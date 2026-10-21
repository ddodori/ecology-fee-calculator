"""Compact QGIS Shapefile export for zoning overlay (V1.4.13).

Keep *original parcel attributes* in their original order and names whenever
DBF's ten-byte field-name limit allows. Korean names use Windows CP949 (two
bytes per Hangul syllable) instead of UTF-8 (three), retaining many more
original labels without resorting to unreadable COL0000001 aliases.

The generated zoned SHP deliberately excludes adjusted area and QA columns.
Detailed adjusted results, verification, and missing/overlap warnings are
available in the separate final Excel report.
"""
from __future__ import annotations

from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import re
import warnings
import zipfile

import geopandas as gpd
import pandas as pd

DBF_ENCODING = 'CP949'
DBF_MAX_NAME_BYTES = 10


def simple_qgis_zone_frame(
    spatial: gpd.GeoDataFrame, source_columns: dict[str, str]
) -> gpd.GeoDataFrame:
    """Only original cadastral fields and the new zone name, no process columns.

    source_columns maps the *original user attribute name* to its internal
    analysis alias after overlay. The exported name is the original label.
    """
    if not source_columns:
        raise ValueError('지적도 원본 필드 정보를 찾을 수 없습니다.')
    if '용도지역명' not in spatial:
        raise ValueError('교차분석 결과에 용도지역명이 없습니다.')
    fields = {}
    for original_name, internal_name in source_columns.items():
        if internal_name not in spatial.columns:
            raise ValueError(f'원본 속성 필드를 복원할 수 없습니다: {original_name} ({internal_name})')
        fields[original_name] = spatial[internal_name].copy()
    zone_field = '용도지역명' if '용도지역명' not in fields else '분석용도'
    fields[zone_field] = spatial['용도지역명'].copy()
    attrs = pd.DataFrame(fields, index=spatial.index)
    return gpd.GeoDataFrame(attrs, geometry=spatial.geometry.copy(), crs=spatial.crs)


def _encodable_name(name: str) -> bool:
    try:
        return len(name.encode(DBF_ENCODING)) <= DBF_MAX_NAME_BYTES
    except UnicodeEncodeError:
        return False


def _valid_shp_name(name: str) -> bool:
    return (
        bool(name.strip())
        and _encodable_name(name)
        and not name.startswith(('.', ' '))
        and not any(ch in name for ch in ':;,.\\/[](){}*?!@#$%&=+<>"\'\n\r\t')
    )


def _max_bytes_prefix(text: str, limit: int) -> str:
    out = ''
    for ch in text:
        try:
            byte_size = len(ch.encode(DBF_ENCODING))
        except UnicodeEncodeError:
            ch = '_'
            byte_size = 1
        if len(out.encode(DBF_ENCODING)) + byte_size > limit:
            break
        out += ch
    return out


def _friendly_prefix(original: str) -> str:
    # These aliases are only used when the actual name cannot fit DBF.
    equivalents = {
        '필지편입면적': '편입면적',
        '필지대장면적': '대장면적',
        '원본편입면적': '편입면적',
        '원본대장면적': '대장면적',
        '건축물 여부(자동)': '건축여부',
        '건축물 여부': '건축여부',
        '건축물존재여부': '건축여부',
    }
    if original in equivalents:
        return equivalents[original]
    # E.g. "1_필지편입면적" becomes "1_편입면적" if the remainder can fit.
    m = re.match(r'^(\d+_)필지(편입면적|대장면적)$', original)
    if m:
        return f'{m.group(1)}{m.group(2)}'
    return re.sub(r'[^0-9A-Za-z_가-힣]', '', original) or '필드'


def _shp_column_names(original_names: list[str]) -> dict[str, str]:
    """Keep every valid original first, shorten only unavoidable long ones."""
    if len(original_names) != len(set(original_names)):
        raise ValueError('같은 이름의 속성 필드가 중복되어 SHP로 내보낼 수 없습니다.')
    used = set()
    names = {}
    # Reserve all already-valid original names before generating any aliases.
    for original in original_names:
        if _valid_shp_name(original) and original.casefold() not in used:
            names[original] = original
            used.add(original.casefold())

    for original in original_names:
        if original in names:
            continue
        candidate_stem = _friendly_prefix(original)
        candidate = _max_bytes_prefix(candidate_stem, DBF_MAX_NAME_BYTES)
        if not candidate or candidate[0] in '._ ': candidate = '필드'
        if candidate.casefold() in used:
            suffix = 2
            while True:
                tail = f'_{suffix}'
                candidate = _max_bytes_prefix(candidate_stem, DBF_MAX_NAME_BYTES-len(tail)) + tail
                if candidate.casefold() not in used:
                    break
                suffix += 1
        names[original] = candidate
        used.add(candidate.casefold())
    return names


def _sanitize_types(frame: gpd.GeoDataFrame, aliases: dict[str, str]) -> gpd.GeoDataFrame:
    frame = frame.rename(columns=aliases).copy()
    for col in aliases.values():
        series = frame[col]
        if pd.api.types.is_datetime64_any_dtype(series):
            frame[col] = series.dt.strftime('%Y-%m-%d %H:%M:%S')
        elif pd.api.types.is_bool_dtype(series):
            frame[col] = series.astype('Int64')
        elif pd.api.types.is_object_dtype(series):
            frame[col] = series.map(
                lambda value: None if value is None or (
                    isinstance(value, float) and pd.isna(value)
                ) else str(value) if isinstance(value, (list, tuple, dict, set)) else value
            )
    return frame


def make_shp_zip(frame: gpd.GeoDataFrame, base_name: str = "용도지역_교차결과") -> bytes:
    """Write an uncomplicated zipped SHP, including only supplied attributes."""
    if frame.empty:
        raise ValueError('내보낼 교차결과가 없습니다.')
    if frame.crs is None:
        raise ValueError('SHP 내보내기용 좌표계가 없습니다.')
    fields = [str(column) for column in frame.columns if column != frame.geometry.name]
    aliases = _shp_column_names(fields)
    export_frame = _sanitize_types(frame, aliases)
    with TemporaryDirectory(prefix='eco_simple_shp_') as folder:
        shp_path = Path(folder) / (base_name + ".shp")
        export_frame.to_file(
            shp_path, driver='ESRI Shapefile', encoding=DBF_ENCODING,
            engine='pyogrio', index=False,
        )
        # Some GDAL installations omit CPG even when an encoding was given.
        cpg_path = shp_path.with_suffix('.cpg')
        cpg_path.write_text(DBF_ENCODING, encoding='ascii')

        mapping = pd.DataFrame([
            {'원본필드명': source, 'SHP필드명': alias,
             '필드명 변경': '유지' if source == alias else '10바이트 제한으로 축약'}
            for source in fields for alias in [aliases[source]]
        ])
        intro = (
            f'QGIS용 공간분석 결과: {base_name}\n'
            '\n'
            '입력 지적도의 원본 속성 필드와 용도지역 구분만 담았습니다.\n'
            'FINAL_M2(최종보정면적), 보정상태, 공간검증, 보류사유 등은 포함하지 않습니다.\n'
            '해당 상세 분석값은 별도로 다운로드한 최종 결과 엑셀에서 확인하세요.\n'
            '\n'
            '1. .shp/.shx/.dbf/.prj/.cpg를 같은 폴더에 함께 압축 해제하세요.\n'
            f'2. QGIS에서 {base_name}.shp를 불러오세요.\n'
            '3. DBF 한글 인코딩은 CP949(.cpg 포함)입니다. 문자가 깨지면 CP949로 설정하세요.\n'
            '4. SHP는 필드명이 10바이트 이하여야 하므로 원본 이름이 긴 필드만 축약합니다.\n'
            '5. 축약된 필드명은 필드명_대조표.csv에서 확인할 수 있습니다.\n'
            '6. 면적 검증용 SHP에는 0㎡ 보정 교차 피처가 제외됩니다.\n'
            '7. 이 파일은 교차 도형 자료이며, SHP 면적을 최종 보정면적으로 해석하면 안 됩니다.\n'
        ).encode('utf-8-sig')
        mem = BytesIO()
        with zipfile.ZipFile(mem, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
            for suffix in ('.shp', '.shx', '.dbf', '.prj', '.cpg'):
                file = shp_path.with_suffix(suffix)
                if not file.is_file():
                    raise RuntimeError(f'SHP 출력 중 {suffix} 파일이 생성되지 않았습니다.')
                archive.write(file, arcname=file.name)
            archive.writestr('필드명_대조표.csv', mapping.to_csv(index=False).encode('utf-8-sig'))
            archive.writestr('사용안내.txt', intro)
        return mem.getvalue()
