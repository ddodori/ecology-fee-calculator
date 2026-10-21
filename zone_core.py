"""2-2 용도지역 교차분석 엔진 (V1.4.9 통합 프로세스).

이 모듈은 교차 결과를 산출하고, 같은 버튼 실행 안에서 zone_adjust의
편입면적 보정으로 연결됩니다. 면적은 EPSG:5179 평면면적(m²)이며
QGIS 타원체 $area 와 차이가 날 수 있고, 법정 부담금은 계산하지 않습니다.
"""
from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import zipfile
import codecs

from shp_export import make_shp_zip, simple_qgis_zone_frame

import geopandas as gpd
import pandas as pd
from shapely.geometry import GeometryCollection, MultiPolygon, Polygon

WORK_CRS = "EPSG:5179"
SHAPEFILE_SUFFIXES = {".shp", ".shx", ".dbf", ".prj", ".cpg", ".qix", ".sbn", ".sbx"}
MAX_UNPACKED_SIZE = 900 * 1024 * 1024


@dataclass
class ZoneResult:
    excel_bytes: bytes
    shp_zip_bytes: bytes | None
    row_count: int
    parcel_count: int
    source_parcels: int
    layer_count: int
    unknown_codes: int
    unmatched_parcels: int
    duplicate_parcels: int
    status_counts: dict[str, int]
    preview: pd.DataFrame
    zone_summary: pd.DataFrame
    detail_full: pd.DataFrame
    validation_full: pd.DataFrame
    geometry_full: gpd.GeoDataFrame
    original_attribute_map: dict[str, str]


def _polygon_only(geom):
    if geom is None or geom.is_empty:
        return None
    if not geom.is_valid:
        geom = geom.buffer(0)
    if isinstance(geom, (Polygon, MultiPolygon)):
        return geom
    if isinstance(geom, GeometryCollection):
        parts = [g for g in geom.geoms if isinstance(g, (Polygon, MultiPolygon)) and not g.is_empty]
        if parts:
            from shapely.ops import unary_union
            return unary_union(parts)
    return None



def _canonical_encoding(value: str | None) -> str | None:
    """Translate common SHP .cpg names/codes into Python/GDAL encoding names."""
    if value is None:
        return None
    value = str(value).strip().replace("\ufeff", "")
    aliases = {
        "949": "CP949", "WINDOWS-949": "CP949", "MS949": "CP949",
        "UHC": "CP949", "EUC-KR": "CP949", "CP949": "CP949",
        "UTF-8": "UTF-8", "UTF8": "UTF-8", "65001": "UTF-8",
        "ANSI 949": "CP949", "ANSI-949": "CP949",
    }
    norm = aliases.get(value.upper(), value)
    try:
        codecs.lookup(norm)
        return norm
    except LookupError:
        return None


def _dbf_encoding_hint(raw: bytes) -> str | None:
    """Infer Korean DBF encoding from raw field labels and character field values.

    When an SHP has no .cpg, GDAL may decode it as Latin-1, yielding
    broken Korean field names in Streamlit. UTF-8 is checked first;
    remaining Korean DBFs conventionally use Windows CP949.
    """
    if len(raw) < 33:
        return None
    header_size = int.from_bytes(raw[8:10], 'little')
    row_size = int.from_bytes(raw[10:12], 'little')
    if not 33 <= header_size <= len(raw) or row_size <= 1:
        return None
    pieces: list[bytes] = []
    fields: list[tuple[int,int,str]] = []
    start=1
    for ix in range(32, header_size-1, 32):
        desc = raw[ix:ix+32]
        if len(desc)<32 or desc[0]==0x0d:
            break
        label=desc[:11].split(b'\x00',1)[0].strip()
        if any(b >= 128 for b in label):
            pieces.append(label)
        ftype=chr(desc[11])
        length=int(desc[16])
        if ftype == 'C' and length>0:
            fields.append((start,length,ftype))
        start += length
    # Include string contents if labels are ASCII (frequent in VWorld layers).
    for n in range(min(25, (len(raw)-header_size)//row_size)):
        row=raw[header_size+n*row_size:header_size+(n+1)*row_size]
        if len(row)!=row_size or row[0:1]==b'*':
            continue
        for offset,size,_ in fields:
            val=row[offset:offset+size].strip().strip(b'\x00')
            if any(b >= 128 for b in val):
                pieces.append(val)
    if not pieces:
        return None
    try:
        for part in pieces:
            part.decode('utf-8', errors='strict')
        return 'UTF-8'
    except UnicodeDecodeError:
        return 'CP949'


def get_shp_encoding(content: bytes, filename: str, preferred: str = 'auto') -> str | None:
    """Return the effective SHP DBF encoding (None for GeoPackage)."""
    if Path(filename).suffix.lower() != '.zip':
        return None
    if preferred.lower() != 'auto':
        encoding=_canonical_encoding(preferred)
        if not encoding:
            raise ValueError(f'지원하지 않는 인코딩입니다: {preferred}')
        return encoding
    with zipfile.ZipFile(BytesIO(content)) as archive:
        shp = [m for m in archive.infolist() if Path(m.filename).suffix.lower()=='.shp' and not m.is_dir()]
        if len(shp)!=1:
            return 'CP949'
        prefix=Path(shp[0].filename).stem.casefold()
        matched={Path(m.filename).suffix.lower():m for m in archive.infolist() if Path(m.filename).stem.casefold()==prefix}
        cpg=matched.get('.cpg')
        cpg_encoding=None
        if cpg:
            try:
                label=archive.read(cpg).decode('utf-8-sig', errors='replace').strip().splitlines()[0]
                cpg_encoding=_canonical_encoding(label)
            except (IndexError, UnicodeError):
                pass
        dbf=matched.get('.dbf')
        # Peek into a bounded amount of the DBF to identify actual byte encoding.
        raw=b''
        if dbf:
            with archive.open(dbf) as src:
                raw=src.read(min(dbf.file_size, 100_000))
        detected=_dbf_encoding_hint(raw)
        # If .cpg is inconsistent with the actual bytes, trust valid UTF-8 first;
        # otherwise keep documented .cpg. Users can manually override both.
        if detected and cpg_encoding and detected != cpg_encoding:
            return detected
        return cpg_encoding or detected or 'CP949'

def read_spatial(content: bytes, filename: str, encoding: str = "auto") -> gpd.GeoDataFrame:
    """Read a single polygon layer from ZIP(SHP) or GeoPackage; safely handles ZIP members."""
    suffix = Path(filename).suffix.lower()
    if suffix not in {".zip", ".gpkg"}:
        raise ValueError(f"{filename}: ZIP(SHP) 또는 GPKG 파일만 지원합니다.")
    with TemporaryDirectory(prefix="eco_gis_") as tmp:
        tmpdir = Path(tmp)
        if suffix == ".gpkg":
            path = tmpdir / "upload.gpkg"
            path.write_bytes(content)
        else:
            with zipfile.ZipFile(BytesIO(content)) as archive:
                members = [m for m in archive.infolist() if not m.is_dir()]
                if sum(m.file_size for m in members) > MAX_UNPACKED_SIZE:
                    raise ValueError(f"{filename}: 압축 해제 총 크기가 허용값을 초과합니다.")
                shp_members = [m for m in members if Path(m.filename).suffix.lower() == ".shp"]
                if len(shp_members) != 1:
                    raise ValueError(f"{filename}: ZIP마다 SHP(.shp) 한 개만 넣어주세요. 발견: {len(shp_members)}개")
                shp_basename = Path(shp_members[0].filename).stem.casefold()
                seen = set()
                for item in members:
                    orig = Path(item.filename)
                    ext = orig.suffix.lower()
                    if ext not in SHAPEFILE_SUFFIXES or orig.stem.casefold() != shp_basename:
                        continue
                    # ZIP 안의 경로를 무시하여 경로 이탈 및 덮어쓰기 방지
                    key = ext.casefold()
                    if key in seen:
                        raise ValueError(f"{filename}: 같은 확장자의 파일이 중복되어 있습니다.")
                    seen.add(key)
                    output = tmpdir / ("layer" + ext)
                    with archive.open(item) as src, output.open("wb") as dst:
                        import shutil
                        shutil.copyfileobj(src, dst)
                if not {".shp", ".shx", ".dbf"}.issubset(seen):
                    raise ValueError(f"{filename}: SHP, SHX, DBF가 모두 필요합니다.")
                if ".prj" not in seen:
                    raise ValueError(f"{filename}: 좌표계 파일(PRJ)이 없습니다. QGIS에서 PRJ를 포함해 다시 압축하세요.")
                path = tmpdir / "layer.shp"
        try:
            effective_encoding = get_shp_encoding(content, filename, encoding)
            frame = gpd.read_file(path, encoding=effective_encoding) if effective_encoding else gpd.read_file(path)
        except Exception as e:
            raise ValueError(f"{filename}: 공간자료를 열 수 없습니다. {e}") from e
        if frame.crs is None:
            raise ValueError(f"{filename}: 좌표계(CRS)가 정의되어 있지 않습니다. QGIS에서 먼저 지정하세요.")
        if frame.empty:
            raise ValueError(f"{filename}: 객체가 없습니다.")
        frame = frame.copy()
        source_count = len(frame)
        frame.geometry = frame.geometry.map(_polygon_only)
        valid_mask = frame.geometry.notna() & ~frame.geometry.is_empty
        if not valid_mask.all():
            invalid_count = int((~valid_mask).sum())
            raise ValueError(
                f"{filename}: {source_count:,}개 도형 중 {invalid_count:,}개가 "
                "폴리곤이 아니거나 사용할 수 없는 도형입니다. QGIS의 '도형 수정' 후 다시 업로드하세요."
            )
        return frame


def list_fields(content: bytes, filename: str, encoding: str = "auto") -> list[str]:
    frame = read_spatial(content, filename, encoding=encoding)
    return [str(c) for c in frame.columns if c != frame.geometry.name]


def _string_pnu(value) -> str:
    if pd.isna(value):
        return ""
    val = str(value).strip()
    if val.endswith(".0") and val[:-2].isdigit():
        val = val[:-2]
    return val


def _parse_area(value):
    if pd.isna(value) or str(value).strip() == "":
        return None
    try:
        number = float(str(value).replace(",", "").strip())
        return number if number >= 0 else None
    except (ValueError, TypeError):
        return None


def analyze_zone_overlay(
    parcel_file: tuple[str, bytes],
    zone_files: list[tuple[str, bytes]],
    pnu_field: str,
    mnum_field: str,
    code_map: dict[str, str],
    landuse_field: str | None = None,
    included_area_field: str | None = None,
    registry_area_field: str | None = None,
    boundary_file: tuple[str, bytes] | None = None,
    include_shp: bool = True,
    parcel_encoding: str = "auto",
    zone_encoding: str = "auto",
) -> ZoneResult:
    if not zone_files:
        raise ValueError("용도지역 SHP를 한 개 이상 업로드하세요.")
    land = read_spatial(parcel_file[1], parcel_file[0], encoding=parcel_encoding)
    if pnu_field not in land:
        raise ValueError(f"지적도에 PNU 필드 '{pnu_field}'가 없습니다.")
    if landuse_field and landuse_field not in land:
        raise ValueError(f"지적도에 지목 필드 '{landuse_field}'가 없습니다.")
    if included_area_field and included_area_field not in land:
        raise ValueError(f"지적도에 편입면적 필드 '{included_area_field}'가 없습니다.")
    if registry_area_field and registry_area_field not in land:
        raise ValueError(f"지적도에 대장면적 필드 '{registry_area_field}'가 없습니다.")

    source_parcels = len(land)
    # Preserve every original parcel field for later Excel outputs, including
    # 법정동, 지번, 건축물 여부 and other user-provided attributes.
    original_columns = [c for c in land.columns if c != land.geometry.name]
    source_id = "__source_row_eco147"
    if source_id in land.columns:
        raise ValueError(f"지적도 속성에 내부 예약 필드 {source_id}가 존재합니다.")
    land[source_id] = range(len(land))
    original_attrs = pd.DataFrame(land[[source_id] + original_columns].copy())
    # Only the four standardized fields are renamed; all other originals remain.
    replacements = {pnu_field: "PNU"}
    if landuse_field:
        replacements[landuse_field] = "지목"
    if included_area_field:
        replacements[included_area_field] = "원본편입면적"
    if registry_area_field:
        replacements[registry_area_field] = "원본대장면적"
    land = land.rename(columns=replacements)
    land["PNU"] = land["PNU"].map(_string_pnu)
    if "지목" not in land:
        land["지목"] = ""
    if "원본편입면적" not in land:
        land["원본편입면적"] = None
    else:
        land["원본편입면적"] = land["원본편입면적"].map(_parse_area)
    if "원본대장면적" not in land:
        land["원본대장면적"] = None
    else:
        land["원본대장면적"] = land["원본대장면적"].map(_parse_area)
    land = land.to_crs(WORK_CRS)
    land = land[land.geometry.area > 0].copy()
    if boundary_file:
        boundary = read_spatial(boundary_file[1], boundary_file[0]).to_crs(WORK_CRS)
        land = gpd.clip(land, boundary, keep_geom_type=True)
        if land.empty:
            raise ValueError("사업구역계와 지적도가 겹치지 않습니다. 좌표계와 위치를 확인하세요.")
    land["분석대상면적_m2"] = land.geometry.area
    land["_parcel_row"] = range(len(land))

    zones = []
    for filename, content in zone_files:
        z = read_spatial(content, filename, encoding=zone_encoding)
        if mnum_field not in z:
            raise ValueError(f"{filename}: 선택한 MNUM 필드 '{mnum_field}'가 없습니다.")
        z = z[[mnum_field, z.geometry.name]].rename(columns={mnum_field: "MNUM"}).copy()
        z["MNUM"] = z["MNUM"].fillna("").astype(str).str.strip()
        zones.append(z.to_crs(WORK_CRS))
    combined = gpd.GeoDataFrame(pd.concat(zones, ignore_index=True), geometry="geometry", crs=WORK_CRS)

    # Prefix all columns to prevent GeoPandas overlay from renaming fields unexpectedly.
    left = land[["PNU", "지목", "원본편입면적", "원본대장면적", "_parcel_row", source_id, "geometry"]].copy()
    right = combined[["MNUM", "geometry"]].copy()
    overlay = gpd.overlay(left, right, how="intersection", keep_geom_type=True, make_valid=True)
    overlay = overlay[overlay.geometry.notna() & ~overlay.geometry.is_empty].copy()
    overlay["교차구적면적_m2"] = overlay.geometry.area
    overlay = overlay.loc[overlay["교차구적면적_m2"] > 1e-7].copy()
    overlay["분석행ID"] = range(1, len(overlay) + 1)
    overlay["용도지역코드"] = overlay["MNUM"].map(lambda x: x[20:26] if len(x) >= 26 else "")
    overlay["용도지역명"] = overlay["용도지역코드"].map(code_map)
    overlay["코드검증"] = overlay.apply(
        lambda r: "MNUM 길이 부족" if not r["용도지역코드"] else (
            "코드표 미등록" if pd.isna(r["용도지역명"]) else "정상"), axis=1
    )
    overlay["용도지역명"] = overlay["용도지역명"].fillna("미확인")

    # Attach the raw source attributes back to every split polygon. Avoid
    # duplicate/ambiguous names when input fields equal derived field names.
    # PNU/지목 are already present as standardized fields; skip those originals
    # only when the source names are identical to the standardized names.
    occupied = set(overlay.columns)
    attr_map = {}
    for field in original_columns:
        if (field == pnu_field and field == "PNU") or (field == landuse_field and field == "지목") or field.startswith("출처파일"):
            continue
        name = field
        if name in occupied or name in attr_map.values():
            name = f"원본_{field}"
            n = 2
            while name in occupied or name in attr_map.values():
                name = f"원본_{field}_{n}"
                n += 1
        attr_map[field] = name
    if attr_map:
        extras = original_attrs[[source_id] + list(attr_map)].rename(columns=attr_map)
        overlay = overlay.merge(extras, on=source_id, how="left", validate="many_to_one")
    retained_fields = list(attr_map.values())
    # Original source name -> overlay/detail attribute name. Keep the exact original
    # ordering and column names for an uncluttered QGIS-only export.
    original_attribute_map = {}
    for field in original_columns:
        if field.startswith("출처파일"):
            continue
        if field in attr_map:
            original_attribute_map[str(field)] = attr_map[field]
        elif field == pnu_field and field == "PNU":
            original_attribute_map[str(field)] = "PNU"
        elif field == landuse_field and field == "지목":
            original_attribute_map[str(field)] = "지목"

    # Check total geometry area of each PNU, comparing spatial coverage (not included-area adjustments).
    valid_land = land.copy()
    valid_land["PNU_검증"] = valid_land["PNU"].map(lambda x: "정상" if len(x) == 19 and x.isdigit() else "PNU 확인 필요")
    land_groups = valid_land.groupby("PNU", dropna=False, as_index=False).agg(
        분석대상면적_m2=("분석대상면적_m2", "sum"),
        원본편입면적=("원본편입면적", "first"),
        원본대장면적=("원본대장면적", "first"),
        필지원본행수=("_parcel_row", "size"),
        PNU_검증=("PNU_검증", "first"),
    )
    overlay_groups = overlay.groupby("PNU", as_index=False)["교차구적면적_m2"].sum().rename(
        columns={"교차구적면적_m2": "용도지역교차합계_m2"}
    )
    validation = land_groups.merge(overlay_groups, on="PNU", how="left")
    validation["용도지역교차합계_m2"] = validation["용도지역교차합계_m2"].fillna(0)
    validation["차이_m2"] = validation["용도지역교차합계_m2"] - validation["분석대상면적_m2"]
    # Area gaps of > max(0.1 m², 0.01% of base area) are examined.
    tolerance = (validation["분석대상면적_m2"] * 0.0001).clip(lower=0.1)
    validation["공간검증"] = "정상"
    validation.loc[validation["차이_m2"] > tolerance, "공간검증"] = "용도지역 중복 의심"
    validation.loc[validation["차이_m2"] < -tolerance, "공간검증"] = "용도지역 누락 의심"
    validation.loc[validation["PNU_검증"] != "정상", "공간검증"] = "PNU 확인 필요"

    detail_columns = (["분석행ID", "PNU"] + retained_fields +
                      ["지목", "원본편입면적", "원본대장면적", "MNUM", "용도지역코드",
                       "용도지역명", "코드검증", "교차구적면적_m2"])
    detail = overlay[detail_columns].copy()
    summary = (
        detail.groupby(["용도지역코드", "용도지역명", "코드검증"], dropna=False, as_index=False)
        .agg(교차건수=("PNU", "size"), 교차면적합계_m2=("교차구적면적_m2", "sum"))
        .sort_values("교차면적합계_m2", ascending=False)
    )
    problems = detail[detail["코드검증"] != "정상"].copy()
    count_unknown = len(problems)
    count_unmatched = int((validation["용도지역교차합계_m2"] == 0).sum())
    count_dup = int((validation["공간검증"] == "용도지역 중복 의심").sum())
    status_counts = {str(k): int(v) for k, v in validation["공간검증"].value_counts().items()}

    instructions = pd.DataFrame([
        ["기능", "2-2 용도지역 지적도 교차분석 (V1.4)"],
        ["법적 확정 여부", "시산/공간분석 전용. 법정 산정 결과가 아닙니다."],
        ["기준 내장 용도지역코드", "2026.10.14 (내장 파일의 기준일)"],
        ["측정 좌표계", f"{WORK_CRS} 투영 평면면적(m²)"],
        ["QGIS $area와 차이", "QGIS 타원체(ellipsoid) 설정에 따라 차이가 발생할 수 있습니다."],
        ["공간검증 비교 기준", "용도지역 교차면적 합계 vs 지적도 SHP 분석대상 도형면적. 조서상 편입면적과의 비교가 아닙니다."],
        ["원본 필드 보존", "지적도에 있던 법정동·지번·건축물 여부 등 전체 속성 필드를 교차 결과에 유지합니다."],
        ["지적도 원본 행수", source_parcels],
        ["분석대상 지적도 도형 행수", len(land)],
        ["용도지역 파일 수", len(zone_files)],
        ["교차분석 결과 행수", len(detail)],
        ["공간누락 의심 필지수", int((validation["공간검증"] == "용도지역 누락 의심").sum())],
        ["용도지역 중복 의심 필지수", count_dup],
        ["MNUM/코드명 검증 필요 행수", count_unknown],
        ["편입면적 보정", "프로그램 2-2 통합 버튼에서 교차분석에 이어 자동 실행합니다."],
    ], columns=["항목", "설명/결과"])

    output_stream = BytesIO()
    with pd.ExcelWriter(output_stream, engine="openpyxl") as writer:
        detail.to_excel(writer, sheet_name="교차분석_상세", index=False)
        validation.to_excel(writer, sheet_name="필지별_공간검증", index=False)
        summary.to_excel(writer, sheet_name="용도지역별_집계", index=False)
        problems.to_excel(writer, sheet_name="코드검토_목록", index=False)
        instructions.to_excel(writer, sheet_name="처리요약", index=False)
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            from openpyxl.styles import Font, PatternFill
            for cell in sheet[1]:
                cell.fill = PatternFill("solid", fgColor="DDEBF2")
                cell.font = Font(bold=True, color="16324A")
            for column in sheet.columns:
                letter = column[0].column_letter
                max_width = max((len(str(c.value)) if c.value is not None else 0) for c in list(column)[:150])
                sheet.column_dimensions[letter].width = min(43, max(13, max_width + 2))
            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    header = str(sheet.cell(1, cell.column).value or "")
                    if "PNU" in header or "코드" in header or "MNUM" in header:
                        cell.number_format = "@"
                    elif ("면적" in header or "차이_m2" in header) and isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                        cell.number_format = "#,##0.00" if "구적" in header or abs(float(cell.value) - round(float(cell.value))) > 1e-9 else "#,##0"
    shp_zip_bytes = None
    if include_shp and not overlay.empty:
        # Even direct engine calls use the same minimal QGIS field layout.
        shp_frame = simple_qgis_zone_frame(overlay, original_attribute_map)
        shp_zip_bytes = make_shp_zip(shp_frame)
    return ZoneResult(
        excel_bytes=output_stream.getvalue(),
        shp_zip_bytes=shp_zip_bytes,
        row_count=len(detail),
        parcel_count=len(validation),
        source_parcels=source_parcels,
        layer_count=len(zone_files),
        unknown_codes=count_unknown,
        unmatched_parcels=count_unmatched,
        duplicate_parcels=count_dup,
        status_counts=status_counts,
        preview=detail.head(100),
        zone_summary=summary.head(50),
        detail_full=detail,
        validation_full=validation,
        geometry_full=gpd.GeoDataFrame(overlay[["분석행ID", "geometry"]].copy(), geometry="geometry", crs=WORK_CRS),
        original_attribute_map=original_attribute_map,
    )
