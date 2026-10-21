"""V1.4.17 | 지번 PNU 복원 및 차수별 제외판정 근거를 명시한 2시트 결과.

수기 업무매뉴얼에서 판독한 지목/용도지역 규칙을 '검증용 가정'으로 제공한다.
법령상 확정 규칙이라고 선언하지 않으며, 건축물 존재 O만으로 전체 필지
제외를 단정하지 않는다. 면적은 2-2에서 이미 배분한 최종보정면적을 사용한다.
"""
from __future__ import annotations
from dataclasses import dataclass
from decimal import Decimal
from io import BytesIO
from typing import Iterable
import re
import unicodedata

import geopandas as gpd
import pandas as pd

from shp_export import make_shp_zip
from zone_adjust import _as_decimal

COMPLETED = ('정상 보정', '보정 완료·공간검토필요')
DEFAULT_FIRST = ('대', '장', '학', '도', '철', '체', '원')
DEFAULT_RETAIN = ('전', '답', '임', '염', '천', '유', '공')
ZONE_TOKENS = ('주거', '상업', '공업', '계획관리')


@dataclass
class DamageResult:
    excel_bytes: bytes
    excluded_shp: bytes | None
    remain_shp: bytes | None
    detail: pd.DataFrame
    parcel_validation: pd.DataFrame
    summary: pd.DataFrame
    excluded_m2: Decimal
    remaining_m2: Decimal
    pending_count: int
    building_review_count: int
    first_m2: Decimal
    second_m2: Decimal
    third_m2: Decimal
    spatial_review_count: int
    missing_building_count: int


def _pnu_text(value) -> str:
    if pd.isna(value):
        return ''
    s = str(value).strip()
    if s.endswith('.0') and s[:-2].isdigit():
        s = s[:-2]
    return s


def normalize_landuse(value) -> str:
    s = '' if pd.isna(value) else str(value).strip()
    # 지적공부 지목 약칭/풀어쓴 이름만 정확히 변환; 애매한 자동 추론은 하지 않는다.
    aliases = {
        '대지': '대', '공장용지': '장', '학교용지': '학',
        '도로': '도', '철도용지': '철', '체육용지': '체', '유원지': '원',
        '전': '전', '답': '답', '임야': '임', '염전': '염',
        '하천': '천', '유지': '유', '공원': '공',
    }
    return aliases.get(s, s)


def read_building_matches(excel_bytes: bytes) -> dict[str, bool]:
    """2-1 대조결과 내 'O' PNU만 반환; 빈 칸을 건축물 없음으로 해석하지 않는다."""
    with pd.ExcelFile(BytesIO(excel_bytes)) as excel:
        name = '대조결과' if '대조결과' in excel.sheet_names else excel.sheet_names[0]
        sheet = pd.read_excel(excel, sheet_name=name, dtype=str)
    pnu_cols = [c for c in sheet.columns if str(c).strip().casefold() == 'pnu']
    building_cols = [c for c in sheet.columns if '건축물' in str(c) and '여부' in str(c)]
    if not pnu_cols or not building_cols:
        raise ValueError('건축물대장 대조 엑셀에 PNU 및 건축물 여부 열이 필요합니다. 2-1의 대조결과 시트를 확인하세요.')
    pnu_col, flag_col = pnu_cols[0], building_cols[0]
    result = {}
    for pnu, flag in zip(sheet[pnu_col], sheet[flag_col]):
        key = _pnu_text(pnu)
        if len(key) == 19 and key.isdigit() and _text_match_building(flag):
            result[key] = True
    return result


def _text_match_building(value) -> bool:
    """Accept common affirmative spellings without mistaking digit 0 for letter O."""
    if pd.isna(value):
        return False
    if isinstance(value, bool):
        return value
    text = unicodedata.normalize('NFKC', str(value)).strip().upper()
    text = re.sub(r'\s+', '', text)
    return text in ('O', '○', '◯', '●', 'Y', 'YES', 'TRUE', '1', '있음', '유', '존재')


def find_building_columns(frame: pd.DataFrame) -> list[str]:
    """Detect original parcel-building flag fields, including truncated SHP names."""
    result = []
    for col in frame.columns:
        compact = re.sub(r'[\s_()\[\]-]+', '', unicodedata.normalize('NFKC', str(col))).lower()
        # Only original building-status columns, never a computed validation column.
        if '검증' in compact or '판정' in compact or '사유' in compact:
            continue
        if ('건축' in compact or '건물' in compact) and any(
            token in compact for token in ('여부', '유무', '유무확인', '존재', '확인')
        ):
            result.append(col)
    return result


def _qgis_frame(spatial: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
    """이전 버전 요구대로 QGIS용 열은 알아보기 쉬운 최소 필드만 수록."""
    names = {
        'PNU': spatial['PNU'].map(_pnu_text),
        '지목': spatial['지목'].fillna('').astype(str),
        '용도지역': spatial['용도지역명'].fillna('').astype(str),
        '구분': spatial['처리구분'].astype(str),
        '사유': spatial['판정사유'].astype(str),
        '보정면적': pd.to_numeric(spatial['최종보정면적_m2'], errors='coerce'),
    }
    return gpd.GeoDataFrame(pd.DataFrame(names, index=spatial.index), geometry=spatial.geometry.copy(), crs=spatial.crs)


def _find_original_field(frame: pd.DataFrame, *names: str):
    """Try exact source field name; otherwise common prefixes and DBF contractions."""
    for name in names:
        if name in frame.columns:
            return name
    for col in frame.columns:
        comp = re.sub(r'[\s_()0-9]+', '', str(col)).casefold()
        if any(comp == re.sub(r'[\s_()0-9]+','', n).casefold() for n in names):
            return col
    return None


def _jibeon_from_pnu(pnu) -> str:
    """Restore cadastral lot number when no original 지번 column exists.

    19-digit PNU = legal dong(10) + land kind(1) + main lot(4) + sublot(4).
    Kind 1 is ordinary lot; kind 2 is mountain (산). Invalid codes are blank.
    """
    key = _pnu_text(pnu)
    if len(key) != 19 or not key.isdigit() or key[10] not in ('1', '2'):
        return ''
    main, sub = int(key[11:15]), int(key[15:19])
    if main <= 0:
        return ''
    return ('산' if key[10] == '2' else '') + str(main) + (f'-{sub}' if sub else '')


def _jibeon_value(row: pd.Series, original_field: str | None) -> str:
    if original_field:
        value = row.get(original_field)
        if pd.notna(value):
            text = str(value).strip()
            if text and text.casefold() not in ('nan', 'none', 'null'):
                if text.endswith('.0') and text[:-2].isdigit():
                    return text[:-2]
                return text
    return _jibeon_from_pnu(row.get('PNU'))


def _short_reason(row: pd.Series) -> str:
    stage = str(row.get('제외차수', ''))
    use = normalize_landuse(row.get('지목', ''))
    if stage == '1차':
        return f'1차: 제외 지목 ({use})'
    if stage == '2차':
        return f"2차: {row.get('용도지역명', '')} · 지목 {use}"
    if stage == '3차':
        return '3차: 건축물 O (1·2차 미해당)'
    if str(row.get('건축물여부_검증')) == 'O':
        return '잔여: 건축물 O · 3차 제외 미적용'
    return '잔여: 1·2차 미해당 · 건축물 미조회'


def _make_statement_style_excel(spatial: gpd.GeoDataFrame,
                                validations: pd.DataFrame, *,
                                first_landuses: set[str],
                                retained_landuses: set[str],
                                apply_landuse_rule: bool,
                                apply_zone_rule: bool,
                                exclude_building_parcels: bool) -> bytes:
    """Small reviewer-friendly XLSX modeled on the two user-provided manual sheets.

    2-4 생태자연도/계수 산정 is intentionally NOT filled to avoid implying a
    verified statutory charge or carrying unrelated project numbers forward.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    book = Workbook()
    book.calculation.fullCalcOnLoad = True
    book.calculation.forceFullCalc = True
    book.calculation.calcMode = 'auto'
    stmt = book.active
    stmt.title = '산정명세서'
    land = book.create_sheet('토지이용계획(용도지역)')
    land.merge_cells('A1:O1')
    land['A1'] = '토지이용계획(용도지역) · 훼손 제외면적 검증'
    land.merge_cells('A2:O2')
    first_description = '·'.join(x for x in DEFAULT_FIRST if x in first_landuses) + ''.join('·'+x for x in sorted(first_landuses-set(DEFAULT_FIRST)))
    retain_description = '·'.join(x for x in DEFAULT_RETAIN if x in retained_landuses) + ''.join('·'+x for x in sorted(retained_landuses-set(DEFAULT_RETAIN)))
    first_rule = f'1차 지목({first_description})' if apply_landuse_rule else '1차 미적용'
    second_rule = f'2차 주거·상업·공업·계획관리 / 보존 지목({retain_description}) 외' if apply_zone_rule else '2차 미적용'
    third_rule = '3차 건축물 O(1·2차 미해당)' if exclude_building_parcels else '3차 미적용'
    land['A2'] = '판정 기준 (중복 제외): ' + ' → '.join((first_rule, second_rule, third_rule))
    land['A2'].alignment = Alignment(wrap_text=True, vertical='center')
    land.row_dimensions[2].height = 38
    headers = ['연번','PNU','법정동','지번','지목','대장면적(㎡)',
               '편입면적(㎡)','용도지역','건축물 여부',
               '1차 제외(㎡)','2차 제외(㎡)','3차 제외(㎡)','제외 합계(㎡)','잔여면적(㎡)','제외판정 근거']
    for j, h in enumerate(headers,1):
        land.cell(3,j,h)
    land.cell(4,1,'합계')
    for j in [7,10,11,12,13,14]:
        col=get_column_letter(j)
        land.cell(4,j, f'=SUM({col}5:{col}{max(5,4+int(spatial["처리구분"].isin(("제외 후보","잔여","잔여·건축물 검토")).sum()))})')
    law = _find_original_field(spatial,'법정동','법정동명','동리','동.리','동·리')
    jibeon = _find_original_field(spatial,'지번','번지','필지지번','지번명')
    valid_rows=spatial.loc[spatial['처리구분'].isin(('제외 후보','잔여','잔여·건축물 검토'))]
    for n,(_,r) in enumerate(valid_rows.iterrows(),1):
        amount = float(r['최종보정면적_m2'])
        stage=str(r['제외차수'])
        values=[n, _pnu_text(r['PNU']), r.get(law,'') if law else '',
                _jibeon_value(r, jibeon),
                str(r['지목']),r.get('원본대장면적'),amount,
                str(r['용도지역명']), 'O' if r['건축물여부_검증']=='O' else '',
                amount if stage=='1차' else 0,
                amount if stage=='2차' else 0,
                amount if stage=='3차' else 0,
                None, None, _short_reason(r)]
        rr=4+n
        for j,v in enumerate(values,1):
            if j not in (13,14): land.cell(rr,j,v)
        land.cell(rr,13,f'=SUM(J{rr}:L{rr})')
        land.cell(rr,14,f'=G{rr}-M{rr}')
        land.cell(rr,2).number_format='@'
    land.freeze_panes='F5'
    land.auto_filter.ref=f'A3:O{max(4,4+len(valid_rows))}'
    for j,width in enumerate([9,24,19,16,10,17,17,23,14,17,17,17,18,19,43],1):
        land.column_dimensions[get_column_letter(j)].width=width
    for j in range(1,16):
        for row in range(4,5+len(valid_rows)):
            cell=land.cell(row,j)
            if j in (6,7,10,11,12,13,14):
                cell.number_format='#,##0.00'
                # Optional hashes with a trailing period render badly in some Excel versions.
                # Explicitly set the displayed precision by value for source and computed cells.
                if row > 4 and j in (6,7,10,11,12):
                    try:
                        d=_as_decimal(cell.value)
                        if d is not None:
                            places=max(0,min(2,-d.normalize().as_tuple().exponent))
                            cell.number_format='#,##0' + ('.'+ '0'*places if places else '')
                    except Exception:
                        pass
            if row >= 5 and row%2 == 0:
                cell.fill=PatternFill('solid',fgColor='F5F8FC')
    def _fmt_for_amount(x):
        d=_as_decimal(x) or Decimal(0)
        places=max(0,min(2,-d.normalize().as_tuple().exponent))
        return '#,##0' + ('.'+'0'*places if places else '')
    totals={7:Decimal(0),10:Decimal(0),11:Decimal(0),12:Decimal(0),13:Decimal(0),14:Decimal(0)}
    for row, (_,record) in enumerate(valid_rows.iterrows(),5):
        val=_as_decimal(record['최종보정면적_m2']) or Decimal(0)
        step=record['제외차수']
        excluded=val if step in ('1차','2차','3차') else Decimal(0)
        remain=val-excluded
        land.cell(row,13).number_format=_fmt_for_amount(excluded)
        land.cell(row,14).number_format=_fmt_for_amount(remain)
        for j,amount in [(7,val),(10,val if step=='1차' else 0),
                         (11,val if step=='2차' else 0),(12,val if step=='3차' else 0),
                         (13,excluded),(14,remain)]:
            totals[j]+=Decimal(str(amount))
    for j,x in totals.items():
        land.cell(4,j).number_format=_fmt_for_amount(x)
    stmt.merge_cells('B2:H2')
    stmt['B2']='생태계보전부담금 산정명세서'
    stmt['B4']='1. 일반현황'
    stmt['B5']='사업명'; stmt['D5']='입력 예정'
    stmt['B6']='사업자 및 인허가 정보'; stmt['D6']='입력 예정'
    stmt['B8']='2. 생태계 훼손면적'
    stmt['B10']='구분'; stmt['F10']='면적(㎡)'; stmt['H10']='비고'
    stmt['B11']='① 전체 편입면적'; stmt['F11']="='토지이용계획(용도지역)'!G4"
    stmt['B12']='② 생태계 훼손 제외면적'; stmt['F12']="='토지이용계획(용도지역)'!M4"
    stmt['C13']='1차 제외대상 지목'; stmt['F13']="='토지이용계획(용도지역)'!J4"
    stmt['H13'] = first_rule + ' (검증용)'
    stmt['C14']='2차 용도지역·지목'; stmt['F14']="='토지이용계획(용도지역)'!K4"
    stmt['H14'] = second_rule + ' (검증용)'
    stmt['C15']='3차 시설물 설치 토지'; stmt['F15']="='토지이용계획(용도지역)'!L4"
    stmt['H15']=third_rule + ' · 법적 제외 증빙 확인 필요'
    stmt['B16']='③ 생태계 훼손면적 (①-②)'; stmt['F16']='=F11-F12'
    stmt['B18']='3. 생태계보전부담금 산정내역'
    stmt['B19']='생태·자연도 분석 및 지역계수 산정 기능 개발 후 작성 예정'
    stmt['B21']='주의: 이 자료의 제외면적은 검증용 후보이며, 실제 부과면적 확정 전 기준·증빙 검토 필요'
    stmt['B22']='2-2 공간검토·보정 보류 필지는 별도 확인 필요 (보류 제외값은 합계에 미반영)'
    stmt.merge_cells('B19:H19'); stmt.merge_cells('B21:H21'); stmt.merge_cells('B22:H22')
    stmt.column_dimensions['A'].width=4
    for k,w in {'B':29,'C':25,'D':24,'E':12,'F':18,'G':9,'H':76}.items():
        stmt.column_dimensions[k].width=w
    for r, amount in [(11,totals[7]),(12,totals[13]),(13,totals[10]),
                      (14,totals[11]),(15,totals[12]),(16,totals[14])]:
        stmt[f'F{r}'].number_format=_fmt_for_amount(amount)
    for row in [2,4,8,10,18]:
        for cell in stmt[row][1:8]:
            cell.fill=PatternFill('solid',fgColor='E6EEF8' if row !=2 else 'D6E6FA')
            cell.font=Font(name='맑은 고딕',bold=True,color='1D3557',size=14 if row==2 else 10)
    for rr in (13,14,15):
        stmt[f'H{rr}'].alignment=Alignment(vertical='center', wrap_text=True)
        stmt.row_dimensions[rr].height=30
    for cell in land[3]:
        cell.fill=PatternFill('solid',fgColor='DFEAF6')
        cell.font=Font(name='맑은 고딕',bold=True,color='17375E')
        cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    for cell in land[4]:
        cell.fill=PatternFill('solid',fgColor='E7F0EA')
        cell.font=Font(name='맑은 고딕',bold=True,color='1F4330')
    land.row_dimensions[3].height=32
    stmt.sheet_view.showGridLines=False
    land.sheet_view.showGridLines=False
    out=BytesIO()
    book.save(out)
    return out.getvalue()


def analyze_damage_filter(
    geometry: gpd.GeoDataFrame,
    adjusted_detail: pd.DataFrame,
    parcel_validation: pd.DataFrame,
    *,
    landuse_excluded: Iterable[str] = DEFAULT_FIRST,
    zone_retain_landuse: Iterable[str] = DEFAULT_RETAIN,
    building_field: str | None = None,
    building_excel: bytes | None = None,
    exclude_building_parcels: bool = True,
    apply_landuse_rule: bool = True,
    apply_zone_rule: bool = True,
) -> DamageResult:
    """교차 조각별 판정. 누적 면적은 오직 2-2 보정면적이며 도형의 $area가 아님."""
    required = {'분석행ID', 'PNU', '지목', '용도지역명', '최종보정면적_m2', '보정상태', '원본편입면적'}
    if not required.issubset(adjusted_detail) or '분석행ID' not in geometry:
        raise ValueError('2-2의 교차·보정 실행 결과가 없습니다. 용도지역 분석을 먼저 실행하세요.')
    if building_field and building_field not in adjusted_detail:
        raise ValueError(f'건축물 여부 필드를 찾을 수 없습니다: {building_field}')
    if (adjusted_detail['분석행ID'].duplicated().any() or geometry['분석행ID'].duplicated().any()):
        raise ValueError('분석행ID가 중복되어 지적도와 보정자료를 안전하게 결합할 수 없습니다.')
    merged = geometry[['분석행ID', geometry.geometry.name]].merge(
        adjusted_detail, on='분석행ID', validate='one_to_one', how='inner', sort=False
    )
    if len(merged) != len(adjusted_detail):
        raise ValueError('교차 도형과 보정 내역의 분석행ID가 일치하지 않습니다.')
    spatial = gpd.GeoDataFrame(merged, geometry=geometry.geometry.name, crs=geometry.crs)
    if spatial.crs is None:
        raise ValueError('교차분석 도형의 좌표계를 확인할 수 없습니다.')
    if spatial.geometry.isna().any() or spatial.geometry.is_empty.any():
        raise ValueError('교차분석 도형에 비어 있는 폴리곤이 있습니다.')

    # Building flags may be present on the original cadastral layer even when the
    # user forgot to pick the field in the UI. Detect every plausible source
    # column and carry any O within a PNU across all of its zoning fragments.
    building_o = read_building_matches(building_excel) if building_excel else {}
    building_columns = find_building_columns(spatial)
    if building_field and building_field not in building_columns:
        building_columns.insert(0, building_field)
    building_pnus = set(building_o)
    for field_name in building_columns:
        positive = spatial[field_name].map(_text_match_building)
        building_pnus.update(spatial.loc[positive, 'PNU'].map(_pnu_text))
    building_pnus.discard('')

    first = {normalize_landuse(x) for x in landuse_excluded if str(x).strip()}
    retained = {normalize_landuse(x) for x in zone_retain_landuse if str(x).strip()}
    if apply_zone_rule and not retained:
        raise ValueError('2차 제외 규칙의 보존 지목이 비어 있습니다.')

    statuses, reasons, flags, stages = [], [], [], []
    for _, row in spatial.iterrows():
        status = str(row['보정상태'])
        value = _as_decimal(row['최종보정면적_m2'])
        if status not in COMPLETED or value is None:
            statuses.append('보정 보류')
            reasons.append(str(row.get('보류사유', '2-2 보정자료 확인 필요')))
            flags.append('미확인')
            stages.append('보정 보류')
            continue
        if value <= 0:
            statuses.append('0면적 제외')
            reasons.append('최종보정면적 0㎡')
            flags.append('미확인')
            stages.append('보정 보류')
            continue
        pnu = _pnu_text(row['PNU'])
        # Recognize a matching O in the original SHP *or* the 2-1 Excel.
        # By PNU, its zoning fragments inherit the same positive result.
        has_building = pnu in building_pnus
        flag = 'O' if has_building else '미조회'
        landuse = normalize_landuse(row['지목'])
        zone = str(row['용도지역명'])
        # The handwritten workbook checks 1차 -> 2차 -> 3차 exclusively.
        # An O on an already-excluded landuse (e.g., 대/학) must NEVER be
        # counted a second time in the 3차 시설물설치 total.
        if apply_landuse_rule and landuse in first:
            status, stage, reason = '제외 후보', '1차', '1차: 제외 지목'
        elif apply_zone_rule and any(token in zone for token in ZONE_TOKENS) and landuse not in retained:
            status, stage, reason = '제외 후보', '2차', '2차: 용도지역·지목 조건'
        elif has_building and exclude_building_parcels:
            status, stage, reason = '제외 후보', '3차', '3차: 시설물 설치(건축물 O, 검증용)'
        elif has_building:
            status, stage, reason = '잔여·건축물 검토', '잔여', '건축물 O: 전체 필지 제외 여부 미확정'
        else:
            status, stage, reason = '잔여', '잔여', '1·2차 제외 조건 미해당 / 건축물 미조회(없음 확정 아님)'
        statuses.append(status)
        reasons.append(reason)
        flags.append(flag)
        stages.append(stage)

    spatial['제외차수'] = stages
    spatial['처리구분'] = statuses
    spatial['판정사유'] = reasons
    spatial['건축물여부_검증'] = flags
    area = pd.to_numeric(spatial['최종보정면적_m2'], errors='coerce')
    candidate = spatial['처리구분'].eq('제외 후보')
    remaining = spatial['처리구분'].isin(('잔여', '잔여·건축물 검토'))
    valid = candidate | remaining
    # 0㎡ 조각과 계산 보류 조각은 제외/잔여 SHP에 포함하지 않는다.
    excluded = spatial.loc[candidate].copy()
    keep = spatial.loc[remaining].copy()
    extras = spatial.loc[~valid].copy()

    validations = parcel_validation[['PNU', '원본편입면적', '보정상태', '공간검증']].copy()
    validations['PNU'] = validations['PNU'].map(_pnu_text)
    totals = spatial.loc[valid, ['PNU','최종보정면적_m2','처리구분']].copy()
    totals['PNU'] = totals['PNU'].map(_pnu_text)
    records = []
    for pnu, grp in totals.groupby('PNU', sort=False):
        excluded_sum = sum((_as_decimal(x) or Decimal(0) for x in grp.loc[grp['처리구분'].eq('제외 후보'), '최종보정면적_m2']), Decimal(0))
        remain_sum = sum((_as_decimal(x) or Decimal(0) for x in grp.loc[~grp['처리구분'].eq('제외 후보'), '최종보정면적_m2']), Decimal(0))
        records.append({'PNU': pnu, '제외후보면적_m2': float(excluded_sum), '잔여면적_m2': float(remain_sum),
                        '_total_decimal': excluded_sum + remain_sum})
    amounts = pd.DataFrame(records)
    if amounts.empty:
        amounts = pd.DataFrame(columns=['PNU','제외후보면적_m2','잔여면적_m2','_total_decimal'])
    validations = validations.merge(amounts, on='PNU', how='left', validate='one_to_one')
    checks = []
    for _, row in validations.iterrows():
        src, total = _as_decimal(row['원본편입면적']), row['_total_decimal']
        if str(row['보정상태']) not in COMPLETED:
            checks.append('2-2 보정 보류')
        elif pd.isna(total) or src is None:
            checks.append('면적 검증 불가')
        elif abs(src - total) <= Decimal('0.00000001'):
            checks.append('일치')
        else:
            checks.append('불일치')
    validations['합계검증'] = checks
    failures = validations.loc[validations['합계검증'] == '불일치']
    if not failures.empty:
        raise ArithmeticError(f"필지별 제외+잔여 합계가 편입면적과 다른 필지가 {len(failures)}개 있습니다. 내보내기를 중단합니다.")
    validations['면적차이_m2'] = validations.apply(
        lambda r: 0.0 if r['합계검증'] == '일치' else None, axis=1
    )
    validations = validations.drop(columns=['_total_decimal'])

    # Keep full detail in memory for audit/UI, but export ONLY the two sheets
    # the user works with: 산정명세서 + 토지이용계획(용도지역).
    details = spatial.drop(columns=[spatial.geometry.name]).copy()
    detail_view = details.copy()
    first_total = sum((_as_decimal(v) or Decimal(0) for v in
                       spatial.loc[spatial['제외차수']=='1차','최종보정면적_m2']), Decimal(0))
    second_total = sum((_as_decimal(v) or Decimal(0) for v in
                        spatial.loc[spatial['제외차수']=='2차','최종보정면적_m2']), Decimal(0))
    third_total = sum((_as_decimal(v) or Decimal(0) for v in
                       spatial.loc[spatial['제외차수']=='3차','최종보정면적_m2']), Decimal(0))
    summary = pd.DataFrame([
        {'구분':'1차 제외 지목', '면적_m2':float(first_total)},
        {'구분':'2차 용도지역·지목', '면적_m2':float(second_total)},
        {'구분':'3차 시설물 설치(건축물 O, 중복 제외)', '면적_m2':float(third_total)},
        {'구분':'제외 후보 총계', '면적_m2':float(first_total+second_total+third_total)},
        {'구분':'잔여(부과대상 검토)', '면적_m2':float(area.loc[remaining].sum())},
    ])
    excel_bytes = _make_statement_style_excel(
        spatial, validations, first_landuses=first, retained_landuses=retained,
        apply_landuse_rule=apply_landuse_rule, apply_zone_rule=apply_zone_rule,
        exclude_building_parcels=exclude_building_parcels) 
    excluded_zip = make_shp_zip(_qgis_frame(excluded), base_name='훼손제외_대상') if not excluded.empty else None
    remain_zip = make_shp_zip(_qgis_frame(keep), base_name='부담금_산정대상면적') if not keep.empty else None
    return DamageResult(
        excel_bytes=excel_bytes,
        excluded_shp=excluded_zip, remain_shp=remain_zip, detail=detail_view,
        parcel_validation=validations, summary=summary,
        excluded_m2=sum((_as_decimal(x) or Decimal(0) for x in excluded['최종보정면적_m2']), Decimal(0)),
        remaining_m2=sum((_as_decimal(x) or Decimal(0) for x in keep['최종보정면적_m2']), Decimal(0)),
        pending_count=int((spatial['처리구분']=='보정 보류').sum()),
        building_review_count=int((spatial['처리구분']=='잔여·건축물 검토').sum()),
        first_m2=first_total, second_m2=second_total, third_m2=third_total,
        spatial_review_count=int(validations['공간검증'].isin(('용도지역 누락 의심', '용도지역 중복 의심')).sum()),
        missing_building_count=int((spatial.loc[valid, '건축물여부_검증']=='미조회').sum()),
    )
