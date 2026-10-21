"""V1.4.2: 건축물대장 PNU와 토지조서/산정명세서 필지를 비교하는 독립 모듈.

대조 결과는 새 XLSX에 저장합니다. 사용자가 업로드한 원본 파일은 수정하지 않습니다.
"""
from __future__ import annotations
from collections import Counter
from io import BytesIO
from pathlib import Path
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from pnu_core import process, read_codes


def _text(v) -> str:
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _pnu(v) -> str:
    """지적 PNU는 19자리 *문자열*이 기본. 엑셀 숫자 손상은 거부."""
    val = _text(v)
    return val if re.fullmatch(r'\d{19}', val) else ''


def _read_registry(building_bytes: bytes, code_bytes: bytes | None) -> tuple[dict[str, int], str]:
    """기본개요/맞춤형 원본 또는 V1.2 PNU 결과로부터 고유 PNU별 행수를 얻는다."""
    try:
        input_wb = load_workbook(BytesIO(building_bytes), data_only=True, read_only=True)
    except Exception as exc:
        raise ValueError('건축물대장 엑셀을 읽지 못했습니다.') from exc
    try:
        sheet_names = input_wb.sheetnames
        if 'PNU_중복제거' in sheet_names:
            sheet = input_wb['PNU_중복제거']
            records = {}
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                code = _pnu(row[0])
                if code:
                    count = row[5] if len(row) > 5 else 1
                    try:
                        records[code] = max(1, int(count))
                    except (ValueError, TypeError):
                        records[code] = 1
            if not records:
                raise ValueError('PNU_중복제거 시트에 유효한 19자리 PNU가 없습니다.')
            return records, 'V1.2 PNU 생성결과'
        # PNU_생성결과 단독 시트도 지원
        if 'PNU_생성결과' in sheet_names:
            sheet = input_wb['PNU_생성결과']
            headers = [_text(v) for v in next(sheet.iter_rows(max_row=1, values_only=True))]
            if 'PNU' not in headers:
                raise ValueError('PNU_생성결과에서 PNU 필드를 찾지 못했습니다.')
            idx = headers.index('PNU')
            counts = Counter()
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if len(row) > idx:
                    p = _pnu(row[idx])
                    if p:
                        counts[p] += 1
            return dict(counts), '기존 PNU_생성결과'
    finally:
        input_wb.close()

    # 건축HUB 원본은 이미 검증된 V1.2 엔진을 재사용한다.
    converted, _, kind = process(building_bytes, code_bytes)
    return _read_registry(converted, None)[0], kind + ' (자동 PNU 생성)'


def _get_law_codes(code_bytes: bytes | None):
    if code_bytes is None:
        code_bytes = Path(__file__).with_name('법정동코드_전체자료.txt').read_bytes()
    codes, _ = read_codes(code_bytes)
    lookup: dict[str, str] = {}
    ambiguous = set()
    for full, code in codes.items():
        parts = full.split()
        if len(parts) < 3:
            continue
        # (시군, 구, 동리), (시군, 읍면, 리) 등 전체명 끝부분에서 후보를 찾되 중복 후보는 제외
        for n in (3, 4):
            if len(parts) >= n:
                key = ' '.join(parts[-n:])
                if key in lookup and lookup[key] != code:
                    ambiguous.add(key)
                else:
                    lookup[key] = code
    for key in ambiguous:
        lookup.pop(key, None)
    return lookup


def _normalize_lot(value):
    """지번 '산 1-2' / '1-2' / '15' 형식만 엄격하게 허용."""
    lot = re.sub(r'\s+', '', _text(value)).replace('번지', '')
    if not lot:
        return None, '지번 누락'
    kind = '2' if lot.startswith('산') else '1'
    if kind == '2':
        lot = lot[1:]
    m = re.fullmatch(r'(\d+)(?:-(\d+))?', lot)
    if not m:
        return None, '지번 형식 확인'
    bun, ji = int(m.group(1)), int(m.group(2) or 0)
    if not (1 <= bun <= 9999 and 0 <= ji <= 9999):
        return None, '본번/부번 범위 오류'
    return (kind, f'{bun:04d}', f'{ji:04d}'), ''


def _style(wb):
    head_fill = PatternFill('solid', fgColor='17324D')
    for sh in wb.worksheets:
        sh.freeze_panes = 'A2'
        sh.auto_filter.ref = sh.dimensions
        sh.row_dimensions[1].height = 28
        for cell in sh[1]:
            cell.font = Font(color='FFFFFF', bold=True)
            cell.fill = head_fill
            cell.alignment = Alignment(vertical='center', wrap_text=True)
        for col in range(1, sh.max_column + 1):
            title = str(sh.cell(1, col).value or '')
            sh.column_dimensions[get_column_letter(col)].width = (
                27 if title in ('대조상태','검증상태','내용','항목') else
                23 if ('PNU' in title or title=='소재지') else
                19 if title in ('용도지역','법정동명') else 16
            )
        for row in sh.iter_rows(min_row=2):
            for cell in row:
                if 'PNU' in str(sh.cell(1, cell.column).value or ''):
                    cell.number_format = '@'
    # 상태 색상은 서식만, 판정값은 원자료에서만
    for sh in wb.worksheets:
        if sh.title not in ('검증상세','미확인_검토'):
            continue
        headers = [_text(x.value) for x in sh[1]]
        if '대조상태' not in headers:
            continue
        col = headers.index('대조상태') + 1
        for row in sh.iter_rows(min_row=2):
            cell = row[col-1]
            if cell.value == '일치':
                cell.font = Font(color='19704C', bold=True)
            else:
                cell.font = Font(color='A45712')


def match_to_ledger(ledger_bytes: bytes, registry_bytes: bytes,
                    code_bytes: bytes | None = None,
                    sheet_name: str | None = None) -> tuple[bytes, dict, str, str]:
    """return: (결과 XLSX, 통계, 대상 시트, 대장 입력 유형)."""
    if not ledger_bytes or not registry_bytes:
        raise ValueError('토지조서와 건축물대장(PNU 결과) 엑셀 두 파일이 필요합니다.')
    registry, source_kind = _read_registry(registry_bytes, code_bytes)
    try:
        source_wb = load_workbook(BytesIO(ledger_bytes), data_only=True, read_only=True)
    except Exception as exc:
        raise ValueError('토지조서 또는 산정명세서 엑셀을 읽지 못했습니다.') from exc

    try:
        if sheet_name and sheet_name in source_wb.sheetnames:
            src = source_wb[sheet_name]
        elif '토지이용계획(용도지역)' in source_wb.sheetnames:
            src = source_wb['토지이용계획(용도지역)']
        else:
            src = source_wb.active

        # 한글 산정명세서 템플릿(원본 3행 머리글, 4행 합계, 5행부터 필지)을 명시적으로 지원
        template = src.title in ('토지이용계획(용도지역)', '토지이용계획(생태·자연도)')
        if template:
            law = _get_law_codes(code_bytes)
            output_header = ['시군','구','동리','번지','지목','편입면적','용도지역',
                             'PNU(자동)','건축물 여부(자동)','기존 건축물 여부']
            startrow=5
        else:
            headerrow, pcol = None, None
            for r_idx, row in enumerate(src.iter_rows(min_row=1, max_row=15, values_only=True), 1):
                labels = [_text(x).upper().replace(' ','') for x in row]
                if 'PNU' in labels:
                    headerrow, pcol = r_idx, labels.index('PNU')
                    break
            if headerrow is None:
                raise ValueError('토지조서에 PNU 열이 없습니다. 산정명세서 원본은 `토지이용계획(용도지역)` 시트를 선택해 주세요.')
            headers = [_text(x) for x in next(src.iter_rows(min_row=headerrow, max_row=headerrow,values_only=True))]
            output_header = headers + ['건축물 여부(자동)']
            startrow=headerrow+1

        out = Workbook()
        result = out.active
        result.title = '대조결과'
        detail = out.create_sheet('검증상세')
        uncertain = out.create_sheet('미확인_검토')
        summary = out.create_sheet('처리요약')
        result.append(output_header)
        detail_header = ['원본행', 'PNU', '건축물대장 행수', '대조상태', '기존 건축물 여부']
        detail.append(detail_header)
        uncertain.append(detail_header)
        summary.append(['항목','내용'])
        stats = Counter()

        for rowidx, row in enumerate(src.iter_rows(min_row=startrow, values_only=True), startrow):
            vals = list(row)
            if not any(_text(v) for v in vals):
                continue
            if template:
                def get(i):
                    return vals[i] if i < len(vals) else None
                city, gu, dong, lot = map(_text, (get(1),get(2),get(3),get(5)))
                if not city or not dong or not lot or city in ('합계','시.군'):
                    continue
                match_key = ' '.join(x for x in (city,gu,dong) if x)
                code10 = law.get(match_key)
                parts, reason = _normalize_lot(lot)
                if not code10:
                    pnu, reason = '', '법정동코드 미매칭'
                elif parts is None:
                    pnu = ''
                else:
                    pnu = code10 + ''.join(parts)
                original_o = _text(get(19)) if src.title=='토지이용계획(용도지역)' else ''
                output = [city,gu,dong,lot,_text(get(6)),get(8),_text(get(9)),pnu]
            else:
                if len(vals) <= pcol:
                    continue
                raw=vals[pcol]
                if _text(raw)=='' and not any(_text(v) for v in vals):
                    continue
                pnu=_pnu(raw)
                if not pnu:
                    reason = 'PNU 누락/형식 오류 (19자리 문자열 필요)'
                else:
                    reason = ''
                output = vals[:len(headers)]

            stats['전체'] += 1
            if reason:
                mark, count, status='',0,reason
                stats['검토 필요'] += 1
            elif pnu in registry:
                mark,count,status='O',registry[pnu],'일치'
                stats['일치'] += 1
            else:
                mark,count,status='',0,'대장 미조회 (건축물 없음 확정 아님)'
                stats['대장 미조회'] += 1
            if template:
                final = output + [mark, original_o]
                if original_o.upper() == 'O' and mark != 'O':
                    stats['기존 O·신규 미조회'] += 1
            else:
                final = output + [mark]
            result.append(final)
            verification = [rowidx, pnu, count, status, original_o if template else '']
            detail.append(verification)
            if status != '일치':
                uncertain.append(verification)

        summary.append(['대상 시트',src.title])
        summary.append(['건축물대장 입력 형태',source_kind])
        summary.append(['건축물대장 고유 PNU 수',len(registry)])
        summary.append(['대조한 토지조서 행',stats['전체']])
        summary.append(['대장 일치 O',stats['일치']])
        summary.append(['대장 미조회',stats['대장 미조회']])
        summary.append(['코드/PNU 오류로 검토 필요',stats['검토 필요']])
        if template:
            summary.append(['기존 O이나 신규 대장 미조회',stats['기존 O·신규 미조회']])
        summary.append(['주의 1','대장 미조회는 건축물이 없다는 뜻이 아닙니다. 다운로드 지역·기준일·대장종류를 확인해야 합니다.'])
        summary.append(['주의 2','대장에 해당 PNU가 존재해도 훼손 제외지역이 법적으로 확정되는 것은 아닙니다.'])
        summary.append(['결과 시트 안내','대조결과: 건축물 여부 중심 / 검증상세: 원본행·대장 행수·대조상태 / 미확인_검토: 미조회·오류만'])
        summary.append(['주의 3','원본 엑셀은 변경하지 않으며 분석 결과를 별도 엑셀에 저장합니다.'])
        _style(out)
        data=BytesIO()
        out.save(data)
        return data.getvalue(),dict(stats),src.title,source_kind
    finally:
        source_wb.close()
