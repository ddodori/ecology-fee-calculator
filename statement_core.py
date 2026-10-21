"""Final review statement assembled from verified step 2-3 and 2-4 results.

Manual benchmark rates only; must be reviewed for the relevant approval date.
"""
from __future__ import annotations
from decimal import Decimal, ROUND_DOWN
from io import BytesIO
from pathlib import Path
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ZONE_LABELS = ('주거·상업·공업·계획관리지역', '녹지지역', '생산관리지역', '농림지역', '보전관리지역', '자연환경보전지역')
ZONE_FACTORS = (1, 2, 2.5, 3, 3.5, 4)
GRADE_LABELS = ('1등급','2등급','3등급','별도관리지역')
GRADE_FACTORS = (4, 3, 2, 5)


def dec(x):
    if x is None or str(x).strip() == '':return Decimal(0)
    return Decimal(str(x).replace(',',''))


def classify_zone(name):
    s=str(name or '').replace(' ','')
    if any(x in s for x in ('주거','상업','공업','계획관리')):return 0
    if '녹지' in s:return 1
    if '생산관리' in s:return 2
    if '농림' in s:return 3
    if '보전관리' in s:return 4
    if '자연환경보전' in s:return 5
    return None


def build_statement(damage, ecology, *, project_name='', project_location='', owner='', owners=None, authority='', unit_price=300, approval_note='', region_factors=None):
    """Return Excel bytes; fail closed on area mismatches or unclassified land zones."""
    if damage.pending_count:
        raise ValueError(f'2-3 보정 보류 {damage.pending_count}건이 있어 명세서를 만들 수 없습니다.')
    if ecology.overlap_fragments:
        raise ValueError(f'생태자연도 등급 중첩 {ecology.overlap_fragments}건이 있습니다. 원본 경계를 검토해 주세요.')
    remain=dec(damage.remaining_m2)
    if remain != dec(ecology.total_m2):
        raise ValueError(f'2-3 산정대상면적({remain}㎡)과 2-4 등급합계({ecology.total_m2}㎡)가 일치하지 않습니다.')
    excluded=dec(damage.first_m2)+dec(damage.second_m2)+dec(damage.third_m2)
    if excluded != dec(damage.excluded_m2):
        raise ValueError('1·2·3차 제외 합계와 제외 총계가 일치하지 않습니다.')
    totals=[Decimal(0) for _ in ZONE_LABELS]
    unknown=[]
    for _,r in damage.detail.iterrows():
        if str(r.get('처리구분','')) not in ('잔여','잔여·건축물 검토'):continue
        area=dec(r.get('최종보정면적_m2'))
        if area==0:continue
        name=r.get('용도지역명','')
        idx=classify_zone(name)
        if idx is None:unknown.append(f"{r.get('PNU','')} ({name})")
        else:totals[idx]+=area
    if unknown:
        raise ValueError('지역계수 미분류 용도지역이 있습니다: '+', '.join(unknown[:6]))
    if sum(totals)!=remain:
        raise ValueError(f'용도지역별 대상면적 합계 {sum(totals)}㎡와 2-3 산정대상면적 {remain}㎡ 불일치')
    eco=[dec(ecology.class_totals[g]) for g in GRADE_LABELS]
    if sum(eco)!=remain:
        raise ValueError('생태자연도 등급별 면적 합계가 산정대상면적과 일치하지 않습니다.')
    if region_factors is None:
        region_factors = ZONE_FACTORS
    if len(region_factors) != len(ZONE_LABELS):
        raise ValueError('지역계수는 6개 용도지역에 모두 입력해야 합니다.')
    region_factors = tuple(dec(f) for f in region_factors)
    if any(f < 0 for f in region_factors):
        raise ValueError('지역계수는 음수일 수 없습니다.')
    price=dec(unit_price)
    if price<=0:raise ValueError('단위면적당 부과금액은 0보다 커야 합니다.')

    book=load_workbook(BytesIO(damage.excel_bytes))
    original=book['산정명세서']
    book.remove(original)
    sheet=book.create_sheet('산정명세서',0)
    eco_sheet=book.create_sheet('토지이용계획(생태·자연도)')
    sheet.merge_cells('A1:H1')
    sheet['A1']='생태계보전부담금 산정명세서 (검토용)'
    sheet.merge_cells('A2:H2')
    sheet['A2']='수작업 명세서의 계수·산식을 참고한 검증안 · 적용 시점의 법령 및 단가는 별도 확인 필요'
    for row,label,value in [(4,'사업명',project_name),(5,'사업위치',project_location),(6,'사업자',owner),(7,'적용 기준/승인일',approval_note)]:
        sheet.cell(row,1,label);sheet.cell(row,3,value)
    sheet['A9']='면적 구분';sheet['C9']='면적(㎡)';sheet['E9']='비고'
    for r,label,value,desc in [
        (10,'전체 편입면적',remain+excluded,'1·2·3차 제외 + 산정대상면적'),
        (11,'1차 제외면적 · 지목별',dec(damage.first_m2),'지목: 대·장·학·도·철·체·원 해당 토지'),
        (12,'2차 제외면적 · 용도지역별',dec(damage.second_m2),'주거·상업·공업·계획관리지역 중 전·답·임·염·천·유·공 외 지목'),
        (13,'3차 제외면적 · 시설물',dec(damage.third_m2),'1·2차 미해당 토지 중 시설물 설치 확인 구간(건축물 O 참고)'),
        (14,'훼손 제외면적 합계',None,'1·2·3차 제외면적 합계(중복 차감 없음)'),
        (15,'생태계보전부담금 산정대상면적',None,'전체 편입면적에서 훼손 제외면적을 차감한 면적')]:
        sheet.cell(r,1,label);sheet.cell(r,5,desc)
        if value is not None:sheet.cell(r,3,float(value))
    sheet['C14']='=SUM(C11:C13)'
    sheet['C15']='=C10-C14'
    sheet['A17']='단위면적당 부과금액(원/㎡)';sheet['C17']=float(price);sheet['E17']='「자연환경보전법 시행령」 제38조제1항: 300원/㎡ (적용 시점 확인)'
    sheet['A19']='지역계수별 산정';sheet['B20']='용도지역';sheet['C20']='면적(㎡)';sheet['D20']='지역계수';sheet['E20']='산정액(원)'
    for i,(name,factor,area) in enumerate(zip(ZONE_LABELS,region_factors,totals),21):
        sheet.cell(i,2,name);sheet.cell(i,3,float(area));sheet.cell(i,4,factor)
        sheet.cell(i,5,f'=C{i}*$C$17*D{i}/2')
    sheet['B27']='지역계수 소계';sheet['C27']='=SUM(C21:C26)';sheet['E27']='=SUM(E21:E26)'
    sheet['A29']='생태자연도계수별 산정';sheet['B30']='생태등급';sheet['C30']='면적(㎡)';sheet['D30']='등급계수';sheet['E30']='산정액(원)'
    for i,(name,factor,area) in enumerate(zip(GRADE_LABELS,GRADE_FACTORS,eco),31):
        sheet.cell(i,2,name);sheet.cell(i,3,float(area));sheet.cell(i,4,factor)
        sheet.cell(i,5,f'=C{i}*$C$17*D{i}/2')
    sheet['B35']='생태자연도 소계';sheet['C35']='=SUM(C31:C34)';sheet['E35']='=SUM(E31:E34)'
    sheet['A37']='산정금액(원)';sheet['E37']='=E27+E35'
    sheet['A38']='납기 내 부담금(10원 미만 절사)';sheet['E38']='=ROUNDDOWN(E37,-1)'
    sheet['A39']='가산금(검토용 3%, 10원 미만 절사)';sheet['E39']='=ROUNDDOWN(E37*0.03,-1)'
    sheet['A40']='납기 후 합계(검토용)';sheet['E40']='=E38+E39'
    sheet['A42']='면적 검증';sheet['B42']='용도지역 합계 - 산정대상면적';sheet['E42']='=C27-C15'
    sheet['B43']='생태등급 합계 - 산정대상면적';sheet['E43']='=C35-C15'
    sheet['A45']='※ 단위면적당 금액: 「자연환경보전법 시행령」 제38조제1항(300원/㎡). 지역계수: 같은 조 제2항 및 별표 2의5. 계수·가산금·절사 및 적용 시점의 법령은 별도 확인하세요.'
    sheet.merge_cells('A45:H45')
    sheet['A45'].alignment=Alignment(wrap_text=True)
    sheet.row_dimensions[45].height=36

    eco_sheet.append(['연번','PNU','지번','지목','용도지역','산정대상면적(㎡)','1등급(㎡)','2등급(㎡)','3등급(㎡)','별도관리지역(㎡)','등급면적합계(㎡)','검토'])
    for i,(_,r) in enumerate(ecology.details.iterrows(),2):
        vals=[i-1,str(r.get('PNU','')),str(r.get('지번','')),str(r.get('지목','')),str(r.get('용도지역','')),
              r.get('잔여면적(㎡)',0),r.get('1등급(㎡)',0),r.get('2등급(㎡)',0),r.get('3등급(㎡)',0),r.get('별도관리(㎡)',0),None,r.get('검토','')]
        eco_sheet.append(vals)
        eco_sheet.cell(i,11,f'=SUM(G{i}:J{i})')
        eco_sheet.cell(i,2).number_format='@'
    eco_sheet.freeze_panes='F2';eco_sheet.auto_filter.ref=f'A1:L{eco_sheet.max_row}'
    widths=[8,24,16,10,25,22,16,16,16,19,21,26]
    for n,w in enumerate(widths,1):eco_sheet.column_dimensions[get_column_letter(n)].width=w
    for i,w in enumerate([31,31,20,16,24,18,18,15],1):sheet.column_dimensions[get_column_letter(i)].width=w
    for row in (1,9,19,20,29,30):
        for cell in sheet[row][:8]:
            cell.fill=PatternFill('solid',fgColor='244F64');cell.font=Font(bold=True,color='FFFFFF')
    for cell in eco_sheet[1]:
        cell.fill=PatternFill('solid',fgColor='244F64');cell.font=Font(bold=True,color='FFFFFF');cell.alignment=Alignment(wrap_text=True)
    eco_sheet.row_dimensions[1].height=33
    sheet.row_dimensions[1].height=34
    for row in sheet:
        for cell in row:
            if cell.coordinate.startswith(('C','E')) and (isinstance(cell.value,(float,int)) or (isinstance(cell.value,str) and cell.value.startswith('='))):
                cell.number_format='#,##0.00'
    for row in eco_sheet.iter_rows(min_row=2,min_col=6,max_col=11):
        units=[dec(c.value) for c in row[:5]]
        summed=sum(units)
        for cell in row:
            value=summed if cell.column==11 else dec(cell.value)
            if value == value.to_integral_value():
                cell.number_format='#,##0'
            elif value*10 == (value*10).to_integral_value():
                cell.number_format='#,##0.0'
            else:
                cell.number_format='#,##0.00'
    # Final statement: structured, print-friendly calculation sheet using a separate remarks column.
    navy='183A50'; blue='245B74'; ice='EAF2F7'; light='F5F8FA'; ink='223548'; muted='56697A'
    gridline=Side(style='thin',color='DBE4E9')
    sheet.sheet_view.showGridLines=False
    sheet.freeze_panes='C10'
    sheet.sheet_properties.pageSetUpPr.fitToPage=True
    sheet.page_setup.orientation='landscape';sheet.page_setup.paperSize=sheet.PAPERSIZE_A3
    sheet.page_setup.fitToWidth=1;sheet.page_setup.fitToHeight=1
    sheet.print_options.horizontalCentered=True;sheet.print_area='A1:H45'
    for col,width in {'A':32,'B':25,'C':20,'D':16,'E':35,'F':16,'G':16,'H':15}.items():
        sheet.column_dimensions[col].width=width
    sheet['A1']='생태계보전부담금 산정명세서'
    sheet['A2']='산정 대상면적·제외면적·계수별 부과금액 산정 / 검토용'
    sheet.row_dimensions[1].height=44;sheet.row_dimensions[2].height=26
    for r in [4,5,6,7]:
        sheet.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
        sheet.merge_cells(start_row=r,start_column=5,end_row=r,end_column=8)
        sheet.cell(r,1).font=Font(name='맑은 고딕',bold=True,size=10,color=navy)
        sheet.cell(r,3).font=Font(name='맑은 고딕',size=11,color=ink)
        sheet.row_dimensions[r].height=24
    sheet['E4']='사업명·위치·사업자는 사용자가 입력한 값입니다.'
    sheet['E5']='사업지구 위치 및 적용 지역 확인'
    sheet['E6']='산정 의뢰 주체 확인'
    sheet['E7']='부과 기준일 및 적용 법령 확인'
    for r in range(9,16):
        sheet.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
        sheet.merge_cells(start_row=r,start_column=5,end_row=r,end_column=8)
        sheet.row_dimensions[r].height=32 if r in (11,12,13) else 28
    sheet.merge_cells('E17:H17')
    sheet.row_dimensions[17].height=34
    sheet['A9']='구분';sheet['C9']='면적 (㎡)';sheet['E9']='비고 / 제외 적용 기준'
    sheet['A19']='지역계수별 산정';sheet['A29']='생태자연도계수별 산정'
    for r in (19,29):
        sheet.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)
        sheet.row_dimensions[r].height=29
    sheet['B20']='용도지역';sheet['C20']='면적 (㎡)';sheet['D20']='지역계수';sheet['E20']='부과금액 (원)'
    sheet['B30']='생태·자연도 등급';sheet['C30']='면적 (㎡)';sheet['D30']='등급계수';sheet['E30']='부과금액 (원)'
    for r in list(range(10,16))+list(range(20,28))+list(range(30,36))+[17,37,38,39,40,42,43]:
        for c in range(1,9):
            cl=sheet.cell(r,c)
            cl.border=Border(bottom=gridline)
            if r%2==0:cl.fill=PatternFill('solid',fgColor=light)
            if cl.value is not None:
                cl.font=Font(name='맑은 고딕',size=10,color=ink,bold=(r in (14,15,17,27,35,37,38,40)))
                cl.alignment=Alignment(vertical='center',wrap_text=(c in (1,2,5)))
        if r in (14,15,27,35,37,38,40):
            for c in range(1,9):sheet.cell(r,c).fill=PatternFill('solid',fgColor=ice)
    for r in (9,19,20,29,30):
        for c in range(1,9):
            cl=sheet.cell(r,c);cl.fill=PatternFill('solid',fgColor=navy if r in (9,19,29) else blue)
            if cl.__class__.__name__!='MergedCell':cl.font=Font(name='맑은 고딕',bold=True,size=10,color='FFFFFF')
    sheet['A1'].font=Font(name='맑은 고딕',bold=True,size=19,color='FFFFFF')
    sheet['A2'].font=Font(name='맑은 고딕',size=10,color=muted)
    sheet['A1'].fill=PatternFill('solid',fgColor=navy)
    sheet['A1'].alignment=Alignment(vertical='center',horizontal='left',indent=1)
    for r in (10,11,12,13,14,15,17):
        sheet.cell(r,5).font=Font(name='맑은 고딕',size=9,color=muted)
        sheet.cell(r,5).alignment=Alignment(vertical='center',wrap_text=True,indent=1)
    sheet.row_dimensions[45].height=42
    sheet['A45'].font=Font(name='맑은 고딕',size=9,color=muted)
    sheet['A45'].alignment=Alignment(wrap_text=True,vertical='center')
    for r in range(21,27):sheet.cell(r,5).number_format='#,##0'
    for r in range(31,35):sheet.cell(r,5).number_format='#,##0'
    for r in (27,35,37,38,39,40):sheet.cell(r,5).number_format='#,##0'
    for r in (10,11,12,13,14,15,17,21,22,23,24,25,26,27,31,32,33,34,35):
        cl=sheet.cell(r,3)
        if isinstance(cl.value,(int,float)):
            v=dec(cl.value)
            cl.number_format='#,##0' if v==v.to_integral_value() else ('#,##0.0' if v*10==(v*10).to_integral_value() else '#,##0.00')
        else:cl.number_format='#,##0.00'
    sheet['A37']='산정금액 (원)';sheet['A38']='납기 내 납부금액 (원)'
    sheet['A39']='가산금 (검토용)';sheet['A40']='납기 후 납부금액 (검토용)'
    sheet['E38'].font=Font(name='맑은 고딕',bold=True,size=13,color=navy)
    # Landscape tabulation; no stray punctuation for integer values.
    eco_sheet.sheet_view.showGridLines=False
    eco_sheet.sheet_properties.pageSetUpPr.fitToPage=True
    eco_sheet.page_setup.orientation='landscape';eco_sheet.page_setup.fitToWidth=1
    # Rename earlier audit wording without altering underlying quantities or formulas.
    land=book['토지이용계획(용도지역)']
    for row in land.iter_rows(min_row=1,max_row=min(4,land.max_row)):
        for cell in row:
            if isinstance(cell.value,str) and '잔여면적' in cell.value:
                cell.value=cell.value.replace('잔여면적','산정대상면적')

    # Apply the user's supplied manual statement worksheet as the first-sheet template.
    template_file=Path(__file__).with_name('statement_template.xlsx')
    if not template_file.exists():
        raise FileNotFoundError('명세서 원본 양식(statement_template.xlsx)이 없습니다. V1.4.21 전체 파일을 적용해 주세요.')
    source_book=load_workbook(template_file)
    source=source_book.worksheets[0]
    book.remove(sheet)
    sheet=book.create_sheet('산정명세서',0)
    for row in source:
        for cell in row:
            if cell.__class__.__name__=='MergedCell':continue
            dest=sheet[cell.coordinate]
            dest.value=cell.value
            if cell.has_style:
                # Style indexes are only valid within the source workbook.
                # Copy each style component to register it in the destination workbook.
                dest.font = copy(cell.font)
                dest.fill = copy(cell.fill)
                dest.border = copy(cell.border)
                dest.alignment = copy(cell.alignment)
                dest.protection = copy(cell.protection)
                dest.number_format = cell.number_format
    for block in source.merged_cells.ranges:sheet.merge_cells(str(block))
    for col,dimension in source.column_dimensions.items():
        sheet.column_dimensions[col].width=dimension.width
        sheet.column_dimensions[col].hidden=dimension.hidden
    for index,dimension in source.row_dimensions.items():
        if dimension.height:sheet.row_dimensions[index].height=dimension.height
    sheet.sheet_view.showGridLines=False
    sheet.freeze_panes=None
    sheet.sheet_properties.pageSetUpPr.fitToPage=True
    sheet.page_setup.orientation='portrait'
    sheet.page_setup.fitToWidth=1
    sheet.print_area='B1:K52'
    # Populate the exact cells of the provided manual template. No original example names survive.
    sheet['D5']=project_name
    sheet['D6']=project_location
    sheet['B6']='위         치'
    sheet['D16']=approval_note
    sheet['D17']=authority
    sheet['F22']=float(remain+excluded)
    sheet['F24']=float(dec(damage.first_m2))
    sheet['F25']=float(dec(damage.second_m2))
    sheet['F26']=float(dec(damage.third_m2))
    # Original-conservation area is outside the currently implemented workflow.
    # Leave it blank and do not silently assume that it is zero.
    sheet['F27']=None
    sheet['H27']='원형보전면적: 별도 확인·입력 필요 (현재 자동 계산 미지원)'
    sheet['F23']='=SUM(F24:F27)'
    sheet['F28']='=F22-F23'
    sheet['H22']='전체 편입면적(지적도·토지조서 기준)'
    sheet['H23']='「자연환경보전법」 관련 훼손 제외 항목 합계'
    sheet['H24']='대·장·학·도·철·체·원 (지목에 의한 제외)'
    sheet['H25']='주거·상업·공업·계획관리지역의 지목별 제외 조건'
    sheet['H26']='1·2차 미해당 구간 중 시설물이 설치된 토지 (건축물대장 등 증빙 검토)'
    sheet['H28']='전체 편입면적에서 훼손 제외면적을 차감한 산정대상면적'
    # Multi-owner data remain readable and grouped by legal identity field.
    owners=owners if owners is not None else ([{'name':owner,'address':'','business_number':'','corporate_number':''}] if owner else [])
    owner_cols={'name':8,'address':10,'business_number':12,'corporate_number':14}
    for field,rownum in owner_cols.items():
        sheet.cell(rownum,4,'\n'.join(f'  ①' if False else f'{idx+1}. {o.get(field,"")}' for idx,o in enumerate(owners)))
        sheet.cell(rownum,4).alignment=Alignment(vertical='center',wrap_text=True)
        sheet.row_dimensions[rownum].height=max(32,20*len(owners))
    for idx,(area,fac) in enumerate(zip(totals,region_factors),36):
        sheet.cell(idx,5,float(area));sheet.cell(idx,6,float(price));sheet.cell(idx,8,fac)
        sheet.cell(idx,9,f'=E{idx}*F{idx}*H{idx}/2')
    for idx,(area,fac) in enumerate(zip(eco,GRADE_FACTORS),43):
        sheet.cell(idx,5,float(area));sheet.cell(idx,6,float(price));sheet.cell(idx,8,fac)
        sheet.cell(idx,9,f'=E{idx}*F{idx}*H{idx}/2')
    sheet['E35']='=SUM(E36:E41)';sheet['I35']='=SUM(I36:I41)'
    sheet['E42']='=SUM(E43:E46)';sheet['I42']='=SUM(I43:I46)'
    sheet['E34']='=(E35+E42)/2';sheet['I34']='=I35+I42'
    sheet['D48']='=ROUNDDOWN(I34*3/100,-1)'
    sheet['F50']='=ROUNDDOWN(I34,-1)'
    sheet['F52']='=F50+D48'
    sheet['B19']=' 2. 생태계 훼손면적 및 제외면적'
    sheet['B28']='③ 산정대상면적 (①-②)'
    sheet['B30']=' 3. 생태계보전부담금 산정내역'
    sheet['H21']='비고'
    for row in [22,24,25,26,27]:
        val=sheet.cell(row,6).value
        if isinstance(val,(float,int)):
            v=dec(val)
            sheet.cell(row,6).number_format='#,##0' if v==v.to_integral_value() else ('#,##0.0' if v*10==(v*10).to_integral_value() else '#,##0.00')
    for row in [23,28]:sheet.cell(row,6).number_format='#,##0.00'
    for row in list(range(36,42))+list(range(43,47)):
        v=dec(sheet.cell(row,5).value)
        sheet.cell(row,5).number_format='#,##0' if v==v.to_integral_value() else ('#,##0.0' if v*10==(v*10).to_integral_value() else '#,##0.00')
        sheet.cell(row,9).number_format='#,##0'
    sheet['B31']='※ 산정방식: 「자연환경보전법 시행령」 제37조·제38조 및 적용 기준에 따라 검토. 300원/㎡: 제38조제1항. 지역계수: 제38조제2항·별표 2의5.'
    sheet['B31'].alignment=Alignment(wrap_text=True,vertical='center')
    sheet.row_dimensions[31].height=42
    source_book.close()

    book.calculation.fullCalcOnLoad=True;book.calculation.forceFullCalc=True;book.calculation.calcMode='auto'
    stream=BytesIO();book.save(stream)
    return stream.getvalue(), {'전체편입면적':remain+excluded,'산정대상면적':remain,'1차제외':dec(damage.first_m2),
                                '2차제외':dec(damage.second_m2),'3차제외':dec(damage.third_m2),'지역별':totals,'생태별':eco}
