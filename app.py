"""환경부담금 산정 자동화 웹사이트 V1.4.18.

V1.3의 PNU 생성/건축물 여부 대조 엔진을 유지하면서,
업무매뉴얼 순서에 따라 기초자료 준비 > 데이터 처리 > 산정명세서 작성으로 UI를 구성합니다.
"""
from __future__ import annotations

import hashlib
import pandas as pd
import streamlit as st

from pnu_core import process
from building_workflow import run_building_workflow
from zone_codes import load_zone_code_map
from zone_core import analyze_zone_overlay, list_fields, get_shp_encoding
from zone_adjust import adjust_zone_areas
from shp_export import make_shp_zip, simple_qgis_zone_frame
from damage_filter import analyze_damage_filter, find_building_columns, _text_match_building, DEFAULT_FIRST, DEFAULT_RETAIN
from ecology_core import analyze_ecology, default_grade_values
from statement_core import build_statement
from zone_core import read_spatial

st.set_page_config(page_title="환경부담금 산정 자동화", page_icon="🌿", layout="wide")

st.title("🌿 환경부담금 산정 자동화")
st.caption("생태계보전부담금 산정 업무 지원 시스템 · V1.4.22")
st.info("사용 가능: **건축물대장 PNU 생성·토지조서 대조 통합**, **용도지역 교차분석·면적 자동 보정**. 훼손 제외면적 분석과 생태자연도 등급별 면적 보정도 사용 가능합니다. 최종 산정명세서 검토용 작성이 가능합니다.")

# 산정 기준 용도지역코드는 프로그램 내부에 포함되어 있으며 업로드를 요구하지 않습니다.
try:
    embedded_zone_count = len(load_zone_code_map())
    embedded_zone_status = f"기본 내장 완료 ({embedded_zone_count:,}개 코드)"
except (OSError, ValueError) as exc:
    embedded_zone_count = 0
    embedded_zone_status = "내장 자료 확인 필요"
    st.warning(f"용도지역코드 내장 자료를 읽지 못했습니다: {exc}")

# 전체 프로세스 안내: 버튼처럼 보이지만 작동하지 않는 컨트롤 대신 단계형 정보 카드로 표현
col_a, col_b, col_c = st.columns(3)
with col_a:
    with st.container(border=True):
        st.markdown("#### ① 기초자료 준비")
        st.caption("건축물대장 · 토지조서 · 용도지역 · 생태자연도")
        st.markdown("**자료 안내 제공 / 일부 준비**")
with col_b:
    with st.container(border=True):
        st.markdown("#### ② 데이터 처리")
        st.caption("건축물대장·토지조서 대조 → 용도지역 → 훼손 제외 → 생태자연도")
        st.markdown("**분석 기능 사용 가능**")
with col_c:
    with st.container(border=True):
        st.markdown("#### ③ 산정명세서 작성")
        st.caption("면적·계수 집계 → 명세서 엑셀 출력")
        st.markdown("**산정명세서 작성 및 엑셀 출력 가능**")

st.divider()
preparation, processing, statement = st.tabs(
    ["① 기초자료 준비", "② 데이터 처리", "③ 산정명세서 작성"]
)

with preparation:
    st.subheader("기초자료 준비")
    st.write("생태계보전부담금 산정 매뉴얼에 따른 입력자료와 다운로드 경로입니다.")
    st.caption("최종 산정명세서는 입력자료에 포함하지 않으며, ③ 산정명세서 작성 단계에서 생성할 예정입니다.")

    st.markdown("##### 1. 엑셀 입력자료")
    st.table([
        {"필요 자료": "건축물대장 (기본개요/맞춤형)", "사용 용도": "PNU 생성 및 토지조서 건축물 여부 대조", "개발 상태": "사용 가능"},
        {"필요 자료": "토지조서 엑셀", "사용 용도": "건축물대장 PNU와 대조하여 건축물 여부 확인", "개발 상태": "일부 사용 가능"},
    ])

    st.markdown("##### 2. 기본 내장 기준자료")
    st.success("**법정동코드와 용도지역코드는 프로그램에 기본 내장되어 있습니다. 별도로 업로드하지 않아도 됩니다.**")
    st.table([
        {
            "기본 내장 자료": "법정동코드 (TXT)",
            "사용 용도": "맞춤형 건축물대장 법정동코드 매칭",
            "내장 파일 기준일": "2026.10.14",
        },
        {
            "기본 내장 자료": "용도지역코드 (XLSX)",
            "사용 용도": "MNUM 코드와 용도지역명 자동 매칭 (사용 가능)",
            "내장 파일 기준일": "2026.10.14",
        },
    ])
    st.caption(
        f"용도지역코드: {embedded_zone_status} · 기준일은 이 프로그램에 등록한 파일의 기준일이며, "
        "공식 고시일·자료의 법적 효력 발생일을 의미하지 않습니다."
    )
    with st.expander("ℹ️ 기본 내장 자료의 출처 및 갱신 방법"):
        st.markdown("""
- **법정동코드 (기준일: 2026.10.14):** [행정표준코드관리시스템](https://www.code.go.kr/stdcode/regCodeL.do)의 전체자료 TXT를 내장했습니다. 맞춤형 건축물대장은 최신 TXT를 별도로 업로드해 교체할 수도 있습니다.
- **용도지역코드 (기준일: 2026.10.14):** 제공받은 용도지역코드 XLSX(`UCODE`, `UNAME`)를 내장했습니다. 용도지역 교차분석에서 코드 자동 변환 기능을 사용할 수 있습니다.

※ 행정구역 및 용도지역 코드가 변경되었다면 실제 업무 적용 전 최신 자료와 비교해야 합니다.
""")

    st.markdown("##### 3. GIS 공간자료")
    st.table([
        {"필요 자료": "연속지적도 + 토지대장 (SHP/GPKG)",
         "준비할 자료": "지적경계, PNU, 지목, 편입면적 또는 별도 토지조서", "분석 기능": "2-2 통합분석 사용 가능"},
        {"필요 자료": "용도지역 (SHP/GPKG)",
         "준비할 자료": "브이월드 (연속주제) 국토계획: 도시지역·관리지역·농림지역·자연환경보전지역", "분석 기능": "2-2 통합분석 사용 가능"},
        {"필요 자료": "생태자연도 (SHP/GPKG)",
         "준비할 자료": "에코뱅크의 사업지구 포함 도엽 또는 행정구역 SHP", "분석 기능": "2-4 생태자연도 분석 사용 가능"},
    ])

    with st.expander("📥 건축물대장 다운로드 안내", expanded=True):
        st.markdown("""
**권장 자료: 건축HUB의 건축물대장 ‘기본개요’**

1. [건축HUB 홈페이지](https://www.hub.go.kr/) 접속
2. **원하는대로 건축데이터 → 유형별 건축데이터 제공 → 건축물대장 → 기본개요** 선택
3. 조회할 지역을 선택하고 엑셀 다운로드
4. **② 데이터 처리 → 2-1 건축물대장 정리·토지조서 대조**에서 엑셀 업로드

[건축물대장 기본개요 바로가기](https://www.hub.go.kr/portal/opn/tyb/idx-bdrg.do)

‘기본개요’에는 보통 **시군구코드·법정동코드·대지구분코드·번·지**가 있어 법정동코드 TXT를 별도로 업로드할 필요가 없습니다. 맞춤형 건축정보 파일도 지원합니다.
""")

    with st.expander("📥 용도지역 SHP 다운로드 안내", expanded=True):
        st.markdown("""
**자료 출처: [브이월드(V-World)](https://www.vworld.kr/)**

**다운로드 경로**

브이월드 → **오픈마켓(공간정보 다운로드)** → **연속주제** 검색 → **국토계획** 관련 자료 선택

**필요한 공간자료 4종** (매뉴얼 기준)

1. **(연속주제) 국토계획/도시지역** — 주거·상업·공업·녹지지역 등 포함
2. **(연속주제) 국토계획/관리지역** — 계획관리·생산관리·보전관리지역 등 포함
3. **(연속주제) 국토계획/농림지역**
4. **(연속주제) 국토계획/자연환경보전지역**

사업대상 시·도에 해당하는 **SHP 파일**을 내려받으세요. 대상지에 해당하지 않는 지역 자료는 생략할 수 있습니다.

**주의:** ‘도시지역’과 ‘도시지역경계’는 다른 데이터일 수 있으므로, 용도지역별 분류 속성(`MNUM` 등)이 포함된 자료인지 확인하세요. 여러 ZIP/GPKG 파일을 업로드하면 프로그램이 용도지역 도형을 합쳐 교차분석합니다. **2-2에서 한 번의 실행으로 교차분석과 필지별 정밀도에 맞춘 면적 보정까지 완료합니다.**
""")

    with st.expander("📥 생태자연도 SHP 다운로드 안내"):
        st.markdown("""
**자료 출처: [에코뱅크(EcoBank)](https://www.nie-ecobank.kr/cmmn/Index.do)**

**다운로드 경로**

에코뱅크 → **데이터개방 → 공개데이터 → 국가중점개방데이터 다운로드** → 생태자연도 선택 → 사업지구가 포함된 **도엽 또는 행정구역별 SHP** 다운로드

**[생태자연도 SHP 다운로드 페이지 바로가기](https://www.nie-ecobank.kr/opn/file/list.do?svcId=101)**

※ 실제 산정에 적용할 고시 시점과 데이터 범위가 맞는지 확인하세요. 생태자연도 교차·보정 기능은 **2-4**에서 사용할 수 있습니다.
""")

    st.caption("※ 외부 자료 다운로드를 프로그램이 자동 수행하는 기능은 아직 (미개발)입니다.")


# 데이터 처리 상단 하위 메뉴 (기존 2-1과 2-3을 통합)
PROCESS_CATEGORIES = [
    {
        "id": "building",
        "name": "01. 건축물대장 정리·토지조서 대조",
        "description": "건축물대장 PNU 자동 생성 → 토지조서와 PNU 비교 → 건축물 여부 O 자동 입력",
        "steps": ["building_workflow"],
    },
    {
        "id": "zone",
        "name": "02. 용도지역 분석",
        "description": "SHP 업로드 한 번으로 교차분석부터 편입면적 비례 보정까지",
        "steps": ["zone_workflow"],
    },
    {
        "id": "damage",
        "name": "03. 훼손 제외면적 분석",
        "description": "검증용 제외 후보 및 잔여 구간을 분류하고 엑셀·SHP로 내보내기",
        "steps": ["damage_filter"],
    },
    {
        "id": "ecology",
        "name": "04. 생태자연도 분석",
        "description": "등급별 면적 산출 및 보정",
        "steps": ["ecology_workflow"],
    },
]

PROCESS_STEPS = {
    "building_workflow": {
        "no": 1,
        "name": "건축물대장 PNU 생성 및 필지별 건축물 여부 대조",
        "developed": True,
        "description": "건축물대장 원본에서 PNU를 생성한 다음 토지조서의 필지와 자동 대조합니다.",
    },
    "zone_workflow": {"no": 2, "name": "용도지역 교차분석·최종 면적 자동 보정", "developed": True,
                      "description": "지적도와 용도지역 SHP를 업로드한 뒤 한 번에 교차분석·코드변환·면적비례보정을 완료합니다."},
    "damage_filter": {"no": 4, "name": "생태계 훼손 제외면적 검증", "developed": True,
                      "description": "용도지역 최종보정면적에 검증용 지목·용도지역·건축물 규칙을 적용해 면적표와 SHP를 만듭니다."},
    "ecology_workflow": {"no": 5, "name": "생태자연도 등급별 교차분석·면적 자동 보정", "developed": True,
                        "description": "산정대상면적 SHP와 생태자연도 SHP를 교차해 1·2·3등급·별도관리 면적을 한 번에 산출합니다."},
}

STEP_ORDER = [step_id for category in PROCESS_CATEGORIES for step_id in category["steps"]]
STEP_TO_CATEGORY = {
    step_id: category["id"] for category in PROCESS_CATEGORIES for step_id in category["steps"]
}


def select_category_step(widget_key: str, step_id: str) -> None:
    """현재 카테고리의 세부 단계 선택을 바꿉니다."""
    if step_id in PROCESS_STEPS:
        st.session_state[widget_key] = step_id


def render_unfinished_step(step_id: str) -> None:
    """기능이 개발되지 않았을 때 오해를 막기 위한 명시적 안내."""
    st.info("**(미개발)** 현재 안내만 제공하며, 파일 분석이나 계산은 수행하지 않습니다.")
    descriptions = {
                "damage_filter": "구현 예정: 지목 대·장·학·도·철·체·원 등의 제외 규칙과 용도지역 조건, 건축물 존재 여부를 단계별로 검토.",
        "ecology_workflow": "산정대상면적 SHP의 보정면적을 기준으로 생태자연도 등급 면적을 산출합니다.",
    }
    st.write(descriptions[step_id])
    st.caption("위의 하위 카테고리 탭이나 다른 세부 단계를 선택할 수 있습니다.")

def render_building_workflow():
    """두 개의 엑셀 파일만으로 PNU 생성부터 건축물 여부 표시까지 한번에 처리."""
    st.write(
        "건축물대장과 토지조서를 올리면 **건축물대장 PNU 자동 생성 → 중복 PNU 정리 "
        "→ 필지별 건축물 여부 대조**를 한 번에 수행합니다."
    )
    st.caption("토지조서가 없으면 PNU 생성만 가능하고, 기존 PNU 생성결과 엑셀도 재사용할 수 있습니다.")
    building_excel = st.file_uploader(
        "① 건축물대장 원본 또는 PNU 생성결과 (.xlsx) · 필수",
        type=["xlsx"], key="combined_building_v143",
    )
    ledger_excel = st.file_uploader(
        "② 토지조서 또는 기존 산정명세서 엑셀 (.xlsx) · 대조할 경우 업로드",
        type=["xlsx"], key="combined_ledger_v143",
    )
    with st.expander("📥 건축물대장 다운로드 안내"):
        st.markdown("""
- [건축HUB](https://www.hub.go.kr/) → **원하는대로 건축데이터 → 유형별 건축데이터 제공 → 건축물대장 → 기본개요**
- [건축물대장 기본개요 바로가기](https://www.hub.go.kr/portal/opn/tyb/idx-bdrg.do)
- 시군구코드와 법정동코드가 있는 **기본개요**, 법정동명이 있는 **맞춤형 건축정보** 모두 지원합니다.
""")
    with st.expander("🔧 법정동코드 TXT 교체 (맞춤형 건축정보만 · 선택)"):
        code_txt = st.file_uploader(
            "최신 법정동코드 TXT", type=["txt"], key="combined_code_v143"
        )
        st.caption("업로드하지 않으면 기본 내장 코드(2026.10.14)를 사용합니다. 기본개요 형식에는 필요 없습니다.")
    with st.expander("ℹ️ 처리 기준 및 결과 안내"):
        st.markdown("""
- **PNU 생성:** 시군구코드·법정동코드·대지구분·본번·부번을 결합합니다. 블록 및 번호 누락 행은 제외 목록으로 분리합니다.
- **토지조서 대조:** `PNU` 필드가 있으면 그대로 사용합니다. 기존 산정명세서의 `토지이용계획(용도지역)` 시트는 주소와 지번으로 PNU를 생성합니다.
- **결과 엑셀:** `대조결과` 시트에는 **건축물 여부(자동)**만 깔끔하게 추가합니다. 상세 원인은 `검증상세`, `미확인_검토` 시트에서 확인합니다.
- **주의:** 미조회는 건축물이 없다는 뜻이 아니며, PNU 일치만으로 훼손 제외 여부가 확정되지 않습니다.
- 입력한 원본 엑셀 파일은 수정하지 않습니다.
""")
    files = [file for file in (building_excel, ledger_excel, code_txt) if file is not None]
    signature = tuple(
        (file.name, hashlib.blake2b(file.getvalue(), digest_size=8).hexdigest())
        for file in files
    )
    if st.session_state.get("building_input_signature_v143") != signature:
        st.session_state.pop("building_result_v143", None)

    if st.button(
        "PNU 생성 및 토지조서 대조 한 번에 실행" if ledger_excel else "건축물대장 PNU 생성 실행",
        type="primary", disabled=building_excel is None,
        use_container_width=True, key="run_combined_v143",
    ):
        try:
            with st.spinner("건축물대장 PNU 생성 및 필지별 대조 중…"):
                result = run_building_workflow(
                    building_excel.getvalue(),
                    ledger_excel.getvalue() if ledger_excel else None,
                    code_txt.getvalue() if code_txt else None,
                )
            st.session_state["building_input_signature_v143"] = signature
            st.session_state["building_result_v143"] = result
        except Exception as exc:
            st.session_state.pop("building_result_v143", None)
            st.error(f"분석 실패: {exc}")

    if "building_result_v143" in st.session_state:
        result = st.session_state["building_result_v143"]
        summary = result.pnu_stats
        st.success(f"PNU 생성 완료 · 건축물대장 형식: {result.source_kind}")
        p1, p2, p3 = st.columns(3)
        p1.metric("건축물대장 원본 행", f"{summary.get('전체', 0):,}")
        p2.metric("정상 PNU 생성 행", f"{summary.get('정상', 0):,}")
        p3.metric("고유 PNU", f"{summary.get('고유 PNU', 0):,}")
        if summary.get("제외", 0):
            st.caption(f"PNU 생성 제외: {summary['제외']:,}행 · PNU 생성 결과의 'PNU_제외목록'에서 확인")

        if result.match_bytes is not None:
            stats = result.match_stats
            st.divider()
            st.markdown("##### 토지조서 대조 결과")
            st.caption(f"대상 시트: {result.ledger_sheet}")
            m1, m2, m3 = st.columns(3)
            m1.metric("토지조서 대조 행", f"{stats.get('전체', 0):,}")
            m2.metric("건축물대장 일치 O", f"{stats.get('일치', 0):,}")
            m3.metric(
                "미조회·검토 필요",
                f"{stats.get('대장 미조회', 0) + stats.get('검토 필요', 0):,}",
            )
            st.warning("미조회 필지는 **건축물 없음 확정**이 아닙니다. 자료 범위와 기준일을 확인해 주세요.")
            st.download_button(
                "📥 필지별 건축물 여부 결과 엑셀 다운로드",
                data=result.match_bytes, file_name="필지별_건축물여부_결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="download_combined_match_v143",
            )

        with st.expander("📦 건축물대장 PNU 생성 결과 다운로드 (추가 자료)", expanded=result.match_bytes is None):
            st.caption("PNU_생성결과 / PNU_제외목록 / PNU_중복제거 / 처리요약")
            st.download_button(
                "건축물대장 PNU 결과 엑셀 다운로드",
                data=result.pnu_bytes, file_name="건축물대장_PNU_자동생성.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="download_combined_pnu_v143",
            )


def _choose_field(options, guesses, fallback=0):
    """Pick a likely field; the user can always change the selector."""
    lowered = {str(name).casefold(): i for i, name in enumerate(options)}
    for name in guesses:
        if name.casefold() in lowered:
            return lowered[name.casefold()]
    return fallback


def render_zone_workflow():
    st.write("**지적도 + 토지대장 SHP**와 **용도지역 SHP**를 업로드하면 용도지역 교차분석, 코드변환, 편입면적 비례 보정을 연속으로 실행합니다.")
    st.info("**한 번에 실행:** SHP 업로드 → 필드 확인 → 분석 및 보정 실행 → 최종면적 엑셀·QGIS용 SHP 다운로드")
    st.caption("입력: 지적도·토지대장 ZIP(SHP) 또는 GPKG 1개 + 용도지역 ZIP(SHP) 또는 GPKG 1개 이상. 원래 필지 면적이 아닌 **사업구역 내 면적**이 필요하면 사업구역계도 업로드하세요.")

    land_file = st.file_uploader(
        "① 지적도 + 토지대장 SHP ZIP 또는 GPKG", type=["zip", "gpkg"],
        key="zone_land_v14"
    )
    zone_files = st.file_uploader(
        "② 용도지역 SHP ZIP 또는 GPKG (여러 파일 동시 업로드 가능)",
        type=["zip", "gpkg"], accept_multiple_files=True, key="zone_map_v14"
    )
    with st.expander("③ 사업구역계 SHP ZIP 또는 GPKG (선택)"):
        boundary_file = st.file_uploader(
            "사업구역계 업로드: 업로드하면 대상 필지만 사업구역으로 잘라 분석",
            type=["zip", "gpkg"], key="zone_boundary_v14"
        )
        st.caption("지적도에 이미 편입구역이 잘려 있거나 전체 필지를 분석한다면 생략 가능합니다.")

    with st.expander("📋 분석 기준과 주의사항"):
        st.markdown("""
- SHP는 `.shp`, `.shx`, `.dbf`, `.prj`를 **한 ZIP에 묶어서** 업로드하세요. 한글 인코딩 정보가 담긴 `.cpg` 파일도 있다면 함께 압축하세요. ZIP 한 개에는 SHP 한 세트만 넣으세요.
- **지적도:** PNU, 지목, 편입면적, 대장면적 필드 필요
- **용도지역:** 브이월드의 도시지역·관리지역·농림지역·자연환경보전지역 자료를 각각 업로드할 수 있습니다. 모두 같은 코드 필드(`MNUM`)를 사용해야 합니다.
- `MNUM` 문자열의 **21번째 문자부터 6자리**를 추출해 내장 용도지역코드(2026.10.14)의 `UCODE / UNAME`과 대조합니다.
- 면적은 좌표계를 **EPSG:5179**로 변환해 계산한 **평면 ㎡**입니다. QGIS `$area`가 타원체 면적을 사용하면 수치가 조금 다를 수 있습니다.
- 용도지역 경계의 누락·도형 중복이 의심되더라도 배분 가능한 필지는 **자동 보정 + 별도 검토 표시**합니다. **공간검증 기준은 용도지역 교차면적 합계와 지적도 SHP 분석대상 도형면적의 비교**이며, 조서상 편입면적과 직접 비교한 것이 아닙니다.
- 원본 지적도에 있는 법정동·지번·건축물 여부 등 **기존 속성 필드를 결과 엑셀에 유지**합니다. 결과는 아직 법정 부담금 산정값이 아닙니다.
- 업로드한 공간자료는 현재 앱 실행에서만 분석합니다. 전국 원본 공간자료의 자동 다운로드 기능은 **미개발**입니다.
""")

    if not (land_file and zone_files):
        st.session_state.pop("zone_workflow_result_v149", None)
        st.caption("① 지적도와 ② 용도지역을 업로드하면 필드 선택과 분석 버튼이 표시됩니다.")
        return
    st.markdown("##### SHP 한글 인코딩 설정")
    st.caption("한글 필드명이 깨지면 인코딩을 **CP949**로 바꿔보세요. 브이월드/QGIS SHP는 파일에 따라 CP949 또는 UTF-8을 사용합니다.")
    en1, en2 = st.columns(2)
    with en1:
        parcel_encoding = st.selectbox("지적도 SHP 인코딩", ["auto", "CP949", "UTF-8"],
                                       format_func=lambda x: {"auto":"자동 감지 (.cpg 우선 확인)", "CP949":"CP949 (한글 Windows)", "UTF-8":"UTF-8"}[x],
                                       key="zone_parcel_encoding_v145")
    with en2:
        zone_encoding = st.selectbox("용도지역 SHP 인코딩", ["auto", "CP949", "UTF-8"],
                                     format_func=lambda x: {"auto":"자동 감지 (.cpg 우선 확인)", "CP949":"CP949 (한글 Windows)", "UTF-8":"UTF-8"}[x],
                                     key="zone_zone_encoding_v145")
    try:
        land_fields = list_fields(land_file.getvalue(), land_file.name, encoding=parcel_encoding)
        zone_fields = list_fields(zone_files[0].getvalue(), zone_files[0].name, encoding=zone_encoding)
        st.caption("읽기 인코딩: 지적도 "
                   f"{get_shp_encoding(land_file.getvalue(),land_file.name,parcel_encoding) or 'GPKG (UTF-8)'}"
                   " · 용도지역 "
                   f"{get_shp_encoding(zone_files[0].getvalue(),zone_files[0].name,zone_encoding) or 'GPKG (UTF-8)'}")
    except Exception as exc:
        st.session_state.pop("zone_workflow_result_v149", None)
        st.error(f"파일 검사 실패: {exc}")
        return
    if not land_fields or not zone_fields:
        st.session_state.pop("zone_workflow_result_v149", None)
        st.error("선택 가능한 속성 필드가 없습니다. QGIS에서 파일을 확인해 주세요.")
        return

    source_files = [land_file, *zone_files] + ([boundary_file] if boundary_file else [])
    source_signature = tuple(
        (f.name, hashlib.blake2b(f.getvalue(), digest_size=8).hexdigest())
        for f in source_files
    )
    source_signature = source_signature + (("encoding", parcel_encoding, zone_encoding),)

    st.markdown("##### 분석에 사용할 필드 확인")
    c1, c2 = st.columns(2)
    with c1:
        pnu_field = st.selectbox(
            "지적도 PNU 필드", land_fields,
            index=_choose_field(land_fields, ["PNU", "pnu", "필지고유번호", "pnu19"]),
            key=f"zone_pnu_field_v145_{parcel_encoding}"
        )
        landuse_field = st.selectbox(
            "지적도 지목 필드 · 필수", land_fields,
            index=_choose_field(land_fields, ["지목", "1_지목", "JIMOK"]),
            key=f"zone_landuse_field_v146_{parcel_encoding}"
        )
        included_field = st.selectbox(
            "원본 편입면적 필드 · 필수", land_fields,
            index=_choose_field(land_fields, ["편입면적", "편입면적_m2", "INCL_AREA", "included_area"]),
            key=f"zone_area_field_v146_{parcel_encoding}"
        )
        registry_field = st.selectbox(
            "원본 대장면적 필드 · 필수", land_fields,
            index=_choose_field(land_fields, ["대장면적", "1_대장면적", "LEDGER_A", "REG_AREA", "토지면적"]),
            key=f"zone_register_area_field_v146_{parcel_encoding}"
        )
    with c2:
        mnum_field = st.selectbox(
            "용도지역 MNUM 필드 (모든 파일에 공통)", zone_fields,
            index=_choose_field(zone_fields, ["MNUM", "mnum"]),
            key=f"zone_mnum_field_v145_{zone_encoding}"
        )
        st.success(f"내장 용도지역코드: {embedded_zone_count:,}개 · 기준일 2026.10.14")
        st.caption(f"공간자료 {len(zone_files)}개 병합 후 교차분석 · 분석 좌표계 EPSG:5179")
    # Fields must be distinct; otherwise the source column may be renamed twice.
    selected_parcel_fields = [pnu_field, landuse_field, included_field, registry_field]
    if len(set(selected_parcel_fields)) != 4:
        st.error("PNU·지목·편입면적·대장면적은 각각 서로 다른 필드를 선택해 주세요.")
        st.session_state.pop("zone_workflow_result_v149", None)
        return
    source_signature = source_signature + (("fields", pnu_field, landuse_field,
                                             included_field, registry_field, mnum_field),)
    st.info("**필지별 면적 정밀도 자동 적용:** 편입면적이 정수인 필지는 1㎡ 단위로, 소수 한 자리인 필지는 0.1㎡ 단위로, 소수 두 자리인 필지는 0.01㎡ 단위로 보정합니다. 원본 편입면적·대장면적은 변경하지 않으며, 보정면적 합계는 필지별 편입면적과 일치시킵니다.")
    st.caption("교차 구적면적은 소수 둘째 자리까지 확인할 수 있습니다. 최종 보정면적은 필지별 편입면적의 소수 자릿수에 따라 계산합니다. 소수 셋째 자리 이상인 편입면적은 임의 반올림하지 않고 보정 보류합니다.")
    st.caption("**최종보정면적이 0㎡인 항목은 최종 Excel 및 QGIS용 SHP에서 제외**합니다. 제외 항목은 결과 엑셀의 '0면적_제외내역'에서 확인할 수 있습니다.")
    source_signature = source_signature + (("precision", "auto_per_parcel_1_0.1_0.01_v1411"),)
    source_signature = source_signature + (("qgis_shp_format", "simple_source_cols_cp949_v1413"),)
    if st.session_state.get("zone_source_signature_v149") != source_signature:
        st.session_state.pop("zone_workflow_result_v149", None)
    if st.button("용도지역 교차분석 및 면적 자동 보정 실행", type="primary", use_container_width=True,
                 key="run_zone_workflow_v149"):
        try:
            with st.spinner("용도지역 교차분석 → 코드 변환 → 필지별 비례 보정 → 엑셀·SHP 생성 중입니다..."):
                overlay = analyze_zone_overlay(
                    parcel_file=(land_file.name, land_file.getvalue()),
                    zone_files=[(f.name, f.getvalue()) for f in zone_files],
                    pnu_field=pnu_field,
                    mnum_field=mnum_field,
                    code_map=load_zone_code_map(),
                    landuse_field=landuse_field,
                    included_area_field=included_field,
                    registry_area_field=registry_field,
                    boundary_file=(boundary_file.name, boundary_file.getvalue()) if boundary_file else None,
                    parcel_encoding=parcel_encoding,
                    zone_encoding=zone_encoding,
                    include_shp=False,
                )
                adjustment = adjust_zone_areas(overlay.detail_full, overlay.validation_full,
                                               precision=2)
                # Keep geometry from original overlay; attach corrected attributes by
                # unique analysis-row ID, retaining pending rows for spatial review.
                # Keep calculated fields only temporarily for the 0㎡ filter.
                # Exported SHP must contain ONLY original source attributes plus
                # a readable zoning name; diagnostics and adjusted areas stay in Excel.
                from geopandas import GeoDataFrame
                spatial = overlay.geometry_full.merge(adjustment.detail, on="분석행ID", how="left",
                                                       validate="one_to_one")
                spatial = GeoDataFrame(spatial, geometry="geometry", crs=overlay.geometry_full.crs)
                # Drop only *successfully adjusted* 0㎡ features from QGIS output.
                # Pending/invalid records (NaN) remain available for review.
                zero_spatial = (spatial['보정상태'].isin(['정상 보정', '보정 완료·공간검토필요'])
                                & pd.to_numeric(spatial['최종보정면적_m2'], errors='coerce').eq(0))
                spatial = spatial.loc[~zero_spatial].copy()
                if not spatial.empty:
                    qgis_frame = simple_qgis_zone_frame(spatial, overlay.original_attribute_map)
                    shp_bytes = make_shp_zip(qgis_frame)
                else:
                    shp_bytes = None
            st.session_state["zone_source_signature_v149"] = source_signature
            st.session_state["zone_workflow_result_v149"] = (overlay, adjustment, shp_bytes)
            st.session_state.pop("damage_workflow_result_v1415", None)
            st.session_state.pop("ecology_workflow_result_v1418", None)
        except Exception as exc:
            st.session_state.pop("zone_workflow_result_v149", None)
            st.error(f"교차분석 또는 면적 보정 실패: {exc}")

    saved = st.session_state.get("zone_workflow_result_v149")
    if saved is None:
        return
    overlay, adjusted, shp_bytes = saved
    st.success("용도지역 교차분석과 편입면적 보정이 완료됐습니다.")
    if adjusted.excluded_zero_segments:
        st.info(f"최종보정면적 0㎡: 용도지역 최종 엑셀 {adjusted.excluded_zero_zones:,}행 / QGIS용 SHP {adjusted.excluded_zero_segments:,}개 피처 제외. 상세 이력은 엑셀 '0면적_제외내역' 시트에 남겼습니다.")
    a, b, c, d = st.columns(4)
    a.metric("분석 대상 필지", f"{adjusted.total_parcels:,}")
    b.metric("면적 보정 완료", f"{adjusted.adjusted_parcels:,}")
    c.metric("공간검토 필요", f"{adjusted.review_parcels:,}")
    d.metric("보정 보류", f"{adjusted.pending_parcels:,}")
    if adjusted.adjusted_parcels:
        st.success("보정 완료된 필지는 **PNU별 용도지역 최종면적 합계 = 편입면적**을 충족합니다.")
        st.dataframe(adjusted.summary, hide_index=True, use_container_width=True)
    if adjusted.review_parcels:
        st.warning(f"{adjusted.review_parcels:,}개 필지는 면적을 비례 보정했지만, 실제 용도지역 누락·중복 가능성을 확인해야 합니다. 검토 전 법정 산정값으로 확정하지 마세요.")
        with st.expander("공간검토 필요 필지 및 사유"):
            st.dataframe(adjusted.review, hide_index=True, use_container_width=True)
    if adjusted.pending_parcels:
        st.warning(f"PNU·코드·편입면적 등 오류가 있는 {adjusted.pending_parcels:,}개 필지는 보류했습니다.")
        with st.expander("보정 보류 필지 및 사유"):
            st.dataframe(adjusted.pending, hide_index=True, use_container_width=True)
    st.download_button(
        "📥 용도지역별 최종 보정면적 엑셀 다운로드",
        data=adjusted.excel_bytes,
        file_name="용도지역별_최종보정면적_V1.4.13.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, key="download_zone_final_v149",
    )
    if shp_bytes:
        st.download_button(
            "📥 QGIS용 간편 교차결과 SHP(ZIP) 다운로드",
            data=shp_bytes,
            file_name="용도지역_교차결과_V1.4.13.zip",
            mime="application/zip", use_container_width=True,
            key="download_zone_final_shp_v149",
        )
        st.caption("SHP에는 **원본 지적도 필드 + 용도지역명**만 들어갑니다. 최종보정면적·보정상태·검증항목은 엑셀에만 남깁니다. 한글 필드명은 CP949 기준으로 원래 이름을 우선 유지하되 10바이트 초과 이름만 줄이며 대조표를 제공합니다.")
    st.caption("원본 지적도의 법정동·지번·건축물 여부 등 기존 속성필드를 유지합니다. 다만 별도 엑셀에만 있는 값은 자동 결합되지 않습니다.")


def render_damage_workflow():
    """2-2 결과를 받아 지목/용도지역/건축물 검토 기준으로 SHP·엑셀 생성."""
    st.write("**2-2 용도지역 분석 결과**를 이어받아 훼손 제외 후보와 산정대상면적 구간을 나눕니다.")
    st.info("실행 순서: 2-2 용도지역 교차분석·보정 실행 → 이 화면에서 검토 조건 확인 → 훼손 제외면적 분석 실행")
    st.warning("**수기 명세서 방식의 검증용 결과입니다.** 1차 → 2차 → 3차 순서로 중복 없이 제외합니다. 건축물 O는 1·2차에서 제외되지 않은 면적만 3차로 계산하며, 실제 법적 확정 여부는 증빙 확인이 필요합니다.")
    current = st.session_state.get("zone_workflow_result_v149")
    if current is None:
        st.error("2-2 용도지역 분석 결과가 없습니다. 같은 웹사이트 실행에서 **2-2 용도지역 분석**을 먼저 완료해 주세요.")
        st.caption("2-2에서 사용한 원본 SHP는 수정되지 않습니다. 여기서는 자동 보정된 면적을 그대로 재사용합니다.")
        return
    overlay, adjustment, _ = current
    st.success(f"2-2 연결 완료: 보정 완료 필지 {adjustment.adjusted_parcels:,}개 · 보류 {adjustment.pending_parcels:,}개")

    with st.expander("검증용 제외 규칙 확인 / 수정", expanded=True):
        st.caption("기본값은 기존 수기 업무매뉴얼에 근거한 검토 가정입니다. 사업에 맞는 지목 목록인지 반드시 확인하세요.")
        apply_first = st.checkbox("1차 제외 지목 규칙 적용", value=True, key="damage_use_first_v1415")
        first_text = st.text_input("1차 제외 후보 지목 (쉼표 구분)", value=",".join(DEFAULT_FIRST), key="damage_first_v1415")
        apply_second = st.checkbox("2차 용도지역·지목 조합 규칙 적용", value=True, key="damage_use_second_v1415")
        st.caption("2차 조건: 주거·상업·공업·계획관리지역에서 아래 보존 지목에 해당하지 않는 구간을 제외 후보로 분류")
        retain_text = st.text_input("2차 보존 지목 (쉼표 구분)", value=",".join(DEFAULT_RETAIN), key="damage_retain_v1415")

    st.markdown("##### 건축물 여부 자동 인식")
    candidate_fields = find_building_columns(adjustment.detail)
    if candidate_fields:
        detected = []
        for col in candidate_fields:
            positives = adjustment.detail[col].map(_text_match_building)
            positive_pnu = adjustment.detail.loc[positives, 'PNU'].astype(str).nunique()
            detected.append(f"{col}: O {positive_pnu:,}개 필지")
        st.success("지적도 원본의 건축물 여부 필드를 자동으로 인식했어: " + " / ".join(detected))
    else:
        st.warning("2-2 자료에서 건축물 여부 필드를 자동으로 찾지 못했어. 아래에서 수동 선택하거나 2-1 결과 엑셀을 추가해 줘.")
    selectable = [c for c in adjustment.detail.columns
                  if c not in ('분석행ID', 'PNU', '지목', '용도지역명', '보정상태', '판정사유', '최종보정면적_m2', '원본편입면적')]
    manual_field = st.selectbox("건축물 여부 필드 수동 지정 (자동 인식 실패 시)",
                                ['자동 감지'] + selectable, index=0,
                                key='damage_building_field_v1415')
    source_field = None if manual_field == '자동 감지' else manual_field
    building_upload = st.file_uploader("또는 2-1 '필지별 건축물 여부 결과' 엑셀 업로드 (.xlsx)", type=["xlsx"], key="damage_match_xlsx_v1415")
    existing_building = st.session_state.get("building_result_v143")
    building_from_previous = (existing_building.match_bytes if existing_building is not None else None)
    building_data = building_upload.getvalue() if building_upload else building_from_previous
    if building_data:
        st.caption("2-1에서 얻은 건축물 여부 엑셀을 PNU로 연결합니다. 엑셀의 건축물 O를 우선 반영합니다.")
    else:
        st.caption("2-1 엑셀을 올리지 않아도 지적도 원본 건축물 여부의 O를 자동 반영합니다. O가 확인되지 않은 값은 미조회이며 건축물 없음 확정은 아닙니다.")
    exclude_buildings = st.checkbox(
        "건축물 O는 1·2차 제외 후 남은 구간만 3차 제외 후보에 포함 (기본 적용)",
        value=True, key="damage_exclude_building_v1415"
    )
    if exclude_buildings:
        st.warning("건축물 O는 1·2차에서 이미 제외된 면적과 중복 계산하지 않습니다. 건축물 존재만으로 법적 제외가 확정되는 것은 아니므로 증빙을 확인해 주세요.")

    first = tuple(z.strip() for z in first_text.replace('，', ',').split(',') if z.strip())
    retained = tuple(z.strip() for z in retain_text.replace('，', ',').split(',') if z.strip())
    source_sig = (st.session_state.get("zone_source_signature_v149"),
                  first, retained, apply_first, apply_second,
                  source_field, bool(exclude_buildings),
                  hashlib.blake2b(building_data, digest_size=8).hexdigest() if building_data else None)
    if st.session_state.get("damage_source_signature_v1415") != source_sig:
        st.session_state.pop("damage_workflow_result_v1415", None)

    if st.button("훼손 제외면적 분석 실행", type="primary", use_container_width=True, key="run_damage_v1415"):
        try:
            with st.spinner("제외 조건 적용 → 필지별 면적 검증 → 엑셀 및 두 종류의 SHP 생성 중..."):
                result = analyze_damage_filter(
                    overlay.geometry_full, adjustment.detail, adjustment.validation,
                    landuse_excluded=first, zone_retain_landuse=retained,
                    building_field=source_field,
                    building_excel=building_data,
                    exclude_building_parcels=exclude_buildings,
                    apply_landuse_rule=apply_first, apply_zone_rule=apply_second,
                )
            st.session_state["damage_source_signature_v1415"] = source_sig
            st.session_state["damage_workflow_result_v1415"] = result
            st.session_state.pop("ecology_workflow_result_v1418", None)
        except Exception as exc:
            st.session_state.pop("damage_workflow_result_v1415", None)
            st.session_state.pop("ecology_workflow_result_v1418", None)
            st.error(f"훼손 제외면적 검증 실패: {exc}")

    result = st.session_state.get("damage_workflow_result_v1415")
    if result is None:
        return
    st.success("훼손 제외면적 검증자료를 생성했어. 제외와 잔여 두 SHP는 서로 다른 레이어야.")
    left, right, pending = st.columns(3)
    left.metric("제외 후보 면적", f"{result.excluded_m2:,.2f}㎡")
    right.metric("생태계보전부담금 산정대상면적", f"{result.remaining_m2:,.2f}㎡")
    pending.metric("3차 제외(시설물 설치)", f"{result.third_m2:,.2f}㎡")
    st.caption(f"제외 단계별 면적: 1차 {result.first_m2:,.2f}㎡ · 2차 {result.second_m2:,.2f}㎡ · 3차 {result.third_m2:,.2f}㎡ / 보정 보류 {result.pending_count:,}구간")
    positive_parcels = result.detail.loc[
        result.detail['건축물여부_검증'].eq('O'), 'PNU'
    ].astype(str).nunique()
    if positive_parcels:
        st.info(f"건축물 O 인식: {positive_parcels:,}개 필지 · 3차 잔여 제외 후보 적용: {'예' if exclude_buildings else '아니오 (검토 대상)'}")
    else:
        st.warning("건축물 O로 인식된 필지가 0개야. 원본 SHP의 필드명·표시값을 확인하거나 2-1 건축물 여부 엑셀을 연결해 줘.")
    st.caption("표시된 합계는 2-2에서 조정한 최종보정면적의 합계입니다. 도형의 QGIS $area와는 다를 수 있습니다.")
    if result.building_review_count:
        st.warning(f"건축물 O로 확인되었으나 전체 구간 제외를 확정하지 않은 {result.building_review_count:,}개 교차 구간이 잔여 레이어에 포함되어 있습니다.")
    if result.spatial_review_count:
        st.warning(f"2-2에서 공간 누락·중복 검토가 필요한 필지 {result.spatial_review_count:,}개를 포함합니다. 숫자상 면적 일치와 공간적 타당성은 별도로 확인해 주세요.")
    st.dataframe(result.summary, hide_index=True, use_container_width=True)
    st.caption("다운로드 Excel은 '산정명세서'와 '토지이용계획(용도지역)' 두 시트만 제공하며, 지번은 원본 속성 우선·없으면 PNU에서 복원합니다. PNU별 검증표·공간 경고 등의 내부 정보는 화면 및 계산 단계에서 유지합니다.")
    st.download_button("📥 훼손 제외면적 검증 엑셀 다운로드", result.excel_bytes,
                       file_name="훼손_제외면적_검증_V1.4.16.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True, key="damage_download_excel_v1415")
    if result.excluded_shp:
        st.download_button("📥 훼손 제외 후보 SHP ZIP 다운로드", result.excluded_shp,
                           file_name="훼손제외_대상_V1.4.17.zip", mime="application/zip",
                           use_container_width=True, key="damage_download_exclude_v1415")
    else:
        st.info("검증 규칙에 해당하는 제외 후보가 없어 제외 대상 SHP는 생성하지 않았습니다.")
    if result.remain_shp:
        st.download_button("📥 부담금 산정대상면적 SHP ZIP 다운로드", result.remain_shp,
                           file_name="부담금_산정대상면적_V1.4.19.zip", mime="application/zip",
                           use_container_width=True, key="damage_download_remain_v1415")
    else:
        st.info("잔여 도형이 없어 산정대상면적 SHP는 생성하지 않았습니다.")
    st.caption("SHP는 PNU·지목·용도지역·구분·사유·보정면적 등 최소 필드만 포함합니다. 보정면적은 2-2 결과로 계산하며, SHP의 도형 면적과 다를 수 있습니다.")
    st.caption("결과 엑셀은 산정명세서와 토지이용계획(용도지역) 2시트로 정리했습니다. 행별 1·2·3차 제외 합계 + 산정대상면적 = 편입면적을 확인할 수 있습니다. 생태자연도 등급별 면적은 다음 2-4에서 분석하며, 법적 제외 확정은 별도 검토입니다.")


@st.cache_data(show_spinner=False)
def _ecology_input_fields(name: str, payload: bytes, encoding: str) -> tuple[list[str], dict[str, list[str]]]:
    frame = read_spatial(payload, name, encoding=encoding)
    cols = [str(c) for c in frame.columns if c != frame.geometry.name]
    categories = {}
    for col in cols:
        vals = frame[col].dropna().astype(str).str.strip().value_counts()
        if 0 < len(vals) <= 60:
            categories[col] = vals.index.tolist()
    return cols, categories


def render_ecology_workflow():
    st.write("**산정대상면적 SHP와 생태자연도 SHP**를 비교해 등급별 면적을 계산하고, "
             "1등급·2등급·별도관리지역 이외의 남는 공간을 **전부 3등급**으로 계산합니다.")
    st.caption("생태자연도 SHP에 3등급 폴리곤이 없어도 잔여 도형의 차집합으로 생성합니다. "
               "필지·용도지역 구간별 등급면적 합계는 2-3에서 보정한 '보정면적'과 정확히 일치시킵니다.")

    previous = st.session_state.get('damage_workflow_result_v1415')
    use_previous = bool(previous is not None and previous.remain_shp)
    if use_previous:
        reuse = st.checkbox("현재 2-3 훼손 산정대상면적 SHP 자동 연결", value=True, key='eco_auto_remain_v1418')
    else:
        reuse = False
    manual_remain = None
    if reuse:
        remaining_name = '부담금_산정대상면적.zip'
        remaining_bytes = previous.remain_shp
        st.success('2-3 산정대상면적 SHP 연결 완료')
    else:
        manual_remain = st.file_uploader('① 산정대상면적 SHP ZIP 또는 GPKG',
                                         type=['zip','gpkg'], key='eco_remaining_v1418')
        remaining_name = manual_remain.name if manual_remain else None
        remaining_bytes = manual_remain.getvalue() if manual_remain else None
    eco_files = st.file_uploader('② 생태자연도 SHP ZIP · 도엽이 여러 개면 모두 선택',
                                 type=['zip','gpkg'], accept_multiple_files=True,
                                 key='eco_grades_v1418')
    if not remaining_bytes or not eco_files:
        st.info('산정대상면적 SHP와 생태자연도 SHP를 준비하면 분석 설정이 나타납니다.')
        return

    enc_col1, enc_col2 = st.columns(2)
    with enc_col1:
        remaining_encoding = st.selectbox('잔여 SHP 인코딩', ['auto','CP949','UTF-8'],
                                           key='eco_remain_encoding_v1418')
    with enc_col2:
        ecology_encoding = st.selectbox('생태자연도 SHP 인코딩', ['auto','CP949','UTF-8'],
                                         key='eco_grade_encoding_v1418')
    try:
        remain_cols, _ = _ecology_input_fields(remaining_name, remaining_bytes, remaining_encoding)
        all_cols = None
        values_per_col = {}
        for item in eco_files:
            fieldnames, cats = _ecology_input_fields(item.name, item.getvalue(), ecology_encoding)
            all_cols = set(fieldnames) if all_cols is None else all_cols.intersection(fieldnames)
            for field, vals in cats.items():
                values_per_col.setdefault(field, set()).update(vals)
        grade_fields = sorted(all_cols or [])
        if not grade_fields:
            st.error('생태자연도 파일에서 공통 속성 필드를 찾지 못했습니다.')
            return
    except Exception as exc:
        st.error(f'공간자료 확인 실패: {exc}')
        return

    if not {'PNU', '보정면적'}.issubset(remain_cols):
        st.warning('2-3에서 다운로드한 잔여 SHP라면 PNU와 보정면적 필드가 있어야 합니다. 다른 필드명인 경우 수동 선택하세요.')
    c1, c2, c3 = st.columns(3)
    with c1:
        pnu_name = st.selectbox('잔여 SHP의 PNU 필드', remain_cols,
                                index=remain_cols.index('PNU') if 'PNU' in remain_cols else 0,
                                key='eco_pnu_v1418')
    with c2:
        area_name = st.selectbox('잔여 SHP의 최종 보정면적 필드', remain_cols,
                                 index=remain_cols.index('보정면적') if '보정면적' in remain_cols else 0,
                                 key='eco_area_v1418')
    with c3:
        suggestions=['ecozmGrd','eczmGr','ecozmGr','ecozmGrd','ECOZMGRD','ECZMGR','등급','생태등급']
        recognized = next((v for v in suggestions if v in grade_fields), None)
        grade_field = st.selectbox('생태자연도 등급 필드', grade_fields,
                                   index=grade_fields.index(recognized) if recognized else 0,
                                   key='eco_grade_field_v1418')
    grade_values = sorted(values_per_col.get(grade_field, set()))
    if not grade_values:
        st.error('등급 필드에서 값 목록을 찾을 수 없습니다. 선택한 필드가 맞는지 확인하세요.')
        return
    defaults = default_grade_values(grade_values)
    st.markdown('##### 생태자연도 등급 코드 확인')
    st.caption('원본 파일의 실제 등급값을 선택하세요. 숫자 0이나 B가 별도관리지역인지는 추측하지 않으며, 직접 확인해야 합니다.')
    c1, c2, c3 = st.columns(3)
    with c1:
        select_1 = st.multiselect('1등급 값', options=grade_values, default=defaults['1등급'],
                                  key=f'eco_value1_{grade_field}_v1418')
    with c2:
        select_2 = st.multiselect('2등급 값', options=grade_values, default=defaults['2등급'],
                                  key=f'eco_value2_{grade_field}_v1418')
    with c3:
        select_s = st.multiselect('별도관리지역 값', options=grade_values, default=defaults['별도관리지역'],
                                  key=f'eco_values_{grade_field}_v1418')
    remaining_values = [v for v in grade_values if v not in (set(select_1)|set(select_2)|set(select_s))]
    st.caption(f'기타 값 (모두 3등급 취급): {", ".join(remaining_values[:20]) or "없음"}. '
               '아예 SHP에 없는 공간도 3등급으로 생성합니다.')
    if not select_1 and not select_2 and not select_s:
        st.warning('1등급·2등급·별도관리지역이 하나도 선택되지 않았습니다. 실제 자료가 전부 3등급인지 확인해 주세요.')

    signatures = ((remaining_name, hashlib.blake2b(remaining_bytes,digest_size=8).hexdigest()),)
    signatures += tuple((f.name,hashlib.blake2b(f.getvalue(),digest_size=8).hexdigest()) for f in eco_files)
    signatures += (grade_field, pnu_name, area_name, remaining_encoding, ecology_encoding,
                   tuple(select_1), tuple(select_2), tuple(select_s))
    if st.session_state.get('ecology_signature_v1418') != signatures:
        st.session_state.pop('ecology_workflow_result_v1418', None)

    if st.button('생태자연도 교차분석 및 등급별 면적 자동 보정 실행',
                 type='primary', use_container_width=True, key='eco_run_v1418'):
        try:
            with st.spinner('생태자연도 공간 교차 → 잔여공간 3등급 생성 → 편입면적 비례 보정 → 엑셀·SHP 출력 중…'):
                report = analyze_ecology(
                    (remaining_name, remaining_bytes),
                    [(f.name, f.getvalue()) for f in eco_files],
                    grade_field=grade_field, special_values=select_s,
                    grade1_values=select_1, grade2_values=select_2,
                    pnu_field=pnu_name, area_field=area_name,
                    remaining_encoding=remaining_encoding, ecology_encoding=ecology_encoding,
                )
            st.session_state['ecology_signature_v1418'] = signatures
            st.session_state['ecology_workflow_result_v1418'] = report
        except Exception as exc:
            st.session_state.pop('ecology_workflow_result_v1418', None)
            st.error(f'생태자연도 분석 실패: {exc}')
    report = st.session_state.get('ecology_workflow_result_v1418')
    if report is None:
        return
    st.success('생태자연도 등급별 면적 분석 완료 · 산정대상면적과 최종 등급면적 합계 일치')
    st.dataframe(report.summary, hide_index=True, use_container_width=True)
    st.metric('최종 생태등급별 면적 합계', f'{report.total_m2:,.2f}㎡')
    if report.overlap_fragments:
        st.warning(f'서로 다른 등급 경계가 겹치는 잔여 구간 {report.overlap_fragments:,}개 (약 {report.overlap_area_m2:,.2f}㎡). '
                   '중복계산은 방지했지만 별도관리 > 1등급 > 2등급 임시 우선순위를 적용했으니 실제 자료를 검토하세요.')
    st.download_button('📥 생태자연도 등급별 면적 엑셀 다운로드', report.excel_bytes,
                       file_name='생태자연도_등급별_면적_V1.4.18.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                       use_container_width=True, key='eco_download_xlsx_v1418')
    if report.shp_zip:
        st.download_button('📥 생태자연도 등급 교차결과 SHP ZIP 다운로드', report.shp_zip,
                           file_name='생태자연도_등급교차결과_V1.4.18.zip',
                           mime='application/zip', use_container_width=True, key='eco_download_shp_v1418')
    st.caption('엑셀에는 생태자연도_집계·필지별_생태등급 2개 시트가 들어갑니다. '
               'SHP에는 PNU·지목·용도지역·생태등급·등급면적만 기록합니다. '
               '도형의 실제 구적면적과 보정된 등급면적은 다를 수 있습니다.')


with processing:
    # ② 데이터 처리 탭을 누르자마자, 제목이나 카드보다 먼저 하위 탭을 표시합니다.
    # 상단: ① 기초자료 준비 | ② 데이터 처리 | ③ 산정명세서 작성
    # 하단: 2-1 건축물대장 정리 | 2-2 용도지역 분석 | ...
    category_labels = []
    for number, category in enumerate(PROCESS_CATEGORIES, start=1):
        undeveloped = all(
            not PROCESS_STEPS[step_id]["developed"] for step_id in category["steps"]
        )
        suffix = " (미개발)" if undeveloped else ""
        category_labels.append(f"2-{number} {category['name'].split('. ', 1)[1]}{suffix}")

    category_tabs = st.tabs(category_labels)
    for number, (tab, category) in enumerate(zip(category_tabs, PROCESS_CATEGORIES), start=1):
        with tab:
            # 하위 카테고리로 들어간 다음에만 상세 작업이 보입니다.
            st.subheader(f"2-{number}. {category['name'].split('. ', 1)[1]}")
            st.caption(category["description"])
            step_ids = category["steps"]
            widget_key = f"process_step_choice_v134_{category['id']}"

            # 세부 작업 번호는 전체 데이터 처리 순번이 아니라 각 2-x 카테고리 안에서 ①부터 시작합니다.
            step_symbols = ("①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩")

            def local_step_label(step_id: str) -> str:
                local_index = step_ids.index(step_id)
                local_number = step_symbols[local_index] if local_index < len(step_symbols) else f"{local_index + 1}."
                step = PROCESS_STEPS[step_id]
                return f"{local_number} {step['name']}{' (미개발)' if not step['developed'] else ''}"

            if len(step_ids) > 1:
                selected_step = st.radio(
                    "세부 작업 단계",
                    options=step_ids,
                    format_func=local_step_label,
                    horizontal=True,
                    key=widget_key,
                )
            else:
                selected_step = step_ids[0]

            local_index = step_ids.index(selected_step)
            step_info = PROCESS_STEPS[selected_step]
            st.caption(
                f"현재 카테고리의 세부 작업 {local_index + 1}/{len(step_ids)}단계 · "
                "진행 순서를 뜻하며 실제 작업 완료율은 아닙니다."
            )
            st.markdown(f"#### {local_step_label(selected_step)}")
            st.caption(step_info["description"])

            if selected_step == "building_workflow":
                render_building_workflow()
            elif selected_step == "zone_workflow":
                render_zone_workflow()
            elif selected_step == "damage_filter":
                render_damage_workflow()
            elif selected_step == "ecology_workflow":
                render_ecology_workflow()
            else:
                render_unfinished_step(selected_step)

            # 순차 작업 버튼은 현재 2-x 카테고리 내부에서만 작동합니다.
            # Streamlit 버전 호환성을 위해 다른 탭을 강제로 여는 버튼은 만들지 않습니다.
            if len(step_ids) > 1:
                st.divider()
                prev_col, next_col = st.columns(2)
                with prev_col:
                    st.button(
                        "← 이전 세부 단계", 
                        key=f"previous_step_v134_{category['id']}",
                        disabled=local_index == 0,
                        on_click=select_category_step,
                        args=(widget_key, step_ids[max(0, local_index - 1)]),
                        use_container_width=True,
                    )
                with next_col:
                    st.button(
                        "다음 세부 단계 →",
                        key=f"next_step_v134_{category['id']}",
                        disabled=local_index == len(step_ids) - 1,
                        on_click=select_category_step,
                        args=(widget_key, step_ids[min(len(step_ids) - 1, local_index + 1)]),
                        use_container_width=True,
                    )
            if number < len(PROCESS_CATEGORIES):
                st.caption(
                    f"다음 카테고리: 위의 **2-{number + 1}** 탭을 선택하세요. "
                    "입력 자료 및 생성 결과는 같은 실행 세션 안에서 유지됩니다."
                )
            else:
                st.caption("데이터 처리 마지막 카테고리입니다. 이후 ③ 산정명세서 작성으로 이어집니다.")

with statement:
    st.subheader('3. 생태계보전부담금 산정명세서 작성')
    st.caption('2-3 훼손 제외면적과 2-4 생태자연도 등급별 면적을 연결해 산정명세서를 작성합니다.')
    dr=st.session_state.get('damage_workflow_result_v1415')
    er=st.session_state.get('ecology_workflow_result_v1418')
    if dr is None or er is None:
        st.warning('먼저 2-3 훼손 제외면적 분석과 2-4 생태자연도 분석을 같은 실행 세션에서 완료해 주세요.')
    else:
        st.metric('생태계보전부담금 산정대상면적',f'{dr.remaining_m2:,.2f}㎡')
        col1,col2=st.columns(2)
        with col1:
            project_name=st.text_input('사업명',key='stmt_project_v1421')
            project_location=st.text_input('위치',key='stmt_location_v1421')
            approval_note=st.text_input('사업승인일 / 적용 기준일',key='stmt_approval_v1421')
        with col2:
            authority=st.text_input('인허가권자',key='stmt_authority_v1421')
        st.markdown('#### 사업자 정보 (여러 사업자 입력 가능)')
        owner_count=st.number_input('사업자 수',min_value=1,max_value=20,value=1,step=1,key='stmt_owner_count_v1421')
        owners=[]
        for n in range(int(owner_count)):
            with st.expander(f'사업자 {n+1}',expanded=n==0):
                c1,c2=st.columns(2)
                with c1:
                    owner_name=st.text_input('사업자명',key=f'stmt_owner_name_{n}_v1421')
                    owner_addr=st.text_input('소재지',key=f'stmt_owner_addr_{n}_v1421')
                with c2:
                    owner_biz=st.text_input('사업자등록번호',key=f'stmt_owner_biz_{n}_v1421')
                    owner_corp=st.text_input('법인등록번호',key=f'stmt_owner_corp_{n}_v1421')
                owners.append({'name':owner_name,'address':owner_addr,'business_number':owner_biz,'corporate_number':owner_corp})
        st.info('「자연환경보전법 시행령」 제38조제1항에 따른 단위면적당 부과금액은 300원/㎡입니다. 지역계수는 같은 조 제2항 및 별표 2의5에 따르며 아래에서 용도지역별로 확인·수정할 수 있습니다. 가산금·절사 등은 적용 시점의 기준을 별도로 확인해야 합니다.')
        unit_price=st.number_input('단위면적당 부과금액 (원/㎡) · 자연환경보전법 시행령 제38조제1항',min_value=1.0,value=300.0,step=10.0,key='stmt_unit_price_v1419')
        st.markdown('#### 지역계수 (용도지역별)')
        st.caption('「자연환경보전법 시행령」 제38조제2항 및 별표 2의5에 따른 지역계수입니다. 사업의 적용 기준일을 확인하고 필요하면 수정하세요.')
        zone_factor_labels = (
            '주거·상업·공업·계획관리지역', '녹지지역', '생산관리지역',
            '농림지역', '보전관리지역', '자연환경보전지역'
        )
        default_region_factors = (1.0, 2.0, 2.5, 3.0, 3.5, 4.0)
        region_factors = []
        for row_start in (0, 3):
            region_cols = st.columns(3)
            for offset, col in enumerate(region_cols):
                idx = row_start + offset
                with col:
                    factor = st.number_input(
                        f'{zone_factor_labels[idx]} 지역계수',
                        min_value=0.0, value=default_region_factors[idx], step=0.5,
                        format='%.2f', key=f'stmt_region_factor_{idx}_v1422'
                    )
                    region_factors.append(factor)
        st.caption('입력한 지역계수는 산정명세서 계산에도 반영됩니다. 기본 계수는 관련 법령을 참고한 값이며 적용 시점에 따라 검토가 필요합니다.')
        confirm=st.checkbox('참고 단가·계수·산식을 검토했으며, 이번 엑셀은 검증용 초안으로 작성합니다.',key='stmt_confirm_v1419')
        if st.button('산정명세서 생성',type='primary',use_container_width=True,key='stmt_run_v1419',disabled=not confirm):
            try:
                excel,totals=build_statement(dr,er,project_name=project_name,project_location=project_location,
                    owners=owners,authority=authority,unit_price=unit_price,approval_note=approval_note,region_factors=region_factors)
                st.session_state['statement_file_v1419']=excel
                st.session_state['statement_signature_v1419']=(project_name,project_location,authority,tuple(tuple(o.values()) for o in owners),unit_price,approval_note,tuple(region_factors))
                st.success('면적 일치 검증 통과 · 산정명세서 작성 완료')
            except Exception as exc:
                st.session_state.pop('statement_file_v1419',None)
                st.error(f'산정명세서 생성 불가: {exc}')
        if st.session_state.get('statement_signature_v1419')==(project_name,project_location,authority,tuple(tuple(o.values()) for o in owners),unit_price,approval_note,tuple(region_factors)):
            payload=st.session_state.get('statement_file_v1419')
            if payload:
                st.download_button('📥 생태계보전부담금 산정명세서 엑셀 다운로드',payload,
                    file_name='생태계보전부담금_산정명세서_V1.4.22.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
        st.caption('출력 시트: 산정명세서 / 토지이용계획(용도지역) / 토지이용계획(생태·자연도). 확인되지 않은 면적·용도지역은 임의로 계산하지 않고 오류로 중단합니다.')

st.divider()
st.caption("V1.4.22 · ③ 생태계보전부담금 산정명세서 작성(검토용) 지원")
