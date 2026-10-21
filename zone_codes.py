"""내장 용도지역코드 데이터 조회 (엑셀 원본 기반)."""
from __future__ import annotations
from functools import lru_cache
from pathlib import Path
import json

BASE_PATH = Path(__file__).resolve().parent / "기준자료"
SOURCE_EXCEL = BASE_PATH / "용도지역코드.xlsx"
MAPPING_JSON = BASE_PATH / "용도지역코드_조회.json"

@lru_cache(maxsize=1)
def load_zone_code_map() -> dict[str, str]:
    """사용자 업로드 없이 번들된 UCODE→UNAME 1,270개 코드를 로드."""
    if not SOURCE_EXCEL.is_file() or not MAPPING_JSON.is_file():
        raise FileNotFoundError("용도지역코드 기본 내장 자료가 없습니다. 기준자료 폴더를 확인하세요.")
    records = json.loads(MAPPING_JSON.read_text(encoding="utf-8"))
    if not isinstance(records, dict):
        raise ValueError("내장 용도지역코드가 올바르지 않습니다.")
    return records

def get_zone_name(ucode: str) -> str | None:
    """향후 공간 분석 단계에서 MNUM에서 추출한 코드를 명칭으로 변환할 때 사용."""
    return load_zone_code_map().get(str(ucode).strip())
