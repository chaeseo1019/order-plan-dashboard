"""
생산입고현황 원본 데이터 전처리
--------------------------------
ERP에서 그대로 뽑은 생산입고현황 시트는 다음 두 가지 문제가 있어서,
계산에 쓰기 전에 정리가 필요합니다.

1. '일자' 컬럼 끝에 ' -1', ' -2' 같은 순번 접미사가 붙어있음
   예: '2025/01/02 -1' -> '2025-01-02'
2. 맨 아래에 합계행이나 날짜만 있고 품목코드가 비어있는 행이 섞여 들어올 수 있음
   -> 품목코드가 없는 행은 통째로 제외

Streamlit(app.py)과 분리된 순수 함수라 pytest 등으로 단위 테스트하기 쉽습니다.
"""

from __future__ import annotations

import re

import pandas as pd

# 일자 문자열 맨 끝에 붙는 ' -1', ' -2', ' -12' 같은 순번 접미사 (공백 포함)
_DATE_SUFFIX_PATTERN = re.compile(r"\s*-\d+\s*$")


def clean_production_history(
    df: pd.DataFrame,
    date_col: str = "일자",
    code_col: str = "품목코드",
) -> pd.DataFrame:
    """
    생산입고현황 원본 데이터프레임을 정리해서 반환합니다.

    - `code_col`이 비어있는 행(합계행, 날짜만 있는 행 등) 제거
    - `date_col` 끝의 ' -숫자' 접미사를 제거하고 'YYYY-MM-DD' 형식으로 통일
    - 날짜 형식이 완전히 깨져서 파싱이 안 되는 행도 함께 제거

    원본은 건드리지 않고 새 데이터프레임을 반환합니다.
    """
    cleaned = df.copy()

    # 1) 품목코드가 비어있는 행 제거 (합계행/날짜만 있는 행 등 방지)
    cleaned = cleaned[cleaned[code_col].notna()]
    cleaned = cleaned[cleaned[code_col].astype(str).str.strip() != ""]

    # 2) 일자 끝의 ' -숫자' 접미사 제거
    raw_dates = cleaned[date_col].astype(str).str.strip()
    stripped_dates = raw_dates.str.replace(_DATE_SUFFIX_PATTERN, "", regex=True)

    # 3) 날짜 파싱 후 'YYYY-MM-DD'로 통일 (구분자가 '/'든 '-'든 상관없이 처리)
    parsed = pd.to_datetime(stripped_dates, errors="coerce")
    cleaned[date_col] = parsed.dt.strftime("%Y-%m-%d")

    # 4) 날짜 파싱에 실패한 행(형식이 완전히 다른 이상치)도 제거
    cleaned = cleaned[cleaned[date_col].notna()]

    return cleaned.reset_index(drop=True)
