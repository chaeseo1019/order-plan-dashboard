"""
발주 계획 계산 로직
-------------------
엑셀 버전(발주계획 시트)에서 쓰던 것과 동일한 계산 규칙을 파이썬 함수로 옮긴 모듈입니다.
Streamlit(app.py)과 완전히 분리되어 있어서, 별도 pytest 등으로 단위 테스트하기 쉽습니다.
"""

from __future__ import annotations

import math
from datetime import date, timedelta
from typing import Optional

import pandas as pd


def _num(value, default: float = 0.0) -> float:
    """None / NaN / 빈 문자열을 안전하게 숫자로 변환합니다."""
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass
    if isinstance(value, str) and value.strip() == "":
        return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def compute_three_month_avg_outbound(
    flow_df: pd.DataFrame,
    item_code: str,
    asof_date: date,
    months: int = 3,
) -> float:
    """
    월별재고수불부에서 asof_date가 속한 달을 포함해 최근 `months`개월의
    평균 출고량을 계산합니다. (엑셀의 AVERAGEIFS(집계기간) 로직과 동일)

    flow_df: 컬럼 ['기준날짜'(YYYY-MM 문자열), '품목코드', '출고', ...]
    """
    if flow_df is None or flow_df.empty:
        return 0.0

    if isinstance(asof_date, pd.Timestamp):
        asof_date = asof_date.date()

    asof_period = pd.Period(asof_date, freq="M")
    start_period = asof_period - (months - 1)
    month_strs = {p.strftime("%Y-%m") for p in pd.period_range(start_period, asof_period, freq="M")}

    sub = flow_df[
        (flow_df["품목코드"].astype(str) == str(item_code))
        & (flow_df["기준날짜"].astype(str).isin(month_strs))
    ]
    if sub.empty:
        return 0.0
    return float(sub["출고"].mean())


def compute_material_requirement(production_qty, unit_usage, loss_rate) -> int:
    """
    자재 소요량 = 생산 지시량 × 개당실제투입량 × (1 + 예상로스율), 올림 처리
    """
    production_qty = _num(production_qty)
    unit_usage = _num(unit_usage)
    loss_rate = _num(loss_rate)
    raw = production_qty * unit_usage * (1 + loss_rate)
    return math.ceil(raw)


def compute_safety_stock(avg_outbound, months_buffer) -> int:
    """안전재고 = 안전재고기준(개월) × 3개월평균출고량"""
    avg_outbound = _num(avg_outbound)
    months_buffer = _num(months_buffer)
    return round(avg_outbound * months_buffer)


def compute_required_order_qty(demand_qty, safety_stock, current_stock, incoming_qty) -> int:
    """
    필요수량 = MAX(0, (소요량 + 안전재고) - (현재고 + 입고예정수량))
    """
    demand_qty = _num(demand_qty)
    safety_stock = _num(safety_stock)
    current_stock = _num(current_stock)
    incoming_qty = _num(incoming_qty)
    return max(0, math.ceil((demand_qty + safety_stock) - (current_stock + incoming_qty)))


def compute_effective_production_qty(target_qty, safety_stock, current_stock) -> int:
    """
    실제 생산 필요량 (자재 소요량 계산의 기준이 되는 수량)
    = MAX(0, (생산 목표 수량 + 안전재고) - 현재고)

    예전 버전의 버그: '생산 지시량'을 그대로 자재 계산에 썼기 때문에
    안전재고 기준(개월)을 바꿔도 결과가 전혀 달라지지 않았습니다.
    지금은 안전재고·현재고까지 반영된 '실제 생산 필요량'을 자재 계산의
    기준으로 사용하도록 수정했습니다.
    """
    target_qty = _num(target_qty)
    safety_stock = _num(safety_stock)
    current_stock = _num(current_stock)
    return max(0, math.ceil((target_qty + safety_stock) - current_stock))


def compute_order_quantity(required_qty, moq) -> int:
    """발주수량: 필요수량을 MOQ의 배수로 올림 (MOQ 없으면 필요수량 그대로)"""
    required_qty = _num(required_qty)
    if required_qty <= 0:
        return 0
    moq = _num(moq)
    if moq > 0:
        return math.ceil(required_qty / moq) * int(moq)
    return math.ceil(required_qty)


def compute_order_by_date(production_start_date: Optional[date], lead_time_days) -> Optional[date]:
    """
    발주권장일(= 예전 '발주예상일'의 의미 수정판)
    ------------------------------------------------
    "생산을 위해 언제 발주하면 좋을지"를 나타내는 날짜입니다.
    = 생산 시작 예정일 - 리드타임(일)

    (기존에는 '오늘+리드타임'으로 계산해서 '언제 자재가 들어올지'에 가까운
    의미였는데, 실제로 필요한 건 '늦어도 언제까지는 발주를 넣어야
    생산 시작일에 자재가 맞춰 들어오는지'이므로 계산식을 이렇게 수정했습니다.)
    """
    if production_start_date is None:
        return None
    if isinstance(production_start_date, pd.Timestamp):
        production_start_date = production_start_date.date()
    lead_time_days = int(_num(lead_time_days))
    return production_start_date - timedelta(days=lead_time_days)


def days_until_order_deadline(order_by_date: Optional[date], today: Optional[date] = None) -> Optional[int]:
    """발주권장일까지 남은 일수 (음수면 이미 늦은 것)"""
    if order_by_date is None:
        return None
    today = today or date.today()
    if isinstance(today, pd.Timestamp):
        today = today.date()
    return (order_by_date - today).days
