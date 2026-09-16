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


def summarize_production_history(history_df: pd.DataFrame, qty_col: str = "수량", date_col: str = "일자") -> dict:
    """
    생산입고현황에서 특정 제품의 이력만 뽑아온 데이터프레임을 받아
    안전재고 수량을 정할 때 참고할 요약 통계를 계산합니다.

    반환: count(생산 횟수), total_qty(총 생산량), avg_qty(1회 평균 생산량),
          latest_date(가장 최근 생산일), latest_qty(가장 최근 생산량)
    """
    if history_df is None or history_df.empty:
        return {"count": 0, "total_qty": 0.0, "avg_qty": 0.0, "latest_date": None, "latest_qty": 0.0}

    count = len(history_df)
    total_qty = float(history_df[qty_col].sum())
    avg_qty = total_qty / count if count else 0.0
    latest_row = history_df.sort_values(date_col).iloc[-1]
    return {
        "count": count,
        "total_qty": total_qty,
        "avg_qty": avg_qty,
        "latest_date": latest_row[date_col],
        "latest_qty": float(latest_row[qty_col]),
    }


def compute_material_requirement(production_qty, unit_usage, loss_rate) -> int:
    """
    자재 소요량 = 생산 지시량 × 개당실제투입량 × (1 + 예상로스율), 올림 처리
    """
    production_qty = _num(production_qty)
    unit_usage = _num(unit_usage)
    loss_rate = _num(loss_rate)
    raw = production_qty * unit_usage * (1 + loss_rate)
    return math.ceil(raw)


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

    안전재고는 이제 사용자가 수량으로 직접 입력한 값입니다
    (예전처럼 '몇 개월치'를 곱해서 자동 산출하지 않습니다).
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
    발주권장일 = 생산 시작일 - 리드타임(일)
    "생산 시작일에 자재가 맞춰 들어오려면 늦어도 언제까지는 발주를 넣어야 하는가"
    """
    if production_start_date is None:
        return None
    if isinstance(production_start_date, pd.Timestamp):
        production_start_date = production_start_date.date()
    lead_time_days = int(_num(lead_time_days))
    return production_start_date - timedelta(days=lead_time_days)


def compute_optimal_production_start(
    production_deadline: Optional[date], production_lead_days
) -> Optional[date]:
    """
    최적 생산 시작일 = 생산마감일자 - 생산소요일

    '생산마감일자'는 완제품이 반드시 준비되어 있어야 하는 진짜 제약(납기)이고,
    '생산소요일'은 자재가 다 갖춰진 상태에서 실제 생산 공정 자체가
    며칠 걸리는지를 나타냅니다. 이 둘로 역산하면, 마감일을 맞추기 위해
    "가장 늦어도 이 날짜에는 생산을 시작해야 한다"는 시작일이 나옵니다.
    """
    if production_deadline is None:
        return None
    if isinstance(production_deadline, pd.Timestamp):
        production_deadline = production_deadline.date()
    production_lead_days = int(_num(production_lead_days))
    return production_deadline - timedelta(days=production_lead_days)


def days_until_order_deadline(order_by_date: Optional[date], today: Optional[date] = None) -> Optional[int]:
    """발주권장일까지 남은 일수 (음수면 이미 늦은 것)"""
    if order_by_date is None:
        return None
    today = today or date.today()
    if isinstance(today, pd.Timestamp):
        today = today.date()
    return (order_by_date - today).days
