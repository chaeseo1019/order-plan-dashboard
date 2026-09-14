"""
제품별 발주 계획 대시보드
--------------------------
월별재고수불부 + 최신재고현황 + 제품_BOM 시트가 포함된 엑셀 파일을 업로드하면,
제품을 선택해 생산 계획을 입력하고, 자재별 발주 조건(로스율/주문업체/국내외/
리드타임/입고예정수량/MOQ)을 입력해서 필요수량·발주수량·발주권장일을 계산합니다.

화면 구성 원칙:
  - "✏️ 직접 입력"과 "🔒 자동으로 채워지는 값"을 항상 별도 영역/표로 분리해서 보여줍니다.
  - 입력값은 st.form 안에 모아두고, [발주 계획 계산하기] 버튼을 눌러야만 재계산됩니다.
"""

from datetime import date, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st

from calculations import (
    compute_three_month_avg_outbound,
    compute_material_requirement,
    compute_safety_stock,
    compute_required_order_qty,
    compute_effective_production_qty,
    compute_order_quantity,
    compute_order_by_date,
    days_until_order_deadline,
)

st.set_page_config(page_title="발주 계획 대시보드", layout="wide")

REQUIRED_BOM_COLS = ["제품코드", "제품명", "품목코드", "품목명", "수량"]
REQUIRED_INV_COLS = ["품목코드", "품목구분", "품목명[규격]", "재고수량"]
REQUIRED_FLOW_COLS = ["기준날짜", "품목코드", "품목명", "입고", "출고", "재고"]
FLOW_SHEET_CANDIDATES = ["월별재고수불부_병합", "월별재고수불부_선별", "월별재고수불부"]


# ============================================================
# 데이터 로드 & 검증
# ============================================================
@st.cache_data(show_spinner="엑셀 파일을 읽는 중입니다...")
def load_excel(file_bytes: bytes) -> dict:
    xls = pd.ExcelFile(BytesIO(file_bytes))
    return {name: pd.read_excel(xls, sheet_name=name) for name in xls.sheet_names}


def find_flow_sheet(sheets: dict) -> str | None:
    for name in FLOW_SHEET_CANDIDATES:
        if name in sheets:
            return name
    for name in sheets:
        if "수불부" in name:
            return name
    return None


def validate_columns(df: pd.DataFrame, required_cols: list[str], sheet_name: str) -> None:
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(
            f"'{sheet_name}' 시트에 필요한 컬럼이 없습니다: {missing}\n\n"
            f"현재 컬럼: {list(df.columns)}"
        )
        st.stop()


def normalize_code(series: pd.Series) -> pd.Series:
    """품목코드/제품코드를 문자열로 통일 (숫자코드/문자코드 혼재 대응)"""
    return series.astype(str).str.strip()


# ============================================================
# UI - 파일 업로드
# ============================================================
st.title("📦 제품별 발주 계획 대시보드")

with st.expander("ℹ️ 사용법 / 엑셀 파일 형식 안내 (클릭하세요)", expanded=False):
    st.markdown(
        """
        업로드하는 엑셀 파일에는 아래 3개 시트가 **정확한 시트명·컬럼명으로** 포함되어야 합니다.

        정확한 발주 계획을 위해 엑셀 시트의 데이터는 최신화한 후 업로드해주세요.

        | 시트명 | 필요 컬럼 |
        |---|---|
        | `제품_BOM` | 제품코드, 제품명, 품목코드, 품목명, 수량 |
        | `최신재고현황` | 품목코드, 품목구분, 품목명[규격], 재고수량 |
        | `월별재고수불부` | 기준날짜(YYYY-MM), 품목코드, 품목명, 입고, 출고, 재고 |

        **화면 보는 법**
        - ✏️ 표시가 있는 영역/칸은 직접 입력해야 하는 값입니다.
        - 🔒 표시가 있는 영역/칸은 엑셀에서 자동으로 조회·계산된 값이라 수정할 수 없습니다.
        - 값을 다 입력한 뒤 **[📊 발주 계획 계산하기]** 버튼을 눌러야 결과가 계산/갱신됩니다.
        """
    )

uploaded = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if not uploaded:
    st.info("엑셀 파일을 업로드하면 시작됩니다.")
    st.stop()

sheets = load_excel(uploaded.getvalue())

if "제품_BOM" not in sheets:
    st.error("'제품_BOM' 시트를 찾을 수 없습니다. 엑셀 파일을 확인해주세요.")
    st.stop()
if "최신재고현황" not in sheets:
    st.error("'최신재고현황' 시트를 찾을 수 없습니다. 엑셀 파일을 확인해주세요.")
    st.stop()

flow_sheet_name = find_flow_sheet(sheets)
if flow_sheet_name is None:
    st.error("'월별재고수불부' 관련 시트를 찾을 수 없습니다. 엑셀 파일을 확인해주세요.")
    st.stop()

bom_df = sheets["제품_BOM"].copy()
inv_df = sheets["최신재고현황"].copy()
flow_df = sheets[flow_sheet_name].copy()

validate_columns(bom_df, REQUIRED_BOM_COLS, "제품_BOM")
validate_columns(inv_df, REQUIRED_INV_COLS, "최신재고현황")
validate_columns(flow_df, REQUIRED_FLOW_COLS, flow_sheet_name)

bom_df["제품코드"] = normalize_code(bom_df["제품코드"])
bom_df["품목코드"] = normalize_code(bom_df["품목코드"])
inv_df["품목코드"] = normalize_code(inv_df["품목코드"])
flow_df["품목코드"] = normalize_code(flow_df["품목코드"])
flow_df["기준날짜"] = flow_df["기준날짜"].astype(str)

inv_lookup = inv_df.set_index("품목코드")

st.success(
    f"업로드 완료 — 제품_BOM {bom_df['제품코드'].nunique()}개 제품 / "
    f"최신재고현황 {len(inv_df)}개 품목 / {flow_sheet_name} {len(flow_df)}행"
)

# ============================================================
# 1. 제품 선택 (화면 가운데 툴박스로 배치 — 즉시 반영, 계산 버튼과 무관)
# ============================================================
product_options = bom_df[["제품코드", "제품명"]].drop_duplicates().reset_index(drop=True)


def _display_name(row) -> str:
    inv_name = inv_lookup["품목명[규격]"].get(row["제품코드"])
    name = inv_name if isinstance(inv_name, str) else row["제품명"]
    return f'{row["제품코드"]} | {name}'


product_options["표시명"] = product_options.apply(_display_name, axis=1)

with st.container(border=True):
    st.markdown("**1️⃣ 제품 선택**")
    selected_label = st.selectbox(
        "제품", product_options["표시명"].tolist(), label_visibility="collapsed"
    )

selected_code = product_options.loc[product_options["표시명"] == selected_label, "제품코드"].iloc[0]
product_name = selected_label.split(" | ", 1)[1]

materials = bom_df[bom_df["제품코드"] == selected_code].copy()
materials = materials.rename(columns={"수량": "개당실제투입량"})
materials["구분"] = materials["품목코드"].map(lambda c: inv_lookup["품목구분"].get(c, "-"))
materials["현재고"] = materials["품목코드"].map(lambda c: float(inv_lookup["재고수량"].get(c, 0) or 0))
materials = materials.reset_index(drop=True)

current_stock_product = float(inv_lookup["재고수량"].get(selected_code, 0) or 0)

st.subheader(f"📌 {selected_code} | {product_name}")

# ------------------------------------------------------------
# 🔒 자동 조회된 정보 (계산 버튼 없이도 바로 보이는 참고용 값)
# ------------------------------------------------------------
with st.container(border=True):
    st.markdown("**🔒 자동으로 조회된 정보 (수정 불가)**")
    c1, c2 = st.columns(2)
    c1.metric("제품 현재고", f"{current_stock_product:,.0f}")
    c2.metric("BOM 자재 수", f"{len(materials)}개")

    st.markdown(
        '<span style="font-weight:700;">자재별 기본 정보</span>'
        '&nbsp;&nbsp;'
        '<span style="font-size:0.875rem; color:#808495;">'
        '(품목코드/품목명/구분/개당실제투입량/현재고는 엑셀에서 자동으로 가져온 값입니다)'
        '</span>',
        unsafe_allow_html=True,
    )

    st.dataframe(
        materials[["품목코드", "품목명", "개당실제투입량", "현재고", "구분"]],
        hide_index=True,
        width="stretch",
    )

if materials.empty:
    st.warning("이 제품에 대한 BOM 자재 정보가 없습니다.")
    st.stop()

# ============================================================
# 2. 입력 폼 (여기 안의 값은 버튼을 눌러야 계산에 반영됨)
#    -> 표(data_editor) 대신, 자재별 카드(expander) + 개별 입력 위젯 방식.
#       카드를 펼치면 안에 있는 건 전부 "입력해야 하는 값"이라 헷갈릴 일이 없습니다.
# ============================================================
prev_key = f"material_inputs_{selected_code}"
prev_df = st.session_state.get(prev_key)


def _prev_value(item_code: str, col: str, default):
    if prev_df is None:
        return default
    row = prev_df[prev_df["품목코드"] == item_code]
    if row.empty:
        return default
    return row.iloc[0][col]


with st.form("calc_form"):
    st.markdown("### ✏️ 직접 입력")

    st.markdown("**2️⃣ 생산 계획**")
    fc1, fc2, fc3, fc4 = st.columns(4)
    production_qty = fc1.number_input("생산 목표 수량(수요)", min_value=0, value=500, step=50)
    production_start = fc2.date_input("생산 시작 예정일", value=date.today() + timedelta(days=14))
    asof_date = fc3.date_input("기준일자 (보통 오늘)", value=date.today())
    safety_months = fc4.number_input(
        "안전재고 기준 (개월)", min_value=0.0, value=1.0, step=0.5,
        help="제품(완제품)에 대해서만 적용됩니다. 이 값에 따라 아래 자재 소요량도 함께 달라집니다.",
    )

    st.markdown("**3️⃣ 자재별 발주 조건** — 자재를 펼쳐서 입력하세요. (카드 안 = 전부 입력값)")

    input_rows = []
    for _, m in materials.iterrows():
        item_code = m["품목코드"]
        item_name = m["품목명"]
        key_prefix = f"{selected_code}_{item_code}"

        with st.expander(
            f"🔧 {item_code} | {item_name}   "
            f"(🔒 참고: 현재고 {m['현재고']:,.0f} · 개당실제투입량 {m['개당실제투입량']:.4f} · 구분 {m['구분']})"
        ):
            ic1, ic2, ic3 = st.columns(3)
            loss_pct = ic1.number_input(
                "예상로스율(%)", min_value=0, max_value=100, step=1,
                value=int(round(_prev_value(item_code, "예상로스율", 0.0) * 100)),
                key=f"loss_{key_prefix}",
            )
            vendor = ic2.text_input(
                "주문업체", value=str(_prev_value(item_code, "주문업체", "")),
                key=f"vendor_{key_prefix}",
            )
            prev_origin = _prev_value(item_code, "국내/외", "국내")
            origin = ic3.selectbox(
                "국내/외", ["국내", "해외"],
                index=0 if prev_origin != "해외" else 1,
                key=f"origin_{key_prefix}",
            )

            ic4, ic5, ic6 = st.columns(3)
            lead_time = ic4.number_input(
                "리드타임(일)", min_value=0, step=1,
                value=int(_prev_value(item_code, "리드타임(일)", 0)),
                key=f"lead_{key_prefix}",
            )
            incoming = ic5.number_input(
                "입고예정수량", min_value=0, step=1,
                value=int(_prev_value(item_code, "입고예정수량", 0)),
                key=f"incoming_{key_prefix}",
            )
            moq = ic6.number_input(
                "MOQ", min_value=0, step=1,
                value=int(_prev_value(item_code, "MOQ", 0)),
                help="0이면 MOQ 제약 없이 필요수량 그대로 발주",
                key=f"moq_{key_prefix}",
            )

        input_rows.append(
            {
                "품목코드": item_code,
                "품목명": item_name,
                "예상로스율": loss_pct / 100,
                "주문업체": vendor,
                "국내/외": origin,
                "리드타임(일)": lead_time,
                "입고예정수량": incoming,
                "MOQ": moq,
            }
        )

    edited_inputs = pd.DataFrame(input_rows)
    submitted = st.form_submit_button("📊 발주 계획 계산하기", type="primary", use_container_width=True)

if submitted:
    st.session_state[prev_key] = edited_inputs
    st.session_state[f"last_inputs_{selected_code}"] = {
        "production_qty": production_qty,
        "production_start": production_start,
        "asof_date": asof_date,
        "safety_months": safety_months,
        "edited": edited_inputs.copy(),
    }

# ============================================================
# 3. 계산 결과 (마지막으로 '계산하기'를 누른 값 기준으로만 표시)
# ============================================================
saved = st.session_state.get(f"last_inputs_{selected_code}")

st.divider()
st.subheader("🧮 계산 결과")

if saved is None:
    st.info("위에서 값을 입력하고 **[📊 발주 계획 계산하기]** 버튼을 눌러주세요.")
    st.stop()

production_qty = saved["production_qty"]
production_start = saved["production_start"]
asof_date = saved["asof_date"]
safety_months = saved["safety_months"]
edited = saved["edited"]
lead_time_max = edited["리드타임(일)"].max()

avg_outbound_3m = compute_three_month_avg_outbound(flow_df, selected_code, asof_date, months=3)
safety_stock_product = compute_safety_stock(avg_outbound_3m, safety_months)
effective_qty = compute_effective_production_qty(production_qty, safety_stock_product, current_stock_product)

with st.container(border=True):
    st.markdown("**🔒 제품 레벨 계산 결과**")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("🏭 실제 생산 필요량", f"{effective_qty:,.0f}")
    c2.metric("제품 현재고", f"{current_stock_product:,.0f}")
    c3.metric("안전재고", f"{safety_stock_product:,.0f}")
    c4.metric("3개월 평균 출고량", f"{avg_outbound_3m:,.1f}")
    c5.metric("자재 최대 리드타임", f"{lead_time_max:,.0f}")
    st.caption(
        "실제 생산 필요량 = MAX(0, (생산 목표 수량 + 안전재고) − 현재고) — "
        "**아래 자재 소요량은 이 수량을 기준으로 계산됩니다.**"
    )

# 자재별 계산 ---------------------------------------------------
results = materials[["품목코드", "품목명", "구분", "개당실제투입량", "현재고"]].merge(
    edited, on=["품목코드", "품목명"], how="left"
)

results["소요량"] = results.apply(
    lambda r: compute_material_requirement(effective_qty, r["개당실제투입량"], r["예상로스율"]), axis=1
)
results["필요수량"] = results.apply(
    lambda r: compute_required_order_qty(r["소요량"], 0, r["현재고"], r["입고예정수량"]), axis=1
)
results["발주수량"] = results.apply(
    lambda r: compute_order_quantity(r["필요수량"], r["MOQ"]), axis=1
)
results["발주권장일"] = results.apply(
    lambda r: compute_order_by_date(production_start, r["리드타임(일)"]) if r["발주수량"] > 0 else None,
    axis=1,
)
# D-day는 화면 표에는 더 이상 표시하지 않지만, 긴급도에 따른 행 강조(빨강/노랑) 색상 계산에는 계속 사용합니다.
results["D-day"] = results["발주권장일"].apply(
    lambda d: days_until_order_deadline(d, asof_date) if pd.notna(d) else None
)

st.caption(
    "**발주권장일** = 생산 시작 예정일 − 리드타임(일)  →  \"생산 시작일에 자재가 맞춰 들어오려면 "
    "늦어도 언제까지는 발주를 넣어야 하는가\"를 의미합니다. "
    "(🔴 빨간색: 이미 발주 시점이 지남 / 🟡 노란색: 3일 이내로 임박)"
)

display_cols = [
    "품목코드", "품목명", "구분", "소요량", "현재고", "입고예정수량",
    "필요수량", "MOQ", "발주수량", "리드타임(일)", "발주권장일",
    "주문업체", "국내/외",
]


def _highlight_urgent(row):
    original = results.loc[row.name]
    if pd.isna(original["D-day"]):
        return [""] * len(row)
    if original["D-day"] < 0:
        return ["background-color: #F8D7DA"] * len(row)
    if original["D-day"] <= 3:
        return ["background-color: #FFF3CD"] * len(row)
    return [""] * len(row)


NO_ORDER_TEXT = "발주 필요없음"

display_df = results[display_cols].copy()
display_df["발주권장일"] = results["발주권장일"].apply(
    lambda d: NO_ORDER_TEXT if pd.isna(d) else d.isoformat()
)

st.dataframe(
    display_df.style.apply(_highlight_urgent, axis=1).format(
        {
            "소요량": "{:,.0f}",
            "현재고": "{:,.0f}",
            "입고예정수량": "{:,.0f}",
            "필요수량": "{:,.0f}",
            "발주수량": "{:,.0f}",
        }
    ),
    width="stretch",
    hide_index=True,
)

need_order = results[results["발주수량"] > 0]
if not need_order.empty:
    earliest = need_order["발주권장일"].min()
    overdue = need_order[need_order["D-day"] < 0]
    msg = f"⚠️ {len(need_order)}개 자재에 대해 발주가 필요합니다. 가장 이른 발주권장일은 **{earliest}** 입니다.\n"
    if not overdue.empty:
        overdue_list = "".join(
            f"\n  - **{row['품목코드']} | {row['품목명']}** "
            f"(발주권장일 {row['발주권장일']}, {abs(row['D-day']):.0f}일 지남)"
            for _, row in overdue.iterrows()
        )
        msg += f"  \n🔴 이미 발주권장일이 지난 자재 {len(overdue)}건:{overdue_list}"
    st.warning(msg)
else:
    st.success("현재 재고(+입고예정)로 충분합니다. 추가 발주가 필요하지 않습니다.")

# ============================================================
# 4. 결과 다운로드
# ============================================================
st.subheader("⬇️ 결과 다운로드")
csv = results[display_cols].to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "CSV로 다운로드",
    csv,
    file_name=f"발주계획_{selected_code}_{asof_date.isoformat()}.csv",
    mime="text/csv",
)