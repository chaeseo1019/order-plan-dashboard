"""
데이터 갱신 가이드
--------------------
제품_BOM / 최신재고현황 / 생산입고현황 3개 엑셀을 ERP(이카운트)에서
어떻게 뽑아서 대시보드에 업로드할 파일로 만드는지 안내하는 페이지입니다.

이 페이지는 엑셀 파일 업로드 여부와 관계없이 항상 접속할 수 있습니다
(dashboard.py의 세션 상태를 전혀 참조하지 않습니다).
"""

from pathlib import Path

import streamlit as st

IMG_DIR = Path(__file__).parent / "assets" / "guide_images"

st.title("📘 데이터 갱신 가이드")
st.caption("발주 계획 대시보드에 업로드할 3개 엑셀 시트를 ERP(이카운트)에서 뽑는 방법입니다.")


def show_step(label: str, text: str, img_name: str, warn: bool = False) -> None:
    if warn:
        st.warning(f"**{label}**  {text}")
    else:
        st.markdown(f"**{label}**  {text}")

    img_path = IMG_DIR / img_name
    if img_path.exists():
        st.image(str(img_path), width="stretch")
    else:
        st.info(f"스크린샷 준비중입니다. ({img_name} 파일을 assets/guide_images 폴더에 넣어주세요)")
    st.divider()


tab_bom, tab_inventory, tab_production = st.tabs(
    ["1️⃣ 제품_BOM", "2️⃣ 최신재고현황", "3️⃣ 생산입고현황"]
)

# ============================================================
# 1. 제품_BOM
# ============================================================
with tab_bom:
    st.subheader("제품_BOM")
    c1, c2 = st.columns(2)
    c1.metric("갱신 시점", "신제품/원재료 변경 시")
    c2.metric("ERP 경로", "재고1 > 생산/외주 > BOM 소요량 조회")
    st.divider()

    show_step("1-1.", "재고1 > 생산/외주 > BOM 소요량 조회 화면으로 이동", "1_1.png")
    show_step("1-2.", "찾는 제품명을 입력한 후 'BOM등록' 열의 조회 버튼 클릭 → 기본버전 선택", "1_2.png")
    show_step("1-3.", "하단에 있는 Excel 버튼 클릭하여 엑셀 파일 다운로드", "1_3.png")
    show_step(
        "1-4.",
        "다운로드한 파일에서 품목코드 · 품목명 · 수량 3개 컬럼만 복사해서 붙여넣기 "
        "(제품코드와 제품명은 전체 행에 동일한 값으로 채우기)",
        "1_4.png",
    )
    show_step("1-5.", "최종적으로 시트에 들어가야 하는 구조는 아래와 같음. (제품코드와 제품명은 전부 동일하게 채워주기)", "1_5.png")

# ============================================================
# 2. 최신재고현황
# ============================================================
with tab_inventory:
    st.subheader("최신재고현황")
    c1, c2 = st.columns(2)
    c1.metric("갱신 시점", "발주 계획 세울 때마다")
    c2.metric("ERP 경로", "재고1 > 출력물 > 재고현황")
    st.divider()

    show_step("2-1.", "재고1 > 출력물 > 재고현황 화면으로 이동", "2_1.png")
    show_step("2-2.", "'품목' 버튼 클릭 → 상품, 무형상품 체크 해제", "2_2.png")
    show_step("2-3.", "하단 Excel 버튼 클릭 → 엑셀 파일로 다운로드", "2_3.png")
    show_step("2-4.", "품목코드, 품목구분, 품목명[규격], 재고수량 4개 컬럼만 복사해서 붙여넣기", "2_4.png")

# ============================================================
# 3. 생산입고현황
# ============================================================
with tab_production:
    st.subheader("생산입고현황")
    c1, c2 = st.columns(2)
    c1.metric("갱신 시점", "발주 계획 세울 때마다")
    c2.metric("ERP 경로", "재고1 > 생산/외주 > 생산입고 > 생산입고현황")
    st.divider()

    show_step(
        "3-1.",
        "재고1 > 생산/외주 > 생산입고 > 생산입고현황 화면으로 이동 — "
        "조회기간은 2025년 1월 1일부터로 설정(변경 가능, 과거 생산이력 전체 확인용), "
        "'품목' 버튼 클릭 → 상품·무형상품 체크 해제",
        "3_1.png",
    )
    show_step(
        "3-2-1.",
        "'적용양식'에서 반드시 '품목코드 추가' 클릭 — 누락 시 추후 프로그램 오류 발생 "
        "(항목이 안 보인다면 오주현 차장님께 문의 후 추가하기) \n" 
        "\n 3-2-2. '정렬/소계기준' → 설정 클릭",
        "3_2.png",
        warn=True,
    )
    show_step("3-3.", "정렬/소계/합계 설정 창에서 1번 항목 클릭 → 선택삭제 → 적용", "3_3.png")
    show_step("3-4.", "하단 Excel 버튼 클릭 → 엑셀 파일로 다운로드", "3_4.png")
