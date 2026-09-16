"""
앱 진입점 (라우터)
--------------------
실제 화면 내용은 각 페이지 파일에 있습니다.
  - dashboard.py : 발주 계획 대시보드 (기존 app.py 내용 전체)
  - guidance.py  : 데이터 갱신 가이드 (엑셀 업로드 없이도 접속 가능)

streamlit run app.py 로 실행하는 건 이전과 동일합니다.
"""

import streamlit as st

st.set_page_config(page_title="발주 계획 대시보드", layout="wide")

guidance_page = st.Page("guidance.py", title="데이터 갱신 가이드", icon="📘")
dashboard_page = st.Page("dashboard.py", title="발주 계획 대시보드", icon="📦", default=True)

pg = st.navigation([guidance_page, dashboard_page])
pg.run()
