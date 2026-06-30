"""
데이터 취합 시스템 - 메인 진입점.

이 파일은 로그인 화면과 탭 구성만 담당한다.
각 탭의 실제 로직은 modules/ 폴더의 각 파일에 분리되어 있다.
"""
import sys
import os
import streamlit as st

# [마법의 경로 설정] 파이썬이 modules 방을 정확히 찾도록 강제 지정
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

# 쪼개놓은 각 탭의 모듈들을 불러옵니다.
from modules import tab1_simple_sum
from modules import tab2_master_template
from modules import tab3_realestate_monthly
from modules import tab4_hwpx_merge
from modules import tab5_jijeok_adjustment
from modules import tab6_jijeok_opinion
from modules import tab7_hwpx_ai  # 🌟 주무관님이 업그레이드하신 7번 탭!

st.set_page_config(layout="wide")

# 🔒 비밀번호 로그인 로직
if "a" not in st.session_state:
    st.session_state.a = False
if not st.session_state.a:
    p = st.text_input("비밀번호", type="password")
    if st.button("입장"):
        if p == "7777":
            st.session_state.a = True
            st.rerun()
    st.stop()

st.title("📊 데이터 취합 및 AI 자동화 시스템")

# 🌟 1번부터 7번까지 탭 메뉴판 만들기
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "① 단순 합산",
    "② 중개사 분기",
    "③ 실거래 월보",
    "④ 한글(HWPX) 병합",
    "⑤ 지적재조사 조정금",
    "⑥ 의견접수·이의신청",
    "⑦ 행사 AI & PPT 생성"
])

# 🌟 각 탭 메뉴를 클릭했을 때 해당 모듈의 화면을 보여주도록 연결
with tab1:
    tab1_simple_sum.render()

with tab2:
    tab2_master_template.render()

with tab3:
    tab3_realestate_monthly.render()

with tab4:
    tab4_hwpx_merge.render()

with tab5:
    tab5_jijeok_adjustment.render()

with tab6:
    tab6_jijeok_opinion.render()

with tab7:
    tab7_hwpx_ai.render()
