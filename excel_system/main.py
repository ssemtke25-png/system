"""
데이터 취합 시스템 - 메인 진입점.
"""
import sys
import os

# 현재 파일이 있는 디렉토리를 가장 우선적인 경로로 설정
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir in sys.path:
    sys.path.remove(script_dir)
sys.path.insert(0, script_dir)

# 이제 이 이후에 import를 실행하면
# 다른 폴더의 영향을 받지 않고 현재 폴더의 모듈을 우선적으로 찾습니다.
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
from modules import tab7_hwpx_ai              # 🌟 7번 탭
from modules import tab8_collection_adjustment
from modules import tab9_attendee_merge       # 🌟 9번 탭
from modules import tab10_spellcheck          # 🌟 10번 탭(맞춤법 검사)
from modules import tab11_crosscheck          # 🌟 11번 탭(정합성 검산)
from modules import tab12_devcharge           # 🌟 12번 탭(개발부담금 수수료 취합)
from modules import tab13_devreport           # 🌟 13번 탭(개발부담금 실적보고 취합)

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

# 🌟 1번부터 12번까지 탭 메뉴판 만들기
(tab1, tab2, tab3, tab4, tab5, tab6, tab7,
 tab8, tab9, tab10, tab11, tab12, tab13) = st.tabs([
    "① 단순 합산",
    "② 중개사 분기",
    "③ 실거래 월보",
    "④ 한글(HWPX) 병합",
    "⑤ 지적재조사 조정금",
    "⑥ 의견접수·이의신청",
    "⑦ 행사 AI 문서생성",
    "⑧ 실명법 취합",
    "⑨ 행사명단 취합",
    "⑩ 맞춤법 검사",
    "⑪ 정합성 검산",
    "⑫ 개발부담금 수수료",   # 🌟 12번 탭 메뉴
    "⑬ 개발부담금 실적보고",  # 🌟 13번 탭 메뉴 추가
])

# 🌟 각 탭 연결
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
    tab7_hwpx_ai.render_tab7()

with tab8:
    tab8_collection_adjustment.render()

with tab9:
    tab9_attendee_merge.render()

with tab10:
    tab10_spellcheck.render()

with tab11:
    tab11_crosscheck.render()

with tab12:
    tab12_devcharge.render()   # 🌟 12번 탭 실행

with tab13:
    tab13_devreport.render()   # 🌟 13번 탭 실행
