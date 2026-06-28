"""
데이터 취합 시스템 - 메인 진입점.
"""
import sys
import os
import streamlit as st

# 모듈 폴더 경로 강제 지정
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

from modules import tab1_simple_sum
from modules import tab2_master_template
from modules import tab3_realestate_monthly
from modules import tab4_hwpx_merge
from modules import tab5_jijeok_adjustment
from modules import tab6_jijeok_opinion
# 🌟 [추가] 7번째 탭 모듈 불러오기
from modules import tab7_hwpx_ai 

st.set_page_config(layout="wide")

if "a" not in st.session_state:
    st.session_state.a = False
if not st.session_state.a:
    p = st.text_input("비밀번호", type="password")
    if st.button("입장"):
        if p == "7777":
            st.session_state.a = True
            st.rerun()
    st.stop()

st.title("📊 데이터 취합 시스템")

# 🌟 [추가] 탭 목록에 7번째 기능 추가
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "① 단순 합산",
    "② 중개사 분기",
    "③ 실거래 월보",
    "④ 한글(HWPX) 병합",
    "⑤ 지적재조사 조정금",
    "⑥ 의견접수·이의신청",
    "⑦ 🤖 AI 문서 작성기"
])

with tab1: tab1_simple_sum.render()
with tab2: tab2_master_template.render()
with tab3: tab3_realestate_monthly.render()
with tab4: tab4_hwpx_merge.render()
with tab5: tab5_jijeok_adjustment.render()
with tab6: tab6_jijeok_opinion.render()

# 🌟 [추가] 7번째 탭 화면 렌더링
with tab7: 
    tab7_hwpx_ai.render()
