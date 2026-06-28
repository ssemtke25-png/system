"""
데이터 취합 시스템 - 메인 진입점.

이 파일은 로그인 화면과 탭 구성만 담당한다.
각 탭의 실제 로직은 modules/ 폴더의 각 파일에 분리되어 있다.

코드 구조:
    main.py                          <- 지금 이 파일 (로그인 + 탭 배치)
    modules/
        common.py                  <- 탭1~3이 공유하는 지역명/숫자 판별 유틸
        tab1_simple_sum.py         <- ① 단순 합산
        tab2_master_template.py   <- ② 총괄표 채우기 (시군구)
        tab3_realestate_monthly.py<- ③ 실거래 월보
        tab4_hwpx_merge.py        <- ④ 한글(HWPX) 파일 병합
        jijeok_shared.py           <- 탭5·6이 공유하는 지적재조사 계열 취합 로직
        tab5_jijeok_adjustment.py <- ⑤ 지적재조사 조정금
        tab6_jijeok_opinion.py    <- ⑥ 의견접수·이의신청
"""
import sys
import os
import streamlit as st

# 🚨 [마법의 코드] 파이썬에게 현재 main.py가 있는 폴더 위치를 강제로 알려줍니다.
# 터미널 실행 위치가 어디든 상관없이, 바로 옆에 있는 modules 폴더를 무조건 100% 찾아냅니다!
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

# 이제 파이썬이 modules 방을 정확히 찾아 들어갈 수 있습니다.
from modules import tab1_simple_sum
from modules import tab2_master_template
from modules import tab3_realestate_monthly
from modules import tab4_hwpx_merge
from modules import tab5_jijeok_adjustment
from modules import tab6_jijeok_opinion

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

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "① 단순 합산",
    "② 중개사 분기",
    "③ 실거래 월보",
    "④ 한글(HWPX) 병합",
    "⑤ 지적재조사 조정금",
    "⑥ 의견접수·이의신청",
])

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
