"""
데이터 취합 시스템 - 메인 진입점.

이 파일은 로그인 화면과 탭 구성만 담당한다.
각 탭의 실제 로직(엑셀 합산, 총괄표 채우기, 실거래 월보, 한글 파일 병합)은
modules/ 폴더의 각 파일에 분리되어 있다.

코드 구조:
    app.py                          <- 지금 이 파일 (로그인 + 탭 배치)
    modules/
        common.py                  <- 탭1~3이 공유하는 지역명/숫자 판별 유틸
        tab1_simple_sum.py         <- ① 단순 합산
        tab2_master_template.py   <- ② 총괄표 채우기 (시군구)
        tab3_realestate_monthly.py<- ③ 실거래 월보
        tab4_hwpx_merge.py        <- ④ 한글(HWPX) 파일 병합

새 탭을 추가하려면: modules/tabN_이름.py 파일을 만들고 render() 함수를 정의한 뒤,
아래에 탭을 하나 추가하고 with tabN: 안에서 render()를 호출하면 된다.
"""
import streamlit as st

from modules import tab1_simple_sum
from modules import tab2_master_template
from modules import tab3_realestate_monthly
from modules import tab4_hwpx_merge

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

tab1, tab2, tab3, tab4 = st.tabs([
    "① 단순 합산",
    "② 총괄표 채우기 (시군구)",
    "③ 실거래 월보",
    "④ 한글(HWPX) 병합",
])

with tab1:
    tab1_simple_sum.render()

with tab2:
    tab2_master_template.render()

with tab3:
    tab3_realestate_monthly.render()

with tab4:
    tab4_hwpx_merge.render()
