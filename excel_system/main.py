"""
모아주이소, AI — 데이터 취합 및 AI 업무자동화 시스템 (메인 진입점)
2026-10 : 🤖 AI 서무비서(tab17_seomu) 맨 앞 탭으로 추가 + 강조 표시, 제목 변경
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
from modules import tab17_seomu               # 🤖 AI 서무비서 (맨 앞 탭)
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
from modules import tab14_newsclip            # 🌟 14번 탭(뉴스 클리핑)
from modules import tab15_compare             # 🌟 15번 탭(분기·연차 비교)
from modules import tab16_pressrelease        # 🌟 16번 탭(보도자료 AI 생성)
from modules import tab18_landpermit          # 🌟 메뉴 ⑰ 토지거래허가 분기보고 취합

st.set_page_config(page_title="모아주이소, AI", page_icon="📊", layout="wide")
# ==========================================================================
# 🔒 개인정보 포함 탭 보호 : 실명법 과징금·개발부담금 자료에는 개인정보가 들어 있어
#    외부 클라우드(웹) 접속에서는 열지 않고, 내 PC에서 실행할 때만 사용
#    - 내 PC 실행 : streamlit run main.py → http://localhost:8501 로 열면 자동 허용
#    - 다른 탭도 막으려면 그 탭의 render() 호출을 _pii_tab(함수, "자료명") 으로 감싸면 됨
# ==========================================================================
def _is_local_run():
    """내 PC(localhost)에서 실행 중이면 True, 외부 웹(클라우드) 접속이면 False"""
    if os.environ.get("MOAJUISO_LOCAL") == "1":
        return True
    try:
        host = (st.context.headers.get("host") or "").split(":")[0].strip().lower()
    except Exception:
        host = ""
    return host in ("localhost", "127.0.0.1")


def _pii_tab(render_fn, 자료명):
    """개인정보 포함 탭 : 내 PC에서만 실행, 웹에서는 안내문만 표시"""
    if _is_local_run():
        st.caption("🔒 개인정보 포함 자료 — 내 PC에서 실행 중이라 자료가 외부로 나가지 않습니다.")
        render_fn()
        return
    st.warning(f"🔒 **{자료명}**에는 개인정보가 포함되어 있어 웹(외부 서버)에서는 사용할 수 없습니다.")
    st.markdown(
        "- 이 기능은 **담당자 PC에서 실행하는 버전**으로만 사용합니다.\n"
        "- 내 PC에서 실행하면 파일이 PC 밖으로 나가지 않습니다.\n"
        "- 사용이 필요하면 도 토지정보과 담당자에게 문의해 주세요."
    )



APP_TITLE = "📊 모아주이소, AI"
APP_SUBTITLE = "AI 서무비서 · 데이터 취합 · 문서 자동생성"

# 🔒 비밀번호 로그인 로직
if "a" not in st.session_state:
    st.session_state.a = False
if not st.session_state.a:
    st.title(APP_TITLE)
    st.caption(APP_SUBTITLE)
    p = st.text_input("비밀번호", type="password")
    if st.button("입장"):
        if p == "7777":
            st.session_state.a = True
            st.rerun()
    st.stop()

st.title(APP_TITLE)
st.caption(APP_SUBTITLE)

# 탭이 많아 한 줄에 안 들어가므로, 넘치면 가로 스크롤 대신 다음 줄로 흐르도록 한다.
# 맨 앞 탭(🤖 AI 서무비서)은 남색 바탕·흰 굵은 글씨로 강조한다.
st.markdown("""
<style>
div[data-testid="stTabs"] div[role="tablist"] {
    flex-wrap: wrap;
    row-gap: 4px;
}
div[data-testid="stTabs"] [role="tablist"] > [role="tab"]:first-child {
    background: #1f3864;
    border-radius: 8px 8px 0 0;
    padding: 0 16px;
    margin-right: 6px;
}
div[data-testid="stTabs"] [role="tablist"] > [role="tab"]:first-child * {
    color: #ffffff !important;
    font-weight: 700 !important;
}
div[data-testid="stTabs"] [role="tablist"] > [role="tab"]:first-child:hover {
    background: #2e74b5;
}
</style>
""", unsafe_allow_html=True)

# 🌟 맨 앞 AI 서무비서 + 1번부터 16번까지 탭 메뉴판 만들기
(tab_ai, tab1, tab2, tab3, tab4, tab5, tab6, tab7,
 tab8, tab9, tab10, tab11, tab12, tab13, tab14, tab15, tab16, tab17) = st.tabs([
    "🤖 AI 서무비서",        # 🤖 맨 앞 · 강조 탭 (접속 시 첫 화면)
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
    "⑬ 개발부담금 실적보고",  # 🌟 13번 탭 메뉴
    "⑭ 뉴스 클리핑",         # 🌟 14번 탭 메뉴
    "⑮ 분기·연차 비교",      # 🌟 15번 탭 메뉴
    "⑯ 보도자료 AI 생성",    # 🌟 16번 탭 메뉴
    "⑰ 토지거래허가 분기",   # 🌟 17번 탭 메뉴 (파일: tab18_landpermit.py)
])
# 🌟 각 탭 연결
with tab_ai:
    tab17_seomu.render()       # 🤖 AI 서무비서
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
    _pii_tab(tab8_collection_adjustment.render, "실명법(과징금) 현황")
with tab9:
    tab9_attendee_merge.render()
with tab10:
    tab10_spellcheck.render()
with tab11:
    tab11_crosscheck.render()
with tab12:
    _pii_tab(tab12_devcharge.render, "개발부담금 징수수수료 자료")   # 🔒 12번 탭
with tab13:
    _pii_tab(tab13_devreport.render, "개발부담금 부과·징수 실적")   # 🔒 13번 탭
with tab14:
    tab14_newsclip.render()    # 🌟 14번 탭 실행
with tab15:
    tab15_compare.render()     # 🌟 15번 탭 실행
with tab16:
    tab16_pressrelease.render_tab16()   # 🌟 16번 탭 실행
with tab17:
    tab18_landpermit.render()           # 🌟 ⑰ 토지거래허가 분기보고 취합
