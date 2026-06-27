"""
탭5: 지적재조사 조정금
시군구별 '지적재조사 조정금' 파일을 받아 총괄표의 각 지역 시트에
징수/지급 내역을 연도별로 취합한다. 실제 취합 로직은 tab6과 동일해서
modules/jijeok_shared.py 에 공유 함수로 정의되어 있다.
"""
import io
import streamlit as st

from modules.jijeok_shared import fill_jijeok_template


def render():
    """탭5 화면을 그린다. app.py에서 with tab5: render() 형태로 호출."""
    st.caption("시군구별 '지적재조사 조정금' 파일을 받아 총괄표의 각 지역 시트에 징수/지급 내역을 취합합니다.")
    st.info("📌 '총괄' 시트는 절대 건드리지 않으며, A열의 '연도(합계, 2012년 등)' 글자를 자동 매칭하여 데이터를 안전하게 끼워 넣습니다.")

    template_file5 = st.file_uploader("① 지적재조사(조정금) 총괄표 서식 업로드", type=["xlsx"], key="template_up5")
    region_files5 = st.file_uploader("② 시군구별 지적재조사(조정금) 파일 업로드", type=["xlsx"], accept_multiple_files=True, key="region_up5")

    if template_file5 and region_files5 and st.button("🚀 조정금 취합 시작", key="btn5"):
        try:
            template_bytes5 = io.BytesIO(template_file5.read())
            region_files_with_names5 = [(f.name, io.BytesIO(f.read())) for f in region_files5]

            result_wb5, log5, warns5 = fill_jijeok_template(template_bytes5, region_files_with_names5)

            o5 = io.BytesIO()
            result_wb5.save(o5)

            st.success("취합이 완료되었습니다.")
            st.download_button("📥 다운로드", o5.getvalue(), "지적재조사_조정금_결과.xlsx", key="dl5")

            if warns5:
                st.warning(f"⚠️ 확인이 필요한 항목 {len(warns5)}건이 발견되었습니다.")
                st.dataframe(warns5, use_container_width=True)
            else:
                st.info("특이사항 없이 정상적으로 취합되었습니다.")

            with st.expander("처리 로그 보기 (성공 내역)"):
                st.dataframe(log5, use_container_width=True)

        except Exception as e:
            st.error(f"오류: {e}")
            st.exception(e)