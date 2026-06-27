"""
탭6: 지적재조사 의견접수·이의신청 현황
시군구별 '의견접수 및 이의신청 현황' 파일을 받아 총괄표의 각 지역 시트에
내역을 연도별로 취합한다. 서식 구조와 취합 방식이 탭5(조정금)와 동일해서
실제 로직은 modules/jijeok_shared.py 의 공유 함수를 그대로 재사용한다.
"""
import io
import streamlit as st

from modules.jijeok_shared import fill_jijeok_template


def render():
    """탭6 화면을 그린다. app.py에서 with tab6: render() 형태로 호출."""
    st.caption("시군구별 '의견접수 및 이의신청 현황' 파일을 받아 총괄표의 각 지역 시트에 내역을 취합합니다.")
    st.info("📌 '총괄(자동입력)' 시트는 절대 건드리지 않으며, A열의 '연도(합계, 2012년 등)' 글자를 자동 매칭하여 데이터를 안전하게 끼워 넣습니다.")

    template_file6 = st.file_uploader("① 의견접수/이의신청 총괄표 서식 업로드", type=["xlsx"], key="template_up6")
    region_files6 = st.file_uploader("② 시군구별 의견접수/이의신청 파일 업로드", type=["xlsx"], accept_multiple_files=True, key="region_up6")

    if template_file6 and region_files6 and st.button("🚀 의견접수/이의신청 취합 시작", key="btn6"):
        try:
            template_bytes6 = io.BytesIO(template_file6.read())
            region_files_with_names6 = [(f.name, io.BytesIO(f.read())) for f in region_files6]

            # 탭5와 서식 구조가 동일하므로 동일한 취합 함수를 호출한다.
            result_wb6, log6, warns6 = fill_jijeok_template(template_bytes6, region_files_with_names6)

            o6 = io.BytesIO()
            result_wb6.save(o6)

            st.success("취합이 완료되었습니다.")
            st.download_button("📥 다운로드", o6.getvalue(), "의견접수_이의신청_결과.xlsx", key="dl6")

            if warns6:
                st.warning(f"⚠️ 확인이 필요한 항목 {len(warns6)}건이 발견되었습니다.")
                st.dataframe(warns6, use_container_width=True)
            else:
                st.info("특이사항 없이 정상적으로 취합되었습니다.")

            with st.expander("처리 로그 보기 (성공 내역)"):
                st.dataframe(log6, use_container_width=True)

        except Exception as e:
            st.error(f"오류: {e}")
            st.exception(e)