import streamlit as st
import pandas as pd
import io
import openpyxl

# [1. 기본 설정]
st.set_page_config(page_title="데이터 자동 검증 및 취합 시스템", layout="wide")

# [2. 보안 잠금장치]
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("## 🔐 시스템 보안 접속")
    pwd = st.text_input("보안 비밀번호", type="password")
    if st.button("시스템 입장"):
        if pwd == "7777":
            st.session_state.authenticated = True
            st.rerun()
    st.stop()

# [3. 메인 화면]
st.markdown("### 📊 지적재조사 만능 데이터 취합 시스템")
run_validation = st.checkbox("🔍 오류 검증 기능 켜기", value=False)
uploaded_files = st.file_uploader("엑셀 파일(.xlsx)을 드래그하세요.", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 취합 시작"):
        with st.spinner("처리 중..."):
            try:
                # 1단계: 검증
                if run_validation:
                    for f in uploaded_files:
                        wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
                        ws = wb.active
                        if (ws.cell(1, 1).value or 0) + (ws.cell(1, 2).value or 0) != (ws.cell(1, 3).value or 0):
                            st.error(f"❌ 검증 실패: {f.name}")
                            st.stop()
                
                # 2단계: 합산
                base_file = io.BytesIO(uploaded_files[0].read())
                wb_base = openpyxl.load_workbook(base_file)
                for f in uploaded_files[1:]:
                    wb_temp = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
                    for sheet in wb_base.sheetnames:
                        if sheet in wb_temp.sheetnames:
                            for r in range(1, wb_base[sheet].max_row + 1):
                                for c in range(1, wb_base[sheet].max_column + 1):
                                    v1 = wb_base[sheet].cell(r, c).value
                                    v2 = wb_temp[sheet].cell(r, c).value
                                    if isinstance(v1, (int, float)) or isinstance(v2, (int, float)):
                                        wb_base[sheet].cell(r, c).value = (v1 if isinstance(v1, (int, float)) else 0) + (v2 if isinstance(v2, (int, float)) else 0)
                
                output = io.BytesIO()
                wb_base.save(output)
                st.download_button("📥 다운로드", output.getvalue(), "결과물.xlsx")
            except Exception as e:
                st.error(f"오류: {e}")
