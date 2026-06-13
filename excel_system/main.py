import streamlit as st
import io
import openpyxl

st.set_page_config(layout="wide")

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    pwd = st.text_input("비밀번호", type="password")
    if st.button("입장"):
        if pwd == "7777":
            st.session_state.auth = True
            st.rerun()
    st.stop()

st.title("📊 데이터 취합 시스템")
files = st.file_uploader("파일 업로드", type=["xlsx"], accept_multiple_files=True)

if files and st.button("🚀 취합 시작"):
    with st.spinner("처리 중..."):
        try:
            base_wb = openpyxl.load_workbook(io.BytesIO(files[0].read()))
            for f in files[1:]:
                wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
                for s in base_wb.sheetnames:
                    if s in wb.sheetnames:
                        ws_b = base_wb[s]
                        ws_t = wb[s]
                        for r in range(1, ws_b.max_row + 1):
                            for c in range(1, ws_b.max_column + 1):
                                v1 = ws_b.cell(r, c).value
                                v2 = ws_t.cell(r, c).value
                                if isinstance(v1, (int, float)) or isinstance(v2, (int, float)):
                                    ws_b.cell(r, c).value = (v1 if isinstance(v1, (int, float)) else 0) + (v2 if isinstance(v2, (int, float)) else 0)
            
            out = io.BytesIO()
            base_wb.save(out)
            st.download_button("📥 결과물 다운로드", out.getvalue(), "result.xlsx")
        except Exception as e:
            st.error(f"오류: {e}")