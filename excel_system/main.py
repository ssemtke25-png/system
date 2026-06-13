import streamlit as st
import io
import openpyxl

st.set_page_config(layout="wide")
if "a" not in st.session_state: st.session_state.a = False

if not st.session_state.a:
    p = st.text_input("비밀번호", type="password")
    if st.button("입장"):
        if p == "7777": st.session_state.a = True; st.rerun()
    st.stop()

st.title("📊 데이터 취합 시스템")
files = st.file_uploader("파일 업로드", type=["xlsx"], accept_multiple_files=True)

if files and st.button("🚀 취합 시작"):
    try:
        b = openpyxl.load_workbook(io.BytesIO(files[0].read()))
        for f in files[1:]:
            t = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
            for s in b.sheetnames:
                if s in t.sheetnames:
                    for r in range(1, b[s].max_row + 1):
                        for c in range(1, b[s].max_column + 1):
                            v1, v2 = b[s].cell(r, c).value, t[s].cell(r, c).value
                            if isinstance(v1, (int, float)) or isinstance(v2, (int, float)):
                                b[s].cell(r, c).value = (v1 if isinstance(v1, (int, float)) else 0) + (v2 if isinstance(v2, (int, float)) else 0)
        o = io.BytesIO()
        b.save(o)
        st.download_button("📥 다운로드", o.getvalue(), "result.xlsx")
    except Exception as e: st.error(f"오류: {e}")