import streamlit as st
import io
import openpyxl
from openpyxl.cell.cell import MergedCell

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
files = st.file_uploader("파일 업로드", type=["xlsx"], accept_multiple_files=True)

if files and st.button("🚀 취합 시작"):
    try:
        # 모든 파일을 data_only=True로 통일 (수식이 아닌 "계산된 값"을 읽음)
        b = openpyxl.load_workbook(io.BytesIO(files[0].read()), data_only=True)

        for f in files[1:]:
            t = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
            for s in b.sheetnames:
                if s not in t.sheetnames:
                    continue

                max_r = max(b[s].max_row, t[s].max_row)
                max_c = max(b[s].max_column, t[s].max_column)

                for r in range(1, max_r + 1):
                    for c in range(1, max_c + 1):
                        cell_b = b[s].cell(r, c)

                        # 병합 셀은 건너뜀 (쓰기 불가)
                        if isinstance(cell_b, MergedCell):
                            continue

                        v1 = cell_b.value
                        v2 = t[s].cell(r, c).value if r <= t[s].max_row and c <= t[s].max_column else None

                        is_num1 = isinstance(v1, (int, float)) and not isinstance(v1, bool)
                        is_num2 = isinstance(v2, (int, float)) and not isinstance(v2, bool)

                        if is_num1 or is_num2:
                            cell_b.value = (v1 if is_num1 else 0) + (v2 if is_num2 else 0)

        o = io.BytesIO()
        b.save(o)
        st.download_button("📥 다운로드", o.getvalue(), "result.xlsx")

    except Exception as e:
        st.error(f"오류: {e}")
        st.exception(e)  # 디버깅 시 전체 traceback 확인용
