import streamlit as st
import io
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

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


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def aggregate(file_bytes_list, file_names):
    """
    file_bytes_list: BytesIO 리스트 (각 파일의 원본 바이트)
    file_names: 같은 순서의 파일명 리스트
    return: (결과 워크북, 경고 리스트)
    """
    warnings = []

    value_wbs = [openpyxl.load_workbook(b, data_only=True) for b in file_bytes_list]
    formula_wbs = [openpyxl.load_workbook(b, data_only=False) for b in file_bytes_list]
    names = file_names
    n_files = len(file_bytes_list)

    base_wb = value_wbs[0]
    sheet_names = base_wb.sheetnames

    for sheet in sheet_names:
        present_idx = [i for i in range(n_files) if sheet in value_wbs[i].sheetnames]
        if len(present_idx) < n_files:
            missing_files = [names[i] for i in range(n_files) if i not in present_idx]
            warnings.append({
                "유형": "시트 누락",
                "시트": sheet,
                "셀": "-",
                "파일": ", ".join(missing_files),
                "설명": f"'{sheet}' 시트가 없어 해당 파일은 이 시트 합산에서 제외됨"
            })

        max_r = max(value_wbs[i][sheet].max_row for i in present_idx)
        max_c = max(value_wbs[i][sheet].max_column for i in present_idx)
        base_ws = base_wb[sheet]

        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                cell_addr = f"{get_column_letter(c)}{r}"
                base_cell = base_ws.cell(r, c)
                if isinstance(base_cell, MergedCell):
                    continue

                total = 0
                any_number = False
                text_files = []
                formula_flags = []

                for i in present_idx:
                    vws = value_wbs[i][sheet]
                    fws = formula_wbs[i][sheet]
                    if r > vws.max_row or c > vws.max_column:
                        continue

                    v = vws.cell(r, c).value
                    fv = fws.cell(r, c).value

                    if is_number(v):
                        total += v
                        any_number = True
                    elif v is not None and isinstance(v, str):
                        text_files.append((names[i], v))

                    formula_flags.append((names[i], is_formula(fv)))

                # 경고1: 숫자 칸에 텍스트 혼입 (다른 파일에 숫자가 있을 때만)
                if any_number and text_files:
                    warnings.append({
                        "유형": "텍스트 혼입",
                        "시트": sheet,
                        "셀": cell_addr,
                        "파일": ", ".join(fn for fn, _ in text_files),
                        "설명": "숫자가 들어가야 할 칸에 텍스트 발견 ("
                                + ", ".join(f"{fn}='{val}'" for fn, val in text_files)
                                + ") → 0으로 처리하여 합산함"
                    })

                # 경고2: 절반 이상 수식인데 일부만 수식 아님
                n_considered = len(formula_flags)
                n_formula = sum(1 for _, isf in formula_flags if isf)
                if n_considered >= 2 and n_formula >= (n_considered / 2) and n_formula < n_considered:
                    non_formula_files = [fn for fn, isf in formula_flags if not isf]
                    warnings.append({
                        "유형": "수식 누락",
                        "시트": sheet,
                        "셀": cell_addr,
                        "파일": ", ".join(non_formula_files),
                        "설명": f"전체 {n_considered}개 파일 중 {n_formula}개가 수식을 사용 중인데 "
                                + ", ".join(non_formula_files)
                                + "에는 수식이 없음 (해당 파일 값 그대로 합산됨)"
                    })

                if any_number:
                    base_cell.value = total

    return base_wb, warnings


if files and st.button("🚀 취합 시작"):
    try:
        file_bytes_list = [io.BytesIO(f.read()) for f in files]
        file_names = [f.name for f in files]

        result_wb, warns = aggregate(file_bytes_list, file_names)

        o = io.BytesIO()
        result_wb.save(o)

        st.success("취합이 완료되었습니다.")
        st.download_button("📥 다운로드", o.getvalue(), "result.xlsx")

        if warns:
            st.warning(f"⚠️ 확인이 필요한 항목 {len(warns)}건이 발견되었습니다. (합산 결과는 정상적으로 생성되었습니다)")
            st.dataframe(warns, use_container_width=True)
        else:
            st.info("특이사항 없이 정상적으로 합산되었습니다.")

    except Exception as e:
        st.error(f"오류: {e}")
        st.exception(e)
