import io
import re
from copy import copy as _copy_style

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter


REGION_ORDER_TAB8 = [
    "포항남", "포항북", "경주", "김천", "안동", "구미", "영주", "영천", "상주", "문경",
    "경산", "의성", "청송", "영양", "영덕", "청도", "고령", "성주", "칠곡",
    "예천", "봉화", "울진", "울릉"
]


def normalize_region_label(text):
    if text is None:
        return None

    s = str(text).strip()
    s = s.replace(" ", "")
    s = s.replace("시", "").replace("군", "").replace("구", "")
    s = s.replace("포항시남", "포항남").replace("포항남구", "포항남")
    s = s.replace("포항시북", "포항북").replace("포항북구", "포항북")

    for region in REGION_ORDER_TAB8:
        if region == s:
            return region
    for region in REGION_ORDER_TAB8:
        if region in s:
            return region
    return None


def extract_region_from_filename(filename):
    base = filename.rsplit(".", 1)[0]

    region = normalize_region_label(base)
    if region:
        return region

    m = re.match(r'^\d{1,3}[._\s-]*([가-힣]+)', base)
    if m:
        return normalize_region_label(m.group(1))

    return None


def sort_key_for_filename(fname):
    m = re.match(r'^(\d{1,3})[._\s-]', fname)
    if m:
        return int(m.group(1))

    region = extract_region_from_filename(fname)
    if region in REGION_ORDER_TAB8:
        return 1000 + REGION_ORDER_TAB8.index(region)

    return 9999


def get_target_sheets_from_total_wb(wb):
    if len(wb.sheetnames) < 2:
        raise ValueError("총괄표 파일은 최소 2개 시트가 필요합니다. (시트1=합산, 시트2=붙이기)")

    ws_sum = wb[wb.sheetnames[0]]
    ws_append = wb[wb.sheetnames[1]]
    return ws_sum, ws_append, wb.sheetnames[0], wb.sheetnames[1]


def get_source_sheets_from_region_wb(wb):
    if len(wb.sheetnames) < 2:
        raise ValueError("시군 파일은 최소 2개 시트가 필요합니다. (시트1=합산원본, 시트2=붙이기원본)")

    ws_sum_src = wb[wb.sheetnames[0]]
    ws_append_src = wb[wb.sheetnames[1]]
    return ws_sum_src, ws_append_src, wb.sheetnames[0], wb.sheetnames[1]


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def is_formula_cell(cell):
    return isinstance(cell.value, str) and cell.value.startswith("=")


def is_protected_target_cell(cell):
    if isinstance(cell, MergedCell):
        return True
    if is_formula_cell(cell):
        return True
    return False


def accumulate_sheet_values(base_ws, src_ws):
    max_row = min(base_ws.max_row, src_ws.max_row)
    max_col = min(base_ws.max_column, src_ws.max_column)
    updated = 0

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            base_cell = base_ws.cell(r, c)
            src_cell = src_ws.cell(r, c)

            if is_protected_target_cell(base_cell):
                continue

            src_val = src_cell.value
            base_val = base_cell.value

            if not is_number(src_val):
                continue

            if base_val in (None, ""):
                base_cell.value = src_val
                updated += 1
            elif is_number(base_val):
                base_cell.value = base_val + src_val
                updated += 1

    return updated


def find_last_data_row(ws):
    for r in range(ws.max_row, 0, -1):
        for c in range(1, ws.max_column + 1):
            if ws.cell(r, c).value not in (None, ""):
                return r
    return 0


def find_append_start_row(ws):
    last_row = find_last_data_row(ws)
    return last_row + 1


def find_source_data_start_row(ws):
    # 시트2에서 실제 표 시작행 탐지:
    # '구분'이 있는 행을 헤더로 보고, 그 다음 행부터 복사
    scan_limit = min(ws.max_row, 50)

    for r in range(1, scan_limit + 1):
        row_values = []
        for c in range(1, min(ws.max_column, 10) + 1):
            v = ws.cell(r, c).value
            if v is not None:
                row_values.append(str(v).strip())

        joined = " ".join(row_values)
        if "구분" in joined:
            return r - 3 if r - 3 >= 1 else 1

    # 못 찾으면 1행부터
    return 1


def copy_cell_style(src_cell, dst_cell):
    dst_cell.font = _copy_style(src_cell.font)
    dst_cell.fill = _copy_style(src_cell.fill)
    dst_cell.border = _copy_style(src_cell.border)
    dst_cell.alignment = _copy_style(src_cell.alignment)
    dst_cell.number_format = src_cell.number_format
    dst_cell.protection = _copy_style(src_cell.protection)


def append_sheet_rows(base_ws, src_ws):
    src_start_row = find_source_data_start_row(src_ws)
    src_last_row = find_last_data_row(src_ws)

    if src_last_row < src_start_row:
        return 0

    start_append_row = find_append_start_row(base_ws)
    max_col = min(base_ws.max_column, src_ws.max_column)
    written = 0

    for sr in range(src_start_row, src_last_row + 1):
        row_has_data = False
        for c in range(1, max_col + 1):
            if src_ws.cell(sr, c).value not in (None, ""):
                row_has_data = True
                break

        if not row_has_data:
            continue

        dr = start_append_row + written

        for c in range(1, max_col + 1):
            src_cell = src_ws.cell(sr, c)
            dst_cell = base_ws.cell(dr, c)
            dst_cell.value = src_cell.value
            copy_cell_style(src_cell, dst_cell)

        base_ws.row_dimensions[dr].height = src_ws.row_dimensions[sr].height
        written += 1

    return written


def fill_tab8_template(template_bytes, region_files_with_names):
    template_bytes.seek(0)
    base_wb = openpyxl.load_workbook(template_bytes, data_only=False)

    ws_sum_target, ws_append_target, sum_sheet_name, append_sheet_name = get_target_sheets_from_total_wb(base_wb)

    sorted_files = sorted(region_files_with_names, key=lambda x: sort_key_for_filename(x[0]))

    log = []
    warnings = []
    processed_regions = []

    for fname, fbytes in sorted_files:
        region = extract_region_from_filename(fname)

        if not region:
            warnings.append({
                "유형": "지역 인식 실패",
                "파일": fname,
                "설명": "파일명에서 지역명을 인식하지 못해 건너뜀"
            })
            continue

        fbytes.seek(0)
        wb_src = openpyxl.load_workbook(fbytes, data_only=False)
        ws_sum_src, ws_append_src, src_sum_name, src_append_name = get_source_sheets_from_region_wb(wb_src)

        sum_updated = accumulate_sheet_values(ws_sum_target, ws_sum_src)
        append_written = append_sheet_rows(ws_append_target, ws_append_src)

        processed_regions.append(region)

        log.append({
            "파일": fname,
            "지역": region,
            "원본시트1": src_sum_name,
            "원본시트2": src_append_name,
            "시트1합산셀수": sum_updated,
            "시트2추가행수": append_written
        })

    missing_regions = [r for r in REGION_ORDER_TAB8 if r not in processed_regions]
    for region in missing_regions:
        warnings.append({
            "유형": "파일 누락",
            "파일": "-",
            "설명": f"{region} 파일이 없어도 취합은 계속 진행됨"
        })

    return base_wb, log, warnings, sum_sheet_name, append_sheet_name


def render():
    st.caption("총괄표 파일의 시트1은 합산, 시트2는 시군별 상세를 순서대로 이어붙입니다.")

    template_file = st.file_uploader(
        "① 총괄표 파일 업로드",
        type=["xlsx"],
        key="tab8_template"
    )

    region_files = st.file_uploader(
        "② 시군별 파일 업로드 (여러 개)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="tab8_regions"
    )

    if template_file and region_files and st.button("🚀 취합 시작", key="tab8_run"):
        try:
            template_bytes = io.BytesIO(template_file.read())
            region_files_with_names = [(f.name, io.BytesIO(f.read())) for f in region_files]

            result_wb, log, warnings, sum_sheet_name, append_sheet_name = fill_tab8_template(
                template_bytes,
                region_files_with_names
            )

            output = io.BytesIO()
            result_wb.save(output)
            output.seek(0)

            st.success("취합이 완료되었습니다.")
            st.info(f"시트1(합산): '{sum_sheet_name}' / 시트2(붙이기): '{append_sheet_name}'")

            if warnings:
                st.warning("일부 경고가 있습니다.")
                st.dataframe(pd.DataFrame(warnings), use_container_width=True)

            if log:
                with st.expander("처리 로그 보기"):
                    st.dataframe(pd.DataFrame(log), use_container_width=True)

            st.download_button(
                label="📥 취합 결과 다운로드",
                data=output.getvalue(),
                file_name="tab8_실명법_취합결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"오류: {e}")
