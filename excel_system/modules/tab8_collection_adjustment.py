import io
import re
from copy import copy as _copy_style

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.cell.cell import MergedCell


REGION_ORDER_TAB8 = [
    "포항남", "포항북", "경주", "김천", "안동", "구미", "영주", "영천", "상주", "문경",
    "경산", "의성", "청송", "영양", "영덕", "청도", "고령", "성주", "칠곡",
    "예천", "봉화", "울진", "울릉"
]

TOTAL_SHEET_NAME = "총괄표"


def normalize_region_label(text):
    if text is None:
        return None

    s = str(text).strip()
    s = s.replace(" ", "")
    s = s.replace("시", "").replace("군", "").replace("구", "")

    s = s.replace("포항시남", "포항남")
    s = s.replace("포항남구", "포항남")
    s = s.replace("포항시북", "포항북")
    s = s.replace("포항북구", "포항북")

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
        region = normalize_region_label(m.group(1))
        if region:
            return region

    return None


def sort_key_for_filename(fname):
    m = re.match(r'^(\d{1,3})[._\s-]', fname)
    if m:
        return int(m.group(1))

    region = extract_region_from_filename(fname)
    if region in REGION_ORDER_TAB8:
        return 1000 + REGION_ORDER_TAB8.index(region)

    return 9999


def get_total_target_sheets(wb):
    if len(wb.sheetnames) < 2:
        raise ValueError("총괄표 파일에는 최소 2개 시트가 필요합니다. (시트1=합산, 시트2=붙이기)")

    ws_sum = wb[wb.sheetnames[0]]
    ws_append = wb[wb.sheetnames[1]]
    return ws_sum, ws_append, wb.sheetnames[0], wb.sheetnames[1]


def get_region_primary_sheet(wb):
    for s in wb.sheetnames:
        if str(s).strip() == TOTAL_SHEET_NAME:
            return wb[s], s

    for s in wb.sheetnames:
        if TOTAL_SHEET_NAME in str(s).strip():
            return wb[s], s

    return wb[wb.sheetnames[0]], wb.sheetnames[0]


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def is_formula_cell(cell):
    return isinstance(cell.value, str) and cell.value.startswith("=")


def has_fill(cell):
    fill = cell.fill
    return bool(getattr(fill, "patternType", None))


def is_protected_target_cell(cell):
    if isinstance(cell, MergedCell):
        return True
    if is_formula_cell(cell):
        return True
    return False


def accumulate_sheet_values(base_ws, src_ws, warnings, fname, src_sheet_name):
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

            if is_number(src_val):
                if base_val is None or base_val == "":
                    base_cell.value = src_val
                    updated += 1
                elif is_number(base_val):
                    base_cell.value = base_val + src_val
                    updated += 1
                else:
                    # 기준표 셀이 문자면 합산 불가 -> 건너뜀
                    continue

    return updated


def detect_header_row_count(ws, scan_limit=15):
    # 단순 기본값: 1행 헤더
    # 필요 시 나중에 2행 헤더로 바꿀 수 있게 함수 분리
    return 1


def find_last_data_row(ws):
    for r in range(ws.max_row, 0, -1):
        for c in range(1, ws.max_column + 1):
            if ws.cell(r, c).value not in (None, ""):
                return r
    return 0


def copy_row_style(src_ws, src_row, dst_ws, dst_row, max_col):
    for c in range(1, max_col + 1):
        src_cell = src_ws.cell(src_row, c)
        dst_cell = dst_ws.cell(dst_row, c)

        if isinstance(dst_cell, MergedCell):
            continue

        dst_cell.font = _copy_style(src_cell.font)
        dst_cell.fill = _copy_style(src_cell.fill)
        dst_cell.border = _copy_style(src_cell.border)
        dst_cell.alignment = _copy_style(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = _copy_style(src_cell.protection)


def append_sheet_rows(base_ws, src_ws, fname, src_sheet_name):
    header_rows = detect_header_row_count(src_ws)
    src_last_row = find_last_data_row(src_ws)
    if src_last_row <= header_rows:
        return 0

    start_append_row = find_last_data_row(base_ws) + 1
    max_col = min(base_ws.max_column, src_ws.max_column)

    written = 0

    for sr in range(header_rows + 1, src_last_row + 1):
        dr = start_append_row + written

        row_has_data = False
        for c in range(1, max_col + 1):
            v = src_ws.cell(sr, c).value
            if v not in (None, ""):
                row_has_data = True
                break

        if not row_has_data:
            continue

        for c in range(1, max_col + 1):
            base_ws.cell(dr, c).value = src_ws.cell(sr, c).value

        copy_row_style(src_ws, sr, base_ws, dr, max_col)
        base_ws.row_dimensions[dr].height = src_ws.row_dimensions[sr].height

        written += 1

    return written


def copy_column_widths(src_ws, dst_ws, max_col=100):
    for c in range(1, min(src_ws.max_column, max_col) + 1):
        col_letter = openpyxl.utils.get_column_letter(c)
        if col_letter in src_ws.column_dimensions:
            dst_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width


def fill_tab8_template(template_bytes, region_files_with_names):
    template_bytes.seek(0)
    base_wb = openpyxl.load_workbook(template_bytes, data_only=False)

    ws_sum, ws_append, sum_sheet_name, append_sheet_name = get_total_target_sheets(base_wb)

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
                "시트": "-",
                "설명": "파일명에서 지역명을 인식하지 못해 건너뜀"
            })
            continue

        fbytes.seek(0)
        wb_src = openpyxl.load_workbook(fbytes, data_only=False)
        src_ws, src_sheet_name = get_region_primary_sheet(wb_src)

        sum_updated = accumulate_sheet_values(
            ws_sum, src_ws, warnings, fname, src_sheet_name
        )

        append_written = append_sheet_rows(
            ws_append, src_ws, fname, src_sheet_name
        )

        processed_regions.append(region)

        log.append({
            "파일": fname,
            "지역": region,
            "사용시트": src_sheet_name,
            "시트1합산셀수": sum_updated,
            "시트2추가행수": append_written
        })

    missing_regions = [r for r in REGION_ORDER_TAB8 if r not in processed_regions]
    for region in missing_regions:
        warnings.append({
            "유형": "파일 누락",
            "파일": "-",
            "시트": "-",
            "설명": f"{region} 파일이 없어도 취합은 계속 진행됨"
        })

    return base_wb, log, warnings, sum_sheet_name, append_sheet_name


def render():
    st.caption("총괄표 파일의 시트1은 합산, 시트2는 시군 순서대로 이어붙이기 방식으로 처리합니다.")

    st.info(
        "📌 총괄표 업로드 파일의 첫 번째 시트는 합산용, 두 번째 시트는 붙이기용으로 사용합니다.\n\n"
        "📌 시군 파일은 '총괄표' 시트를 우선 사용하고, 없으면 첫 번째 시트를 사용합니다.\n\n"
        "📌 시트1은 같은 위치의 숫자 셀끼리 누적 합산합니다.\n\n"
        "📌 시트2는 각 시군 데이터를 순서대로 아래에 이어 붙입니다."
    )

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
