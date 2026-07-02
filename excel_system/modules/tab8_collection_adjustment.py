"""
탭8: 실명법 취합

핵심 원칙
- '총괄표 업로드' 칸에 올린 파일은 무조건 총괄표 기준 파일로 인정
- 파일 안에 '총괄표' 시트가 있으면 우선 사용
- 없으면 첫 번째 시트를 총괄표 시트처럼 사용
- 시군 파일도 '총괄표' 시트가 있으면 우선 사용, 없으면 첫 번째 시트 사용
- 일부 파일이 누락되어도 가능한 범위에서 계속 취합
"""

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

TOTAL_SHEET_NAME = "총괄표"


def is_formula(value):
    return isinstance(value, str) and value.startswith("=")


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


def get_primary_sheet_from_uploaded_total_file(wb):
    # 1순위: 총괄표 완전일치
    for s in wb.sheetnames:
        if str(s).strip() == TOTAL_SHEET_NAME:
            return wb[s], s

    # 2순위: 총괄표 포함
    for s in wb.sheetnames:
        if TOTAL_SHEET_NAME in str(s).strip():
            return wb[s], s

    # 3순위: 첫 번째 시트
    if wb.sheetnames:
        first_name = wb.sheetnames[0]
        return wb[first_name], first_name

    raise ValueError("업로드된 총괄표 파일에 시트가 없습니다.")


def get_primary_sheet_from_region_file(wb):
    # 시군 파일도 동일 원칙 적용
    for s in wb.sheetnames:
        if str(s).strip() == TOTAL_SHEET_NAME:
            return wb[s], s

    for s in wb.sheetnames:
        if TOTAL_SHEET_NAME in str(s).strip():
            return wb[s], s

    if wb.sheetnames:
        first_name = wb.sheetnames[0]
        return wb[first_name], first_name

    raise ValueError("업로드된 시군 파일에 시트가 없습니다.")


def find_table_boundaries(ws, max_scan_row=500):
    title_rows = []
    max_row = min(ws.max_row, max_scan_row)

    for r in range(1, max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and re.match(r'^\d+\.\s*', v.strip()):
            title_rows.append(r)

    if not title_rows:
        return [(1, ws.max_row)]

    boundaries = []
    for i, start_row in enumerate(title_rows):
        end_row = title_rows[i + 1] - 1 if i + 1 < len(title_rows) else ws.max_row
        boundaries.append((start_row, end_row))
    return boundaries


def find_region_blocks_in_range(ws, start_row, end_row):
    blocks = []
    for r in range(start_row, end_row + 1):
        label = ws.cell(r, 1).value
        key = normalize_region_label(label)
        if key in REGION_ORDER_TAB8:
            blocks.append((key, r))
    return blocks


def get_block_size(blocks, end_row):
    if len(blocks) >= 2:
        return blocks[1][1] - blocks[0][1]
    elif len(blocks) == 1:
        return max(1, end_row - blocks[0][1] + 1)
    return 1


def has_protective_color(cell):
    fill = cell.fill
    if not getattr(fill, "patternType", None):
        return False

    fg = getattr(fill, "fgColor", None)
    if not fg:
        return False

    fg_type = getattr(fg, "type", None)

    if fg_type == "theme":
        theme = getattr(fg, "theme", None)
        if theme is not None and theme != 0:
            return True

    elif fg_type == "rgb":
        rgb = getattr(fg, "rgb", None)
        if rgb not in (None, "00000000", "FFFFFFFF"):
            return True

    return False


def is_protected_cell(cell):
    return is_formula(cell.value) or has_protective_color(cell)


def fill_total_sheet(base_ws, src_ws, own_region, warnings, fname, src_sheet_name):
    base_tables = find_table_boundaries(base_ws)
    src_tables = find_table_boundaries(src_ws)

    table_count = min(len(base_tables), len(src_tables))
    copied_rows = 0

    if len(base_tables) != len(src_tables):
        warnings.append({
            "유형": "표 개수 다름",
            "파일": fname,
            "시트": src_sheet_name,
            "설명": f"기준표 {len(base_tables)}개 / 원본표 {len(src_tables)}개 → 앞쪽 {table_count}개만 처리"
        })

    for i in range(table_count):
        b_start, b_end = base_tables[i]
        s_start, s_end = src_tables[i]

        base_blocks = find_region_blocks_in_range(base_ws, b_start, b_end)
        src_blocks = find_region_blocks_in_range(src_ws, s_start, s_end)

        base_map = dict(base_blocks)
        src_map = dict(src_blocks)

        base_start_row = base_map.get(own_region)
        src_start_row = src_map.get(own_region)

        if base_start_row is None:
            warnings.append({
                "유형": "기준표 지역 없음",
                "파일": fname,
                "시트": src_sheet_name,
                "설명": f"기준표에서 '{own_region}' 지역 시작행을 찾지 못함"
            })
            continue

        if src_start_row is None:
            warnings.append({
                "유형": "원본표 지역 없음",
                "파일": fname,
                "시트": src_sheet_name,
                "설명": f"원본표에서 '{own_region}' 지역 시작행을 찾지 못함"
            })
            continue

        block_size = get_block_size(base_blocks, b_end)
        max_col = min(max(base_ws.max_column, src_ws.max_column), 300)

        for offset in range(block_size):
            br = base_start_row + offset
            sr = src_start_row + offset

            if br > b_end or sr > s_end:
                break

            for c in range(1, max_col + 1):
                base_cell = base_ws.cell(br, c)

                if isinstance(base_cell, MergedCell):
                    continue

                if is_protected_cell(base_cell):
                    continue

                src_val = src_ws.cell(sr, c).value

                if src_val is None:
                    continue

                if isinstance(src_val, str) and src_val.startswith("#"):
                    warnings.append({
                        "유형": "오류 값 발견",
                        "파일": fname,
                        "시트": src_sheet_name,
                        "셀": f"{get_column_letter(c)}{sr}",
                        "설명": f"원본 값 '{src_val}' 발견으로 해당 셀 건너뜀"
                    })
                    continue

                base_cell.value = src_val

            copied_rows += 1

    return copied_rows


def copy_sheet_layout_if_needed(base_ws, max_rows=300, max_cols=50):
    template_row = None
    scan_row = min(base_ws.max_row, max_rows)
    scan_col = min(base_ws.max_column, max_cols)

    for r in range(1, scan_row + 1):
        values = [base_ws.cell(r, c).value for c in range(1, min(6, scan_col) + 1)]
        if any(v is not None for v in values):
            template_row = r
            break

    if template_row is None:
        return

    for r in range(template_row + 1, scan_row + 1):
        for c in range(1, scan_col + 1):
            cell = base_ws.cell(r, c)
            tmpl = base_ws.cell(template_row, c)

            if isinstance(cell, MergedCell):
                continue

            if cell.has_style:
                continue

            cell.font = _copy_style(tmpl.font)
            cell.fill = _copy_style(tmpl.fill)
            cell.border = _copy_style(tmpl.border)
            cell.alignment = _copy_style(tmpl.alignment)
            cell.number_format = tmpl.number_format
            cell.protection = _copy_style(tmpl.protection)


def fill_tab8_template(template_bytes, region_files_with_names):
    template_bytes.seek(0)
    base_wb = openpyxl.load_workbook(template_bytes, data_only=False)

    base_ws, selected_base_sheet_name = get_primary_sheet_from_uploaded_total_file(base_wb)

    sorted_files = sorted(region_files_with_names, key=lambda x: sort_key_for_filename(x[0]))

    log = []
    warnings = []
    processed_regions = set()

    for fname, fbytes in sorted_files:
        own_region = extract_region_from_filename(fname)

        if not own_region:
            warnings.append({
                "유형": "지역 인식 실패",
                "파일": fname,
                "시트": "-",
                "설명": "파일명에서 지역명을 인식하지 못해 건너뜀"
            })
            continue

        if own_region not in REGION_ORDER_TAB8:
            warnings.append({
                "유형": "처리 대상 아님",
                "파일": fname,
                "시트": "-",
                "설명": f"'{own_region}'은 탭8 처리 대상이 아니므로 건너뜀"
            })
            continue

        fbytes.seek(0)
        wb_src = openpyxl.load_workbook(fbytes, data_only=False)
        src_ws, selected_src_sheet_name = get_primary_sheet_from_region_file(wb_src)

        copied_rows = fill_total_sheet(
            base_ws=base_ws,
            src_ws=src_ws,
            own_region=own_region,
            warnings=warnings,
            fname=fname,
            src_sheet_name=selected_src_sheet_name,
        )

        processed_regions.add(own_region)

        log.append({
            "파일": fname,
            "지역": own_region,
            "사용시트": selected_src_sheet_name,
            "처리행수": copied_rows
        })

    missing_regions = [r for r in REGION_ORDER_TAB8 if r not in processed_regions]
    for region in missing_regions:
        warnings.append({
            "유형": "파일 누락",
            "파일": "-",
            "시트": "-",
            "설명": f"{region} 파일이 없어도 취합은 계속 진행됨"
        })

    copy_sheet_layout_if_needed(base_ws)

    return base_wb, log, warnings, selected_base_sheet_name


def render():
    st.caption("총괄표 업로드 칸에 올린 파일을 기준으로 실명법 취합을 수행합니다.")

    st.info(
        "📌 총괄표 업로드 칸에 넣은 파일은 무조건 기준 파일로 사용합니다.\n\n"
        "📌 파일 안에 '총괄표' 시트가 있으면 우선 사용하고, 없으면 첫 번째 시트를 사용합니다.\n\n"
        "📌 시군 파일도 같은 방식으로 처리합니다.\n\n"
        "📌 수식 셀, 색이 칠해진 보호 셀은 덮어쓰지 않습니다."
    )

    template_file = st.file_uploader(
        "① 총괄표(기준) 파일 업로드",
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

            result_wb, log, warnings, selected_base_sheet_name = fill_tab8_template(
                template_bytes,
                region_files_with_names
            )

            output = io.BytesIO()
            result_wb.save(output)
            output.seek(0)

            st.success("취합이 완료되었습니다.")
            st.info(f"기준 시트로 '{selected_base_sheet_name}' 시트를 사용했습니다.")

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
