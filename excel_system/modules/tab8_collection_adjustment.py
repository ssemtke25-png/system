"""
탭8: 총괄표 취합

- 총괄표(서식) 파일 1개 + 시군별 파일 여러 개를 받아
  빈 총괄표의 자기 지역 구간만 찾아 값을 채운다.
- 일부 시군 파일이 없어도 나머지는 계속 진행한다.
- 포항은 남구/북구 2개 구간을 모두 처리한다.
- 군위는 제외한다.
"""

import io
import re

import openpyxl
import streamlit as st
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter


# 군위 제외, 포항 남/북 분리
REGION_ORDER_TAB8 = [
    '포항남', '포항북', '경주', '김천', '안동', '구미', '영주', '영천', '상주', '문경',
    '경산', '의성', '청송', '영양', '영덕', '청도', '고령', '성주', '칠곡',
    '예천', '봉화', '울진', '울릉'
]

TOTAL_SHEET_NAME = '총괄표'


def normalize_region_label(text):
    if text is None:
        return None

    s = str(text).strip()
    s = s.replace(" ", "")
    s = s.replace("_", "")
    s = s.replace("-", "")

    mapping = {
        '포항남구': '포항남',
        '포항북구': '포항북',
        '포항남': '포항남',
        '포항북': '포항북',
        '포항시남구': '포항남',
        '포항시북구': '포항북',
        '경주시': '경주',
        '김천시': '김천',
        '안동시': '안동',
        '구미시': '구미',
        '영주시': '영주',
        '영천시': '영천',
        '상주시': '상주',
        '문경시': '문경',
        '경산시': '경산',
        '의성군': '의성',
        '청송군': '청송',
        '영양군': '영양',
        '영덕군': '영덕',
        '청도군': '청도',
        '고령군': '고령',
        '성주군': '성주',
        '칠곡군': '칠곡',
        '예천군': '예천',
        '봉화군': '봉화',
        '울진군': '울진',
        '울릉군': '울릉',
    }

    if s in mapping:
        return mapping[s]

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

    for region in REGION_ORDER_TAB8:
        if region in base:
            return region

    return None


def find_table_boundaries(ws, max_scan_row=500):
    """
    '1. ...', '2. ...' 같은 제목행을 기준으로 표 구간을 나눈다.
    """
    title_rows = []

    for r in range(1, min(ws.max_row, max_scan_row) + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and re.match(r'^\d+\.\s*', v.strip()):
            title_rows.append(r)

    if not title_rows:
        return [(1, ws.max_row)]

    boundaries = []
    for i, start in enumerate(title_rows):
        end = title_rows[i + 1] - 1 if i + 1 < len(title_rows) else ws.max_row
        boundaries.append((start, end))
    return boundaries


def find_region_blocks_in_range(ws, start_row, end_row):
    """
    표 구간 안에서 A열의 시군명을 찾아 시작행을 기록한다.
    """
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
    if not getattr(fill, 'patternType', None):
        return False

    fg = getattr(fill, 'fgColor', None)
    if not fg:
        return False

    fg_type = getattr(fg, 'type', None)
    if fg_type == 'theme':
        theme = getattr(fg, 'theme', None)
        if theme is not None and theme != 0:
            return True
    elif fg_type == 'rgb':
        rgb = getattr(fg, 'rgb', None)
        if rgb not in (None, '00000000', 'FFFFFFFF'):
            return True

    return False


def is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def is_protected_cell(cell):
    return is_formula(cell.value) or has_protective_color(cell)


def fill_total_sheet(base_ws, src_ws, own_region, warnings, fname):
    """
    총괄표의 각 표에서 own_region에 해당하는 구간만 src에서 가져와 채운다.
    """
    base_tables = find_table_boundaries(base_ws)
    src_tables = find_table_boundaries(src_ws)

    total_written = 0
    table_count = min(len(base_tables), len(src_tables))

    if len(base_tables) != len(src_tables):
        warnings.append({
            "유형": "표 개수 다름",
            "파일": fname,
            "설명": f"서식 표 {len(base_tables)}개 / 원본 표 {len(src_tables)}개 → 앞쪽 {table_count}개만 처리"
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

        if base_start_row is None or src_start_row is None:
            warnings.append({
                "유형": "지역 구간 없음",
                "파일": fname,
                "설명": f"{i+1}번 표에서 '{own_region}' 구간을 찾지 못해 건너뜀"
            })
            continue

        block_size = get_block_size(base_blocks, b_end)
        max_col = max(base_ws.max_column, src_ws.max_column)

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

                if isinstance(src_val, str) and src_val.startswith("#"):
                    warnings.append({
                        "유형": "오류 값 발견",
                        "파일": fname,
                        "설명": f"{get_column_letter(c)}{sr} 값이 '{src_val}'라서 건너뜀"
                    })
                    continue

                if src_val is None:
                    continue

                base_cell.value = src_val
                total_written += 1

    return total_written


def sort_key_for_filename(fname):
    m = re.match(r'^(\d{1,3})[._\s-]', fname)
    if m:
        return int(m.group(1))

    region = extract_region_from_filename(fname)
    if region in REGION_ORDER_TAB8:
        return 100 + REGION_ORDER_TAB8.index(region)

    return 999


def fill_tab8_template(template_bytes, region_files_with_names):
    """
    template_bytes: 총괄표(서식) 파일 BytesIO
    region_files_with_names: [(파일명, BytesIO), ...]
    """
    log = []
    warnings = []

    template_bytes.seek(0)
    base_wb = openpyxl.load_workbook(template_bytes, data_only=False)

    if TOTAL_SHEET_NAME not in base_wb.sheetnames:
        raise ValueError(f"서식 파일에 '{TOTAL_SHEET_NAME}' 시트가 없습니다.")

    base_ws = base_wb[TOTAL_SHEET_NAME]

    sorted_files = sorted(region_files_with_names, key=lambda x: sort_key_for_filename(x[0]))

    processed_regions = set()

    for fname, fbytes in sorted_files:
        own_region = extract_region_from_filename(fname)

        if not own_region:
            warnings.append({
                "유형": "지역 인식 실패",
                "파일": fname,
                "설명": "파일명에서 시군명을 인식하지 못해 건너뜀"
            })
            continue

        if own_region == '군위':
            warnings.append({
                "유형": "처리 제외",
                "파일": fname,
                "설명": "군위는 처리 대상이 아니므로 건너뜀"
            })
            continue

        try:
            fbytes.seek(0)
            wb_src = openpyxl.load_workbook(fbytes, data_only=True)
        except Exception as e:
            warnings.append({
                "유형": "파일 열기 오류",
                "파일": fname,
                "설명": str(e)
            })
            continue

        if TOTAL_SHEET_NAME not in wb_src.sheetnames:
            warnings.append({
                "유형": "시트 없음",
                "파일": fname,
                "설명": f"'{TOTAL_SHEET_NAME}' 시트가 없어 건너뜀"
            })
            continue

        src_ws = wb_src[TOTAL_SHEET_NAME]

        written = fill_total_sheet(base_ws, src_ws, own_region, warnings, fname)
        processed_regions.add(own_region)

        log.append({
            "파일": fname,
            "지역": own_region,
            "입력셀수": written
        })

    missing_regions = [r for r in REGION_ORDER_TAB8 if r not in processed_regions]

    for region in missing_regions:
        warnings.append({
            "유형": "누락 지역",
            "파일": "-",
            "설명": f"{region} 파일이 없어 해당 구간은 비워둠"
        })

    return base_wb, log, warnings


def render():
    st.caption(
        "총괄표(서식) 파일 1개와 시군별 파일 여러 개를 받아 "
        "빈 총괄표의 해당 지역 칸만 채웁니다."
    )

    st.info(
        "📌 군위는 제외하고 처리합니다.\n\n"
        "📌 포항은 남구/북구를 각각 별도 지역으로 처리합니다.\n\n"
        "📌 일부 시군 파일이 없어도 취합은 계속 진행되며, 없는 지역은 빈칸으로 남습니다."
    )

    template_file = st.file_uploader(
        "① 총괄표(서식) 파일 업로드",
        type=["xlsx"],
        key="tab8_template"
    )

    region_files = st.file_uploader(
        "② 시군별 파일 업로드 (여러 개)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="tab8_regions"
    )

    if template_file and region_files and st.button("🚀 tab8 취합 시작", key="tab8_run"):
        try:
            template_bytes = io.BytesIO(template_file.read())
            region_files_with_names = [(f.name, io.BytesIO(f.read())) for f in region_files]

            result_wb, log, warnings = fill_tab8_template(template_bytes, region_files_with_names)

            out = io.BytesIO()
            result_wb.save(out)

            st.success("tab8 총괄표 취합이 완료되었습니다.")
            st.download_button(
                "📥 다운로드",
                data=out.getvalue(),
                file_name="tab8_총괄표_취합결과.xlsx",
                key="tab8_download"
            )

            if warnings:
                st.warning(f"확인이 필요한 항목 {len(warnings)}건이 있습니다. (결과 파일은 생성됨)")
                st.dataframe(warnings, use_container_width=True)
            else:
                st.info("특이사항 없이 정상 처리되었습니다.")

            with st.expander("처리 로그 보기"):
                st.dataframe(log, use_container_width=True)

        except Exception as e:
            st.error(f"오류: {e}")
            st.exception(e)
