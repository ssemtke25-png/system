import io
import re
from copy import copy as _copy_style

import openpyxl


TOTAL_SHEET_NAME = "총괄표"

REGION_ORDER = [
    "포항", "경주", "김천", "안동", "구미", "영주", "영천", "상주", "문경",
    "경산", "군위", "의성", "청송", "영양", "영덕", "청도", "고령",
    "성주", "칠곡", "예천", "봉화", "울진", "울릉"
]


def normalize_region_label(text):
    if text is None:
        return ""
    s = str(text).strip()
    s = re.sub(r"\s+", "", s)
    s = s.replace("시청", "").replace("군청", "")
    s = s.replace("시", "").replace("군", "")
    return s


def extract_region_from_filename(filename):
    base = filename.rsplit(".", 1)[0]
    return normalize_region_label(base)


def sort_key_for_filename(fname):
    region = extract_region_from_filename(fname)
    if region in REGION_ORDER:
        return (0, REGION_ORDER.index(region), fname)
    return (1, 999, fname)


def get_total_sheet(wb):
    for s in wb.sheetnames:
        if str(s).strip() == TOTAL_SHEET_NAME:
            return wb[s], s
    for s in wb.sheetnames:
        if TOTAL_SHEET_NAME in str(s).strip():
            return wb[s], s
    first_name = wb.sheetnames[0]
    return wb[first_name], first_name


def get_first_two_sheets(wb):
    names = wb.sheetnames
    if len(names) == 1:
        return wb[names[0]], wb[names[0]], names[0], names[0]
    return wb[names[0]], wb[names[1]], names[0], names[1]


def is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def is_red_rgb(rgb):
    if not rgb:
        return False
    rgb = str(rgb).upper()
    return rgb in ("FFFF0000", "00FF0000")


def is_red_border_cell(cell):
    border = cell.border
    for side in [border.left, border.right, border.top, border.bottom]:
        color = getattr(side, "color", None)
        if color is None:
            continue
        rgb = getattr(color, "rgb", None)
        if is_red_rgb(rgb):
            return True

        indexed = getattr(color, "indexed", None)
        if indexed == 10:
            return True
    return False


def clear_cell_value_only(cell):
    cell.value = None


def accumulate_only_red_border_cells(base_ws, src_ws):
    max_row = min(base_ws.max_row, src_ws.max_row)
    max_col = min(base_ws.max_column, src_ws.max_column)
    updated = 0
    skipped = []

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            base_cell = base_ws.cell(r, c)

            if not is_red_border_cell(base_cell):
                continue

            if is_formula(base_cell.value):
                continue

            src_cell = src_ws.cell(r, c)
            src_val = src_cell.value
            base_val = base_cell.value

            if src_val in (None, ""):
                continue

            if not is_number(src_val):
                skipped.append({
                    "cell": f"{base_cell.coordinate}",
                    "reason": "source_not_number",
                    "value": src_val
                })
                continue

            if base_val in (None, ""):
                base_cell.value = src_val
                updated += 1
            elif is_number(base_val):
                base_cell.value = base_val + src_val
                updated += 1
            else:
                skipped.append({
                    "cell": f"{base_cell.coordinate}",
                    "reason": "target_not_number",
                    "value": base_val
                })

    return updated, skipped


def detect_detail_header_row(ws, search_max_row=20, search_max_col=20):
    for r in range(1, min(ws.max_row, search_max_row) + 1):
        texts = []
        for c in range(1, min(ws.max_column, search_max_col) + 1):
            v = ws.cell(r, c).value
            texts.append("" if v is None else str(v).strip())
        joined = " ".join(texts)

        score = 0
        if "시군" in joined or "시,군" in joined or "시, 군" in joined:
            score += 1
        if "년도" in joined:
            score += 1
        if "성명" in joined:
            score += 1
        if "부과액" in joined:
            score += 1
        if score >= 2:
            return r
    return None


def extract_detail_rows(src_ws, max_cols=None):
    if max_cols is None:
        max_cols = src_ws.max_column

    header_row = detect_detail_header_row(src_ws)
    if not header_row:
        return []

    rows = []
    for r in range(header_row + 1, src_ws.max_row + 1):
        vals = [src_ws.cell(r, c).value for c in range(1, max_cols + 1)]
        texts = [str(v).strip() if v is not None else "" for v in vals]
        joined = " ".join(texts)
        compact = "".join(texts)

        if all(v in (None, "") for v in vals):
            continue

        if "합계" in compact:
            continue
        if "작성방법" in compact or "작성 방법" in joined:
            continue
        if "미수납현황" in compact or "미수납 현황" in joined:
            continue
        if "사용자:" in joined:
            continue
        if "페이지" in joined:
            continue

        rows.append(vals)

    return rows


def find_sheet2_data_start_row(base_ws, search_max_row=30, search_max_col=20):
    header_row = detect_detail_header_row(base_ws, search_max_row, search_max_col)
    if header_row:
        return header_row + 1
    return 11


def clear_sheet2_data_area(base_ws, start_row, max_cols):
    for r in range(start_row, base_ws.max_row + 1):
        for c in range(1, max_cols + 1):
            clear_cell_value_only(base_ws.cell(r, c))


def write_detail_rows(base_ws, all_rows, start_row=None, max_cols=None):
    if max_cols is None:
        max_cols = base_ws.max_column
    if start_row is None:
        start_row = find_sheet2_data_start_row(base_ws)

    clear_sheet2_data_area(base_ws, start_row, max_cols)

    current_row = start_row
    for row_vals in all_rows:
        for c in range(1, max_cols + 1):
            value = row_vals[c - 1] if c - 1 < len(row_vals) else None
            base_ws.cell(current_row, c).value = value
        current_row += 1

    return current_row - start_row


def copy_sheet_layout_if_needed(ws):
    # 현재 요청 기준에서는 시트2 레이아웃 복사 없이
    # 총괄표 기존 틀 유지가 원칙이라 별도 동작 없음
    return ws


def build_tab8_total_workbook(template_bytes, region_files):
    template_bytes.seek(0)
    base_wb = openpyxl.load_workbook(template_bytes)
    ws_sum_target, total_sheet_name = get_total_sheet(base_wb)

    if len(base_wb.sheetnames) >= 2:
        ws_detail_target = base_wb[base_wb.sheetnames[1]]
    else:
        ws_detail_target = ws_sum_target

    sorted_files = sorted(region_files, key=lambda x: sort_key_for_filename(x[0]))

    all_detail_rows = []
    log = []
    warnings = []

    for fname, fbytes in sorted_files:
        try:
            if hasattr(fbytes, "seek"):
                fbytes.seek(0)

            wb_src = openpyxl.load_workbook(fbytes, data_only=False)
            ws_sum_src, ws_detail_src, src_sum_name, src_detail_name = get_first_two_sheets(wb_src)

            updated, skipped = accumulate_only_red_border_cells(ws_sum_target, ws_sum_src)
            detail_rows = extract_detail_rows(ws_detail_src, max_cols=ws_detail_target.max_column)
            all_detail_rows.extend(detail_rows)

            log.append({
                "file": fname,
                "region": extract_region_from_filename(fname),
                "sheet1_updates": updated,
                "sheet2_rows_added": len(detail_rows),
                "source_sheet1": src_sum_name,
                "source_sheet2": src_detail_name
            })

            for item in skipped:
                warnings.append({
                    "file": fname,
                    "region": extract_region_from_filename(fname),
                    "sheet": src_sum_name,
                    "cell": item["cell"],
                    "reason": item["reason"],
                    "value": item["value"]
                })

        except Exception as e:
            warnings.append({
                "file": fname,
                "region": extract_region_from_filename(fname),
                "sheet": "",
                "cell": "",
                "reason": "load_or_process_error",
                "value": str(e)
            })

    written_count = write_detail_rows(
        ws_detail_target,
        all_detail_rows,
        start_row=find_sheet2_data_start_row(ws_detail_target),
        max_cols=ws_detail_target.max_column
    )

    log.append({
        "file": "TOTAL",
        "region": "",
        "sheet1_updates": "",
        "sheet2_rows_added": written_count,
        "source_sheet1": total_sheet_name,
        "source_sheet2": ws_detail_target.title
    })

    output = io.BytesIO()
    base_wb.save(output)
    output.seek(0)

    return output, log, warnings
