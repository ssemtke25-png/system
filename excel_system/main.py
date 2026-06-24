import streamlit as st
import io
import re
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

tab1, tab2 = st.tabs(["① 단순 합산", "② 총괄표 채우기 (시군구)"])

# =========================================================================
# 공통 유틸
# =========================================================================
PREFIX_SPECIAL = {
    '포항시남구': '포항남', '포항남구': '포항남', '포항남': '포항남',
    '포항시북구': '포항북', '포항북구': '포항북', '포항북': '포항북',
}
VALID_REGION_KEYS = {
    '포항남', '포항북', '경주', '김천', '안동', '구미', '영주', '영천', '상주', '문경',
    '경산', '의성', '청송', '영양', '영덕', '청도', '고령', '성주', '칠곡', '예천',
    '봉화', '울진', '울릉'
}

def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)

def is_formula(v):
    return isinstance(v, str) and v.startswith("=")

def get_safe_value(v):
    """값을 안전하게 스마트 변환 ('-' 표기는 0으로, 에러는 무시, 나머진 그대로)"""
    if v is None:
        return None
    if is_number(v):
        return v
    if isinstance(v, str):
        s = v.strip()
        if s in ['-', '']:  # 하이픈이나 빈칸은 0으로
            return 0
        if s.startswith('#'):  # #VALUE! 등 수식 에러는 무시
            return None
    return v

def region_key(name):
    if not name or not isinstance(name, str):
        return None
    n = re.sub(r'\s+', '', name.strip())
    n = n.replace('광역시', '').replace('특별시', '')
    if n in PREFIX_SPECIAL:
        return PREFIX_SPECIAL[n]
    n2 = re.sub(r'(시|군|구)$', '', n)
    return n2 if n2 else None

def is_valid_region(label):
    return region_key(label) in VALID_REGION_KEYS

def extract_own_region_from_filename(filename):
    # 1. 숫자와 지역명 조합 찾기
    m = re.match(r'^\d{1,3}[_.\s]+([가-힣]+)', filename)
    if m:
        key = region_key(m.group(1))
        if key:
            return key
            
    # 2. 파일명 내 키워드 직접 검색
    for raw_name, mapped_key in PREFIX_SPECIAL.items():
        if raw_name in filename.replace(' ', ''):
            return mapped_key
            
    for k in VALID_REGION_KEYS:
        if k in filename:
            return k
    return None

def get_sheet_by_index(wb, idx):
    if idx < len(wb.sheetnames):
        return wb[wb.sheetnames[idx]]
    return None

def target_keys_for_region(own_key):
    if own_key == '포항':
        return {'포항남', '포항북'}
    return {own_key} if own_key else set()

# =========================================================================
# 탭1: 단순 합산
# =========================================================================
def aggregate(file_bytes_list, file_names):
    warnings = []
    n_files = len(file_bytes_list)

    value_wbs = [openpyxl.load_workbook(b, data_only=True) for b in file_bytes_list]
    file_bytes_list[0].seek(0)
    base_wb = openpyxl.load_workbook(file_bytes_list[0], data_only=False)
    names = file_names

    all_sheets = []
    for vwb in value_wbs:
        for s in vwb.sheetnames:
            if s not in all_sheets:
                all_sheets.append(s)

    for sheet in all_sheets:
        present_idx = [i for i in range(n_files) if sheet in value_wbs[i].sheetnames]

        if len(present_idx) < n_files:
            missing_files = [names[i] for i in range(n_files) if i not in present_idx]
            warnings.append({
                "유형": "시트 누락", "시트": sheet, "셀": "-",
                "파일": ", ".join(missing_files),
                "설명": f"'{sheet}' 시트가 없어 해당 파일은 이 시트 합산에서 제외됨"
            })

        if not present_idx:
            continue

        if sheet not in base_wb.sheetnames:
            base_wb.create_sheet(sheet)

        # 무한 로딩 방지 안전장치
        max_r = min(max(value_wbs[i][sheet].max_row for i in present_idx), 300)
        max_c = min(max(value_wbs[i][sheet].max_column for i in present_idx), 100)
        base_ws = base_wb[sheet]

        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                cell_addr = f"{get_column_letter(c)}{r}"
                base_cell = base_ws.cell(r, c)

                if isinstance(base_cell, MergedCell):
                    continue
                if is_formula(base_cell.value):
                    continue

                total = 0
                any_number = False
                text_files = []

                for i in present_idx:
                    vws = value_wbs[i][sheet]
                    if r > vws.max_row or c > vws.max_column:
                        continue
                    v = vws.cell(r, c).value
                    if is_number(v):
                        total += v
                        any_number = True
                    elif v is not None and isinstance(v, str):
                        text_files.append((names[i], v))

                if any_number and text_files:
                    warnings.append({
                        "유형": "텍스트 혼입", "시트": sheet, "셀": cell_addr,
                        "파일": ", ".join(fn for fn, _ in text_files),
                        "설명": "숫자가 들어가야 할 칸에 텍스트 발견 ("
                                + ", ".join(f"{fn}='{val}'" for fn, val in text_files)
                                + ") → 0으로 처리하여 합산함"
                    })

                if any_number:
                    base_cell.value = total

    formula_wbs = []
    for b in file_bytes_list:
        b.seek(0)
        formula_wbs.append(openpyxl.load_workbook(b, data_only=False))

    return base_wb, warnings

with tab1:
    st.caption("여러 파일의 같은 위치 숫자를 모두 더합니다. 수식이 있는 칸은 건드리지 않고 그대로 보존합니다.")
    files1 = st.file_uploader("파일 업로드", type=["xlsx"], accept_multiple_files=True, key="up1")

    if files1 and st.button("🚀 취합 시작", key="btn1"):
        try:
            file_bytes_list = [io.BytesIO(f.read()) for f in files1]
            file_names = [f.name for f in files1]

            result_wb, warns = aggregate(file_bytes_list, file_names)

            o = io.BytesIO()
            result_wb.save(o)

            st.success("취합이 완료되었습니다.")
            st.download_button("📥 다운로드", o.getvalue(), "result.xlsx", key="dl1")

            if warns:
                st.warning(f"⚠️ 확인이 필요한 항목 {len(warns)}건이 발견되었습니다. (결과는 정상 생성됨)")
                st.dataframe(warns, use_container_width=True)
            else:
                st.info("특이사항 없이 정상적으로 합산되었습니다.")

        except Exception as e:
            st.error(f"오류: {e}")
            st.exception(e)

# =========================================================================
# 탭2: 총괄표 채우기
# =========================================================================
def detect_layout(ws, max_scan_row=30, max_scan_col=30):
    region_like_count_in_col_a = 0
    for r in range(1, min(ws.max_row, max_scan_row) + 1):
        v = ws.cell(r, 1).value
        if is_valid_region(v):
            region_like_count_in_col_a += 1
    if region_like_count_in_col_a >= 3:
        return 'row'

    for r in range(1, min(ws.max_row, max_scan_row) + 1):
        hits = sum(
            1 for c in range(1, min(ws.max_column, max_scan_col) + 1)
            if is_valid_region(ws.cell(r, c).value)
        )
        if hits >= 3:
            return 'col'
    return None

def find_region_start_row(ws, target_keys, max_scan_row=120):
    max_row = min(ws.max_row, max_scan_row)
    for r in range(1, max_row + 1):
        label = ws.cell(r, 1).value
        key = region_key(label) if label else None
        if key in target_keys:
            return r
    return None

def get_base_group_size(base_ws, max_scan_row=120):
    max_row = min(base_ws.max_row, max_scan_row)
    region_rows = []
    for r in range(1, max_row + 1):
        label = base_ws.cell(r, 1).value
        if is_valid_region(label):
            region_rows.append(r)
    if len(region_rows) >= 2:
        return region_rows[1] - region_rows[0]
    return 1

def find_header_row(ws, target_keys=None, max_scan_row=30):
    for r in range(1, min(ws.max_row, max_scan_row) + 1):
        if target_keys:
            hits = sum(
                1 for c in range(1, ws.max_column + 1)
                if region_key(ws.cell(r, c).value) in target_keys
            )
            if hits >= 1:
                return r
        else:
            hits = sum(
                1 for c in range(1, ws.max_column + 1)
                if is_valid_region(ws.cell(r, c).value)
            )
            if hits >= 3:
                return r
    return None

def find_region_col_in_sheet(ws, header_row, target_keys):
    if header_row is None:
        return {}
    col_for_key = {}
    for c in range(1, ws.max_column + 1):
        label = ws.cell(header_row, c).value
        key = region_key(label) if label else None
        if key in target_keys:
            col_for_key.setdefault(key, c)
    return col_for_key

def fill_row_layout(base_ws, src_ws, own_key, warnings, sheet_title):
    target_keys = target_keys_for_region(own_key)
    if not target_keys:
        return 0

    max_col = base_ws.max_column
    group_size = get_base_group_size(base_ws)
    count = 0

    for key in target_keys:
        base_start = find_region_start_row(base_ws, {key})
        src_start = find_region_start_row(src_ws, {key})

        if base_start is None or src_start is None:
            continue

        for offset in range(group_size):
            gr_base = base_start + offset
            gr_src = src_start + offset
            for c in range(2, max_col + 1):
                base_cell = base_ws.cell(gr_base, c)
                if isinstance(base_cell, MergedCell):
                    continue
                if is_formula(base_cell.value):
                    continue
                
                src_val = src_ws.cell(gr_src, c).value
                src_val = get_safe_value(src_val)
                
                if src_val is None:
                    continue
                    
                base_cell.value = src_val
        count += 1
        
    return count

def fill_col_layout(base_ws, src_ws, own_key, warnings, sheet_title):
    target_keys = target_keys_for_region(own_key)
    if not target_keys:
        return 0

    base_header_row = find_header_row(base_ws)
    src_header_row = find_header_row(src_ws, target_keys=target_keys)

    if base_header_row is None or src_header_row is None:
        return 0

    base_col_for_key = find_region_col_in_sheet(base_ws, base_header_row, target_keys)
    src_col_for_key = find_region_col_in_sheet(src_ws, src_header_row, target_keys)

    max_row = min(base_ws.max_row, 120)
    count = 0
    for key in target_keys:
        base_col = base_col_for_key.get(key)
        src_col = src_col_for_key.get(key)
        if base_col is None or src_col is None:
            continue
            
        row_offset = src_header_row - base_header_row
        for r in range(base_header_row + 1, max_row + 1):
            base_cell = base_ws.cell(r, base_col)
            if isinstance(base_cell, MergedCell):
                continue
            if is_formula(base_cell.value):
                continue
                
            src_r = r + row_offset
            
            src_val = src_ws.cell(src_r, src_col).value if src_r <= src_ws.max_row else None
            src_val = get_safe_value(src_val)
            
            if src_val is None:
                continue
                
            base_cell.value = src_val
        count += 1
    return count

def is_standard_structure(src_ws, max_scan_row=120):
    max_row = min(src_ws.max_row, max_scan_row)
    region_row_count = 0
    for r in range(1, max_row + 1):
        if is_valid_region(src_ws.cell(r, 1).value):
            region_row_count += 1
    return region_row_count >= 5

def check_structure_mismatch(src_ws, sheet_title, fname, warnings):
    if not is_standard_structure(src_ws):
        warnings.append({
            "유형": "⚠ 양식 임의변경 감지", "시트": sheet_title, "셀": "-",
            "파일": fname,
            "설명": "이 파일은 자기 지역만 남기고 삭제되는 등 편집이 감지되었습니다. (데이터는 제자리에 정상 취합됩니다)"
        })

def fill_master_template(template_bytes, region_files_with_names):
    log = []
    warnings = []

    base_wb = openpyxl.load_workbook(template_bytes, data_only=False)
    sheet_count = len(base_wb.sheetnames)

    layouts = {}
    for idx in range(sheet_count):
        ws = get_sheet_by_index(base_wb, idx)
        layouts[idx] = detect_layout(ws)

    for fname, fbytes in region_files_with_names:
        own_key = extract_own_region_from_filename(fname)
        if not own_key:
            warnings.append({
                "유형": "지역 인식 실패", "시트": "-", "셀": "-",
                "설명": f"'{fname}' 파일명에서 지역명을 추출할 수 없어 건너뜀"
            })
            continue

        fbytes.seek(0)
        wb_values = openpyxl.load_workbook(fbytes, data_only=True)

        if len(wb_values.sheetnames) != sheet_count:
            warnings.append({
                "유형": "시트 개수 다름", "시트": "-", "셀": "-",
                "설명": f"'{fname}'은 시트 {len(wb_values.sheetnames)}개 (총괄표는 {sheet_count}개)"
            })

        first_sheet = get_sheet_by_index(wb_values, 0)
        if first_sheet is not None and layouts.get(0) == 'row':
            check_structure_mismatch(first_sheet, "전체 파일", fname, warnings)

        for idx in range(min(sheet_count, len(wb_values.sheetnames))):
            base_ws = get_sheet_by_index(base_wb, idx)
            src_ws = get_sheet_by_index(wb_values, idx)
            layout = layouts.get(idx)
            
            if layout == 'row':
                n = fill_row_layout(base_ws, src_ws, own_key, warnings, base_ws.title)
            elif layout == 'col':
                n = fill_col_layout(base_ws, src_ws, own_key, warnings, base_ws.title)
            else:
                n = 0
            if n:
                log.append({"파일": fname, "시트": base_ws.title, "방식": layout, "처리건수": n})

    return base_wb, log, warnings

with tab2:
    st.caption("시군구별로 작성된 여러 파일을 하나의 '총괄표' 서식에 자동으로 채워 넣습니다.")
    st.info("📌 파일명은 'NN_지역명_...' 형식이어야 합니다 (예: 01_포항시_...).")

    template_file = st.file_uploader("① 총괄표(서식) 파일 업로드", type=["xlsx"], key="template_up")
    region_files = st.file_uploader("② 시군구별 파일 업로드 (여러 개)", type=["xlsx"], accept_multiple_files=True, key="region_up")

    if template_file and region_files and st.button("🚀 총괄표 채우기 시작", key="btn2"):
        try:
            template_bytes = io.BytesIO(template_file.read())
            region_files_with_names = [(f.name, io.BytesIO(f.read())) for f in region_files]

            result_wb, log, warns = fill_master_template(template_bytes, region_files_with_names)

            o = io.BytesIO()
            result_wb.save(o)

            st.success("총괄표 채우기가 완료되었습니다.")
            st.download_button("📥 다운로드", o.getvalue(), "총괄표_결과.xlsx", key="dl2")

            if warns:
                st.warning(f"⚠️ 확인이 필요한 항목 {len(warns)}건이 발견되었습니다. (결과는 정상 생성됨)")
                st.dataframe(warns, use_container_width=True)
            else:
                st.info("특이사항 없이 정상적으로 채워졌습니다.")

            with st.expander("처리 로그 보기 (성공 내역)"):
                st.dataframe(log, use_container_width=True)

        except Exception as e:
            st.error(f"오류: {e}")
            st.exception(e)
