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
def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def is_safe_to_copy(v):
    """SUM 등 수식 계산에 안전하게 들어갈 수 있는 값인지
    (#VALUE! 같은 에러 문자열이나 '-' 같은 빈값 표기는 위험하다고 판단)"""
    if v is None:
        return True
    if is_number(v):
        return True
    if isinstance(v, str):
        if v.startswith('#'):
            return False
        if v.strip() == '-':
            return False
        return True
    return True


# =========================================================================
# 탭1: 단순 합산 (여러 파일의 숫자를 모두 더함)
# =========================================================================
def aggregate(file_bytes_list, file_names):
    """
    동작 원칙:
    - base(첫 파일)에서 셀이 수식이면 절대 건드리지 않고 그대로 둔다.
    - 수식이 아닌 순수 데이터 칸만 모든 파일의 숫자를 더해서 채운다.
    - 시트는 "모든 파일에 등장하는 시트의 합집합" 기준으로 처리한다.
    """
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

    for sheet in all_sheets:
        present_idx = [i for i in range(n_files) if sheet in formula_wbs[i].sheetnames]
        if not present_idx:
            continue

        max_r = max(formula_wbs[i][sheet].max_row for i in present_idx)
        max_c = max(formula_wbs[i][sheet].max_column for i in present_idx)

        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                cell_addr = f"{get_column_letter(c)}{r}"
                formula_flags = []

                for i in present_idx:
                    fws = formula_wbs[i][sheet]
                    if r > fws.max_row or c > fws.max_column:
                        continue
                    fv = fws.cell(r, c).value
                    formula_flags.append((names[i], is_formula(fv)))

                n_considered = len(formula_flags)
                n_formula = sum(1 for _, isf in formula_flags if isf)

                if n_considered >= 2 and n_formula >= (n_considered / 2) and n_formula < n_considered:
                    non_formula_files = [fn for fn, isf in formula_flags if not isf]
                    warnings.append({
                        "유형": "수식 누락", "시트": sheet, "셀": cell_addr,
                        "파일": ", ".join(non_formula_files),
                        "설명": f"전체 {n_considered}개 파일 중 {n_formula}개가 수식을 사용 중인데 "
                                + ", ".join(non_formula_files)
                                + "에는 수식이 없음 (해당 파일 값은 그대로 합산에 반영됨)"
                    })

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
            st.caption("※ 합계 등 수식이 있던 칸은 수식 그대로 보존됩니다. Excel에서 파일을 열면 자동으로 재계산됩니다.")

            if warns:
                st.warning(f"⚠️ 확인이 필요한 항목 {len(warns)}건이 발견되었습니다. (합산 결과는 정상적으로 생성되었습니다)")
                st.dataframe(warns, use_container_width=True)
            else:
                st.info("특이사항 없이 정상적으로 합산되었습니다.")

        except Exception as e:
            st.error(f"오류: {e}")
            st.exception(e)


# =========================================================================
# 탭2: 총괄표 채우기 (시군구별 자료를 정해진 총괄표의 자기 자리에 채움)
# =========================================================================

# 지역명 표기 차이(포항시 남구/포항남구/포항남 등)를 통일하기 위한 정규화 키.
# ⚠ 이 목록은 "경상북도 22개 시군" 기준입니다. 다른 지역/다른 양식의 총괄표를 쓰려면
#   아래 VALID_REGION_KEYS와 PREFIX_SPECIAL을 그 지역에 맞게 바꿔야 합니다.
PREFIX_SPECIAL = {
    '포항시남구': '포항남', '포항남구': '포항남', '포항남': '포항남',
    '포항시북구': '포항북', '포항북구': '포항북', '포항북': '포항북',
}
VALID_REGION_KEYS = {
    '포항남', '포항북', '경주', '김천', '안동', '구미', '영주', '영천', '상주', '문경',
    '경산', '의성', '청송', '영양', '영덕', '청도', '고령', '성주', '칠곡', '예천',
    '봉화', '울진', '울릉'
}


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
    """파일명에서 지역명 추출. 숫자.지역명, 숫자_지역명, 지역명 단독 등 유연하게 처리."""
    # 1. '숫자.지역명' 또는 '숫자_지역명' 패턴 시도 (예: 11.의성군.xlsx, 01_포항시.xlsx)
    m = re.match(r'^\d{1,3}[_.\s]+([가-힣]+)', filename)
    if m:
        key = region_key(m.group(1))
        if key:
            return key
            
    # 2. 정규식에 안 맞더라도 파일명 전체에서 키워드 직접 검색 (예: "의성군_최종.xlsx")
    # '포항남구' 같은 특수 표기를 먼저 검사
    for raw_name, mapped_key in PREFIX_SPECIAL.items():
        if raw_name in filename.replace(' ', ''):
            return mapped_key
            
    # 일반 지역명 검사
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


def detect_layout(ws, max_scan_row=30, max_scan_col=30):
    """시트가 '지역이 행에 나열'인지 '지역이 열에 나열'인지 자동 판별."""
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
    """워크시트에서 target_keys에 해당하는 지역명이 처음 나오는 행 번호를 찾음."""
    max_row = min(ws.max_row, max_scan_row)
    for r in range(1, max_row + 1):
        label = ws.cell(r, 1).value
        key = region_key(label) if label else None
        if key in target_keys:
            return r
    return None


def is_standard_structure(src_ws, max_scan_row=120):
    """src 파일이 표준 구조(전체 지역을 모두 나열)인지 판별.
    유효 지역명이 적힌 행이 5개 미만이면 '자기 지역만 압축해서 작성한 비표준 파일'로 간주."""
    max_row = min(src_ws.max_row, max_scan_row)
    region_row_count = 0
    for r in range(1, max_row + 1):
        if is_valid_region(src_ws.cell(r, 1).value):
            region_row_count += 1
    return region_row_count >= 5


def check_structure_mismatch(src_ws, sheet_title, fname, warnings):
    """src 파일 자체가 표준 구조(전체 지역 나열)와 다르면 경고.
    이런 파일은 자기 지역 행/그룹은 찾아 복사하지만, 행 길이나 칸 배치가
    총괄표가 기대하는 표준 구조와 달라 일부 항목이 누락될 수 있다."""
    if not is_standard_structure(src_ws):
        warnings.append({
            "유형": "⚠ 표준 구조와 다름", "시트": sheet_title, "셀": "-",
            "파일": fname,
            "설명": "이 파일은 전체 지역을 나열하는 표준 양식과 다르게, 자기 지역만 압축해서 "
                    "작성되어 있습니다. 자기 지역 칸은 찾아서 복사했지만, 칸 배치가 표준과 달라 "
                    "일부 세부 항목이 누락될 수 있으니 결과를 직접 확인해 주세요."
        })


def get_base_group_size(base_ws, max_scan_row=120):
    """base 시트에서 지역 하나가 차지하는 고정 행 수를 계산 (지역명이 적힌 행들 간의 간격).
    시트1,2,4,5처럼 지역당 1행인 시트와, 시트3처럼 지역당 4행인 시트를 자동으로 구분."""
    max_row = min(base_ws.max_row, max_scan_row)
    region_rows = []
    for r in range(1, max_row + 1):
        label = base_ws.cell(r, 1).value
        if is_valid_region(label):
            region_rows.append(r)
    if len(region_rows) >= 2:
        return region_rows[1] - region_rows[0]
    return 1


def fill_row_layout(base_ws, src_ws, own_key, warnings, sheet_title):
    """지역이 행에 나열된 시트: 자기 지역 행(들)을 그대로 복사."""
    target_keys = target_keys_for_region(own_key)
    if not target_keys:
        return 0

    max_col = base_ws.max_column
    group_size = get_base_group_size(base_ws)
    count = 0

    # 포항시처럼 target_keys가 2개(포항남, 포항북)인 경우를 위해 각각 독립적으로 처리
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
                if src_val is None:
                    continue
                    
                if not is_safe_to_copy(src_val):
                    warnings.append({
                        "유형": "비정상 값 건너뜀", "시트": sheet_title,
                        "셀": f"{get_column_letter(c)}{gr_base}",
                        "설명": f"원본 값 '{src_val}'이 숫자 계산에 부적합하여 건너뜀"
                    })
                    continue
                    
                base_cell.value = src_val
        count += 1
        
    return count


def find_header_row(ws, target_keys=None, max_scan_row=30):
    """
    워크시트에서 헤더 행을 찾음.
    - target_keys가 주어지면 (개별 파일): 해당 지역명이 1개라도 포함된 행을 찾음
    - 주어지지 않으면 (총괄표): 지역명이 3개 이상 나열된 표준 헤더 행을 찾음
    """
    for r in range(1, min(ws.max_row, max_scan_row) + 1):
        if target_keys:
            hits = sum(
                1 for c in range(1, ws.max_column + 1)
                if region_key(ws.cell(r, c).value) in target_keys
            )
            if hits >= 1:  # 자기 지역 1개만 있어도 헤더로 인정!
                return r
        else:
            hits = sum(
                1 for c in range(1, ws.max_column + 1)
                if is_valid_region(ws.cell(r, c).value)
            )
            if hits >= 3:  # 총괄표는 3개 이상 있어야 표준 헤더로 인정
                return r
    return None


def find_region_col_in_sheet(ws, header_row, target_keys):
    """헤더 행에서 target_keys에 해당하는 지역명이 있는 열 번호를 찾음."""
    if header_row is None:
        return {}
    col_for_key = {}
    for c in range(1, ws.max_column + 1):
        label = ws.cell(header_row, c).value
        key = region_key(label) if label else None
        if key in target_keys:
            col_for_key.setdefault(key, c)
    return col_for_key


def fill_col_layout(base_ws, src_ws, own_key, warnings, sheet_title):
    """지역이 열에 나열된 시트: 자기 지역 열을 그대로 복사."""
    target_keys = target_keys_for_region(own_key)
    if not target_keys:
        return 0

    # 총괄표(base)는 지역명 3개 이상인 줄을 찾고, 개별파일(src)은 자기 지역명만 찾도록 수정!
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
            
            # [최적화 포인트] openpyxl은 빈 셀을 조회하면 메모리에 새 셀을 강제 생성하므로,
            # 원본 파일의 실제 행 범위를 넘어가는 조회는 아예 차단하여 메모리 초과(다운) 방지
            src_val = src_ws.cell(src_r, src_col).value if src_r <= src_ws.max_row else None
            
            if src_val is None:
                continue
            if not is_safe_to_copy(src_val):
                warnings.append({
                    "유형": "비정상 값 건너뜀", "시트": sheet_title,
                    "셀": f"{get_column_letter(base_col)}{r}",
                    "설명": f"원본 값 '{src_val}'이 숫자 계산에 부적합하여 건너뜀"
                })
                continue
            base_cell.value = src_val
        count += 1
    return count


def fill_master_template(template_bytes, region_files_with_names):
    """
    template_bytes: 총괄표(BytesIO)
    region_files_with_names: [(파일명, BytesIO), ...]

    규칙:
    - 모든 시트에 동일하게 적용: 시트별로 '지역이 행' 또는 '지역이 열'인지 자동 판별
    - 자기 지역에 해당하는 행(그룹) 또는 열을 그대로 복사
    - 총괄표 셀이 수식이면 절대 덮어쓰지 않음 (작성금지/계산 칸 보호)
    - 시트명이 달라도 시트 순서(인덱스)로 매칭 (시트 개수가 같으면 그대로 진행)
    """
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
                "설명": f"'{fname}'은 시트 {len(wb_values.sheetnames)}개 (총괄표는 {sheet_count}개) - 시트 순서대로 처리"
            })

        # 파일 전체에 대해 한 번만 표준 구조 여부 판단 (시트1 기준)
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
    st.caption(
        "시군구별로 작성된 여러 파일을 하나의 '총괄표' 서식에 자동으로 채워 넣습니다. "
        "총괄표에서 수식이 들어간 칸(계산 칸, 작성금지 칸)은 절대 건드리지 않고 보존합니다."
    )
    st.info(
        "📌 파일명은 'NN_지역명_...' 형식이어야 합니다 (예: 01_포항시_...).\n\n"
        "📌 현재는 경상북도 22개 시군 기준으로 지역명을 인식합니다. 다른 지역 데이터에 쓰려면 코드의 지역 목록을 수정해야 합니다."
    )

    template_file = st.file_uploader("① 총괄표(서식) 파일 업로드", type=["xlsx"], key="template_up")
    region_files = st.file_uploader(
        "② 시군구별 파일 업로드 (여러 개)", type=["xlsx"], accept_multiple_files=True, key="region_up"
    )

    if template_file and region_files and st.button("🚀 총괄표 채우기 시작", key="btn2"):
        try:
            template_bytes = io.BytesIO(template_file.read())
            region_files_with_names = [(f.name, io.BytesIO(f.read())) for f in region_files]

            result_wb, log, warns = fill_master_template(template_bytes, region_files_with_names)

            o = io.BytesIO()
            result_wb.save(o)

            st.success("총괄표 채우기가 완료되었습니다.")
            st.download_button("📥 다운로드", o.getvalue(), "총괄표_결과.xlsx", key="dl2")
            st.caption("※ 총괄표에서 수식이었던 칸은 그대로 보존됩니다. Excel에서 파일을 열면 자동으로 재계산됩니다.")

            if warns:
                st.warning(f"⚠️ 확인이 필요한 항목 {len(warns)}건이 발견되었습니다. (결과는 정상적으로 생성되었습니다)")
                st.dataframe(warns, use_container_width=True)
            else:
                st.info("특이사항 없이 정상적으로 채워졌습니다.")

            with st.expander("처리 로그 보기"):
                st.dataframe(log, use_container_width=True)

        except Exception as e:
            st.error(f"오류: {e}")
            st.exception(e)
