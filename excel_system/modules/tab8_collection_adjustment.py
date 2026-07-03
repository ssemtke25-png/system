"""
탭8: 과징금 취합 (버그 수정본 v2)

수정사항:
1. has_protective_color - theme색 보호셀 오판 수정
2. SHEET1_BASE_COLS - 열3~16으로 수정
3. accumulate_sheet1 - offset 제거, 동일 열 직접 참조
4. 시트2 셀 서식·열너비·행높이까지 원본 그대로 복사
5. 이행강제금 시트 없는 파일 → 과징금만 처리
6. [신규] shrink_styles에서 docProps/custom.xml의 name=None 오류 파일 처리
   (한글오피스/특수 저장 파일에서 발생하는 openpyxl 로딩 오류 방지)
"""
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from copy import copy

import openpyxl
import streamlit as st
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

# ── 지역 설정 ────────────────────────────────────────────────────────
REGION_ORDER = [
    '포항남','포항북','경주','김천','안동','구미','영주','영천','상주','문경',
    '경산','군위','의성','청송','영양','영덕','청도','고령','성주','칠곡',
    '예천','봉화','울진','울릉',
]

SHEET1_CANDIDATES = ['현황보고', '경상북도']
SHEET2_CANDIDATES = ['미수납조서', '미수납 조서']

TOTAL_SHEET1_CANDIDATES = ['경상북도', '현황보고(경상북도)', '합계', '현황보고']
TOTAL_SHEET2_CANDIDATES = ['미수납조서', '미수납 조서', '미수납조서(합계)']

SHEET1_DATA_ROWS = range(9, 24)
SHEET1_BASE_COLS = range(3, 17)

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"


# ── 유틸 ─────────────────────────────────────────────────────────────
def region_key(label):
    if not label or not isinstance(label, str):
        return None
    label = label.strip()
    if re.search(r'포항.*(남|남구)', label): return '포항남'
    if re.search(r'포항.*(북|북구)', label): return '포항북'
    for k in REGION_ORDER:
        if k in label:
            return k
    return None


def extract_region_from_filename(fname):
    m = re.match(r'^\d{0,3}[_.\s]*(.+?)\.', fname)
    raw = m.group(1) if m else fname
    key = region_key(raw)
    if key:
        return key
    for k in REGION_ORDER:
        if k and k in fname:
            return k
    return None


def sort_key(fname):
    m = re.match(r'^(\d{1,3})[_.\s]', fname)
    if m:
        return int(m.group(1))
    key = extract_region_from_filename(fname)
    if key in REGION_ORDER:
        return 100 + REGION_ORDER.index(key)
    return 999


def has_protective_color(cell):
    fill = cell.fill
    if not getattr(fill, 'patternType', None):
        return False
    if fill.patternType == 'none':
        return False
    fg = getattr(fill, 'fgColor', None)
    if not fg:
        return False
    fg_type = getattr(fg, 'type', None)
    if fg_type == 'theme':
        return False
    if fg_type == 'rgb':
        try:
            rgb = getattr(fg, 'rgb', None)
            if rgb not in (None, '00000000', 'FFFFFFFF'):
                return True
        except Exception:
            return False
    return False


def is_formula(v):
    return isinstance(v, str) and v.startswith('=')


def is_protected(cell):
    return is_formula(cell.value) or has_protective_color(cell)


def find_sheet(wb, candidates):
    for name in candidates:
        if name in wb.sheetnames:
            return wb[name]
    for name in wb.sheetnames:
        for c in candidates:
            if c.replace(' ', '') in name.replace(' ', ''):
                return wb[name]
    return None


def find_base_sheets(base_wb):
    ws1 = find_sheet(base_wb, TOTAL_SHEET1_CANDIDATES)
    ws2 = find_sheet(base_wb, TOTAL_SHEET2_CANDIDATES)
    return ws1, ws2


def is_base_file(wb):
    return False


def fix_custom_props(zin, zout):
    """
    docProps/custom.xml에서 name 속성 없는 <property> 제거
    → 한글오피스 저장 파일에서 openpyxl TypeError 방지
    """
    CUSTOM_PATH = 'docProps/custom.xml'
    if CUSTOM_PATH not in zin.namelist():
        return False  # 파일 없으면 처리 불필요

    try:
        raw = zin.read(CUSTOM_PATH)
        ET.register_namespace('', NS_CUSTOM)
        ET.register_namespace('vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes')
        root = ET.fromstring(raw)

        to_remove = [p for p in root if not p.get('name')]
        if not to_remove:
            return False  # 문제 없으면 그대로

        for p in to_remove:
            root.remove(p)

        xml_decl = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        new_bytes = xml_decl + ET.tostring(root, encoding='UTF-8')
        zout.writestr(CUSTOM_PATH, new_bytes)
        return True
    except Exception:
        return False


def shrink_styles(file_bytes):
    """styles.xml 최적화 + custom.xml name=None 오류 수정"""
    file_bytes.seek(0)
    try:
        zin = zipfile.ZipFile(file_bytes, 'r')
        needs_custom_fix = 'docProps/custom.xml' in zin.namelist()

        # styles.xml 처리
        styles_changed = False
        new_styles_bytes = None
        if 'xl/styles.xml' in zin.namelist():
            styles_bytes = zin.read('xl/styles.xml')
            ET.register_namespace('', NS_MAIN)
            root = ET.fromstring(styles_bytes)
            cell_style_xfs = root.find(f'{{{NS_MAIN}}}cellStyleXfs')
            if cell_style_xfs is not None and len(cell_style_xfs) > 200:
                first = cell_style_xfs[0] if len(cell_style_xfs) > 0 else None
                for xf in list(cell_style_xfs): cell_style_xfs.remove(xf)
                if first is not None: cell_style_xfs.append(first)
                cell_style_xfs.set('count', '1')
                cell_xfs = root.find(f'{{{NS_MAIN}}}cellXfs')
                if cell_xfs:
                    for xf in cell_xfs:
                        if xf.get('xfId'): xf.set('xfId', '0')
                xml_decl = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                new_styles_bytes = xml_decl + ET.tostring(root, encoding='UTF-8')
                styles_changed = True

        if not styles_changed and not needs_custom_fix:
            file_bytes.seek(0)
            return file_bytes

        out = io.BytesIO()
        zout = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)

        custom_written = False
        for item in zin.namelist():
            if item == 'xl/styles.xml' and styles_changed:
                zout.writestr(item, new_styles_bytes)
            elif item == 'docProps/custom.xml':
                # custom.xml name=None 수정 시도
                fixed = fix_custom_props(zin, zout)
                if not fixed:
                    zout.writestr(item, zin.read(item))
                custom_written = True
            else:
                zout.writestr(item, zin.read(item))

        zout.close()
        zin.close()
        out.seek(0)
        return out
    except Exception:
        file_bytes.seek(0)
        return file_bytes


def load_wb(file_bytes, data_only=True):
    shrunk = shrink_styles(file_bytes)
    return openpyxl.load_workbook(shrunk, data_only=data_only)


# ── 시트1: 현황보고 취합 ─────────────────────────────────────────────
def get_input_cells(base_ws):
    cells = []
    for r in SHEET1_DATA_ROWS:
        for c in SHEET1_BASE_COLS:
            cell = base_ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            if not is_protected(cell):
                cells.append((r, c))
    return cells


def reset_sheet1(base_ws, input_cells):
    for r, c in input_cells:
        base_ws.cell(r, c).value = 0


def accumulate_sheet1(base_ws, src_ws, input_cells, warnings, fname):
    added = 0
    for r, c in input_cells:
        src_val = src_ws.cell(r, c).value
        if isinstance(src_val, str) and src_val.startswith('#'):
            warnings.append({"유형": "수식오류", "파일": fname,
                             "셀": f"{get_column_letter(c)}{r}",
                             "설명": f"'{src_val}' → 0으로 처리"})
            src_val = 0
        if src_val is None:
            src_val = 0
        try:
            if not isinstance(src_val, (int, float)):
                src_val = 0
            base_val = base_ws.cell(r, c).value or 0
            base_ws.cell(r, c).value = base_val + src_val
            added += 1
        except Exception:
            pass
    return added


# ── 시트2: 미수납조서 이어붙이기 (서식 포함 완전 복사) ───────────────
def copy_cell_style(src_cell, dst_cell):
    dst_cell.value = src_cell.value
    try:
        if src_cell.has_style:
            dst_cell.font      = copy(src_cell.font)
            dst_cell.fill      = copy(src_cell.fill)
            dst_cell.border    = copy(src_cell.border)
            dst_cell.alignment = copy(src_cell.alignment)
            dst_cell.number_format = src_cell.number_format
    except Exception:
        pass


def copy_col_widths(src_ws, dst_ws):
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            dst_ws.column_dimensions[col_letter].width = dim.width


def find_data_start_row(ws):
    for r in range(1, min(ws.max_row + 1, 20)):
        v = ws.cell(r, 1).value or ws.cell(r, 2).value
        if v and isinstance(v, str) and '합계' in str(v):
            return r + 1
    return 7


def clear_data_area(ws, start_row):
    if start_row > ws.max_row:
        return
    for r in range(start_row, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            if not isinstance(cell, MergedCell):
                cell.value = None


def is_valid_row(ws, row, max_col):
    return any(
        ws.cell(row, c).value not in (None, '')
        for c in range(1, max_col + 1)
    )


def find_src_data_range(src_ws):
    start = None
    for r in range(1, min(src_ws.max_row + 1, 20)):
        v = src_ws.cell(r, 2).value
        if v and '합계' in str(v):
            start = r + 1
            break
    if start is None:
        start = 7
    end = src_ws.max_row
    return start, end


def append_sheet2(base_ws, src_ws, next_row, warnings, fname):
    src_start, src_end = find_src_data_range(src_ws)
    max_col = max(base_ws.max_column or 14, src_ws.max_column or 14)

    copy_col_widths(src_ws, base_ws)

    added = 0
    for sr in range(src_start, src_end + 1):
        if not is_valid_row(src_ws, sr, max_col):
            continue
        br = next_row + added

        src_row_height = src_ws.row_dimensions[sr].height
        if src_row_height:
            base_ws.row_dimensions[br].height = src_row_height

        for c in range(1, max_col + 1):
            base_cell = base_ws.cell(br, c)
            if isinstance(base_cell, MergedCell):
                continue
            src_cell = src_ws.cell(sr, c)
            if isinstance(src_cell, MergedCell):
                continue
            sv = src_cell.value
            if isinstance(sv, str) and sv.startswith('#'):
                warnings.append({"유형": "수식오류", "파일": fname,
                                  "셀": f"{get_column_letter(c)}{sr}",
                                  "설명": f"'{sv}' → 빈칸 처리"})
                base_cell.value = None
            else:
                copy_cell_style(src_cell, base_cell)

        added += 1

    return next_row + added, added


def update_sum_row(base_ws, data_start_row):
    sum_row = data_start_row - 1
    if sum_row < 1:
        return

    count_징수 = 0
    amt_징수   = 0
    count_이행 = 0
    amt_이행   = 0

    def safe_num(v):
        if v is None or v == '' or v == '-':
            return 0
        try:
            return float(v)
        except (ValueError, TypeError):
            return 0

    for r in range(data_start_row, base_ws.max_row + 1):
        if not is_valid_row(base_ws, r, 14):
            continue
        count_징수 += safe_num(base_ws.cell(r, 5).value)
        amt_징수   += safe_num(base_ws.cell(r, 7).value)
        count_이행 += safe_num(base_ws.cell(r, 8).value)
        amt_이행   += safe_num(base_ws.cell(r, 10).value)

    for col, val in [
        (5, int(count_징수)),
        (7, int(amt_징수)),
        (8, int(count_이행)),
        (10, int(amt_이행)),
    ]:
        cell = base_ws.cell(sum_row, col)
        if not isinstance(cell, MergedCell):
            cell.value = val if val != 0 else None


# ── 메인 취합 함수 ───────────────────────────────────────────────────
def aggregate(template_bytes, region_files):
    log, warnings = [], []

    template_bytes = shrink_styles(template_bytes)
    base_wb = openpyxl.load_workbook(template_bytes, data_only=False)

    base_ws1, base_ws2 = find_base_sheets(base_wb)
    if not base_ws1:
        sheet_names = ", ".join(base_wb.sheetnames)
        raise ValueError(
            f"총괄표 파일에서 현황보고 시트를 찾지 못했습니다.\n"
            f"현재 시트 목록: [{sheet_names}]"
        )
    if not base_ws2:
        sheet_names = ", ".join(base_wb.sheetnames)
        raise ValueError(
            f"총괄표 파일에서 미수납조서 시트를 찾지 못했습니다.\n"
            f"현재 시트 목록: [{sheet_names}]"
        )

    sorted_files = sorted(region_files, key=lambda x: sort_key(x[0]))

    loaded = {}
    for fname, fbytes in sorted_files:
        try:
            wb = load_wb(fbytes, data_only=True)
            if is_base_file(wb):
                warnings.append({"유형": "파일 오류", "파일": fname,
                                  "설명": "합계(서식) 파일이 시군 파일 목록에 포함됨 → 건너뜀"})
                continue
            loaded[fname] = wb
        except Exception as e:
            warnings.append({"유형": "파일 오류", "파일": fname, "설명": str(e)})

    # 시트1: 현황보고 취합
    input_cells = get_input_cells(base_ws1)
    reset_sheet1(base_ws1, input_cells)

    for fname, _ in sorted_files:
        if fname not in loaded:
            continue
        src_ws1 = find_sheet(loaded[fname], SHEET1_CANDIDATES)
        if not src_ws1:
            warnings.append({"유형": "시트 없음", "파일": fname, "설명": "현황보고 시트를 찾지 못함"})
            continue
        n = accumulate_sheet1(base_ws1, src_ws1, input_cells, warnings, fname)
        log.append({"파일": fname, "시트": "현황보고", "누적셀수": n})

    # 시트2: 미수납조서 이어붙이기
    data_start = find_data_start_row(base_ws2)
    clear_data_area(base_ws2, data_start)
    next_row = data_start

    for fname, _ in sorted_files:
        if fname not in loaded:
            continue
        src_ws2 = find_sheet(loaded[fname], SHEET2_CANDIDATES)
        if not src_ws2:
            log.append({"파일": fname, "시트": "미수납조서", "추가행수": 0,
                        "비고": "미수납조서 시트 없음 (이행강제금 없는 파일)"})
            continue
        next_row, n = append_sheet2(base_ws2, src_ws2, next_row, warnings, fname)
        log.append({"파일": fname, "시트": "미수납조서", "추가행수": n})

    update_sum_row(base_ws2, data_start)

    return base_wb, log, warnings


# ── Streamlit UI ─────────────────────────────────────────────────────
def render():
    st.caption(
        "시군별 과징금 파일을 받아 합계 파일에 취합합니다. "
        "현황보고는 지역 행을 찾아 채우고, 미수납조서는 시군 순서대로 이어붙입니다."
    )
    st.info(
        "📌 셀에 색이 채워진 칸(계산용)은 수식과 마찬가지로 절대 덮어쓰지 않습니다.\n\n"
        "📌 미수납조서가 없는 시군 파일(이행강제금 없음)은 현황보고만 취합합니다.\n\n"
        "📌 원본 값이 수식 오류(#VALUE! 등)인 경우, 해당 칸만 비우고 계속 진행됩니다."
    )

    template_file = st.file_uploader(
        "① 합계(서식) 파일 업로드",
        type=["xlsx"],
        key="tpl_up8"
    )
    region_files = st.file_uploader(
        "② 시군별 파일 업로드 (여러 개, xls/xlsx 모두 가능)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="region_up8"
    )

    if template_file and region_files and st.button("🚀 취합 시작", key="btn8"):
        converted = []
        for f in region_files:
            fname = f.name
            raw = f.read()
            if fname.lower().endswith('.xls'):
                try:
                    import xlrd
                    book = xlrd.open_workbook(file_contents=raw)
                    wb_new = openpyxl.Workbook()
                    wb_new.remove(wb_new.active)
                    for sheet_name in book.sheet_names():
                        ws_src = book.sheet_by_name(sheet_name)
                        ws_dst = wb_new.create_sheet(sheet_name)
                        for r in range(ws_src.nrows):
                            for c in range(ws_src.ncols):
                                ws_dst.cell(r + 1, c + 1, ws_src.cell_value(r, c))
                    buf = io.BytesIO()
                    wb_new.save(buf)
                    buf.seek(0)
                    converted.append((fname, buf))
                except ImportError:
                    st.warning(f"⚠️ xlrd 미설치로 {fname} 변환 실패.")
                    converted.append((fname, io.BytesIO(raw)))
                except Exception as e:
                    st.warning(f"⚠️ {fname} 변환 오류: {e}")
                    converted.append((fname, io.BytesIO(raw)))
            else:
                converted.append((fname, io.BytesIO(raw)))

        try:
            tpl_bytes = io.BytesIO(template_file.read())
            result_wb, log, warns = aggregate(tpl_bytes, converted)

            out = io.BytesIO()
            result_wb.save(out)

            st.success("✅ 취합이 완료되었습니다.")
            st.download_button(
                "📥 취합 결과 다운로드",
                data=out.getvalue(),
                file_name="과징금_취합결과.xlsx",
                key="dl8"
            )

            if warns:
                st.warning(f"⚠️ 확인 필요 항목 {len(warns)}건 (취합 결과는 정상 생성됨)")
                st.dataframe(warns, use_container_width=True)
            else:
                st.info("특이사항 없이 정상 취합되었습니다.")

            with st.expander("처리 로그 보기"):
                st.dataframe(log, use_container_width=True)

        except ValueError as e:
            st.error(f"❌ 파일 오류: {e}")
            st.warning("💡 ① 총괄표(빈 서식) 파일을 업로드하고, ② 시군 파일들을 업로드하세요.")
        except Exception as e:
            st.error(f"오류: {e}")
            st.exception(e)
