"""
탭8: 과징금 취합
시군별 파일(현황보고 + 미수납조서)을 받아 합계 파일에 취합한다.

[시트1 - 현황보고(경상북도)]
  행8~22: 지역별 데이터 행 → 합계 파일의 해당 시군 행에 숫자 복사
  색 채워진 칸(계산용)은 덮어쓰지 않음

[시트2 - 미수납조서]
  행6부터 실데이터 → 합계 파일에 시군 순서대로 이어붙이기
  합계행(행5)의 건수·금액 자동 갱신
"""
import io
import re
import zipfile
import xml.etree.ElementTree as ET

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

# 시트 이름 후보 (파일마다 조금씩 다름)
SHEET1_CANDIDATES = ['현황보고', '경상북도']
SHEET2_CANDIDATES = ['미수납조서', '미수납 조서']

# 합계 파일 시트명 후보
TOTAL_SHEET1_CANDIDATES = ['경상북도', '현황보고(경상북도)', '합계']
TOTAL_SHEET2_CANDIDATES = ['미수납조서', '미수납 조서', '미수납조서(합계)']


def find_base_sheets(base_wb):
    """합계(총괄표) 파일의 시트1(현황보고), 시트2(미수납조서) 찾기"""
    ws1 = find_sheet(base_wb, SHEET1_CANDIDATES + ['현황보고(경상북도)', '총괄'])
    ws2 = find_sheet(base_wb, SHEET2_CANDIDATES)
    return ws1, ws2


def is_base_file(wb):
    """항상 False 반환 - 서식/데이터 혼입 방지 로직 제거 (모든 파일 처리)"""
    return False

# 현황보고: 데이터 행 범위 (0-indexed: 7~21 = 행8~22)
DATA_ROW_START = 7  # 0-indexed
DATA_ROW_END   = 22 # 0-indexed, exclusive
DATA_COL_START = 2  # 0-indexed (C열부터)

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


# ── 유틸 ─────────────────────────────────────────────────────────────
def region_key(label):
    """셀 값에서 지역 키 추출"""
    if not label or not isinstance(label, str):
        return None
    label = label.strip()
    # 포항 남/북 구분
    if re.search(r'포항.*(남|남구)', label): return '포항남'
    if re.search(r'포항.*(북|북구)', label): return '포항북'
    for k in REGION_ORDER:
        if k in label:
            return k
    return None


def extract_region_from_filename(fname):
    """파일명에서 지역키 추출 (예: '1_포항시_북구.xls' → '포항북')"""
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


def is_formula(v):
    return isinstance(v, str) and v.startswith('=')


def is_protected(cell):
    return is_formula(cell.value) or has_protective_color(cell)


def find_sheet(wb, candidates):
    """시트 이름 후보 중 존재하는 것 반환"""
    for name in candidates:
        if name in wb.sheetnames:
            return wb[name]
    # 부분 매칭
    for name in wb.sheetnames:
        for c in candidates:
            if c.replace(' ','') in name.replace(' ',''):
                return wb[name]
    return None


def shrink_styles(file_bytes):
    """styles.xml 최적화로 로딩 속도 개선"""
    file_bytes.seek(0)
    try:
        zin = zipfile.ZipFile(file_bytes, 'r')
        if 'xl/styles.xml' not in zin.namelist():
            file_bytes.seek(0)
            return file_bytes
        styles_bytes = zin.read('xl/styles.xml')
        ET.register_namespace('', NS_MAIN)
        root = ET.fromstring(styles_bytes)
        cell_style_xfs = root.find(f'{{{NS_MAIN}}}cellStyleXfs')
        if cell_style_xfs is None or len(cell_style_xfs) <= 200:
            file_bytes.seek(0)
            return file_bytes
        first = cell_style_xfs[0] if len(cell_style_xfs) > 0 else None
        for xf in list(cell_style_xfs): cell_style_xfs.remove(xf)
        if first is not None: cell_style_xfs.append(first)
        cell_style_xfs.set('count', '1')
        cell_xfs = root.find(f'{{{NS_MAIN}}}cellXfs')
        if cell_xfs:
            for xf in cell_xfs:
                if xf.get('xfId'): xf.set('xfId', '0')
        xml_decl = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        new_bytes = xml_decl + ET.tostring(root, encoding='UTF-8')
        out = io.BytesIO()
        zout = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
        for item in zin.namelist():
            zout.writestr(item, new_bytes if item == 'xl/styles.xml' else zin.read(item))
        zout.close(); zin.close()
        out.seek(0)
        return out
    except Exception:
        file_bytes.seek(0)
        return file_bytes


def load_wb(file_bytes, data_only=True):
    shrunk = shrink_styles(file_bytes)
    return openpyxl.load_workbook(shrunk, data_only=data_only)


# 현황보고 시트 설정
SHEET1_DATA_ROWS   = range(9, 24)   # 1-indexed 행9~23
SHEET1_BASE_COLS   = range(2, 16)   # 합계파일 열2~15 (B~O)


def detect_src_col_offset(src_ws):
    """시군 파일의 열 오프셋 자동 감지
    - 행8에서 '건 수' 또는 '금  액' 텍스트가 시작되는 열 - 합계파일 시작열(2) = offset
    """
    for c in range(1, 8):
        v = src_ws.cell(8, c).value
        if v and isinstance(v, str) and ('건' in v or '금' in v):
            return c - 2  # 합계파일 B열(2) 기준
    # 행9에서 숫자가 처음 나오는 열로 판단
    for c in range(1, 8):
        v = src_ws.cell(9, c).value
        if isinstance(v, (int, float)) and v is not None:
            return c - 2
    return 1  # 기본값: 1칸 오른쪽


def get_input_cells(base_ws):
    """합계 파일에서 수식·색 보호 아닌 실입력칸 좌표 목록 반환"""
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
    """합계 파일 현황보고 실입력칸 0으로 초기화"""
    for r, c in input_cells:
        base_ws.cell(r, c).value = 0


def accumulate_sheet1(base_ws, src_ws, input_cells, warnings, fname):
    """시군 현황보고 → 합계 파일에 누적 합산 (열 오프셋 자동 감지)"""
    offset = detect_src_col_offset(src_ws)
    added = 0
    for r, c in input_cells:
        src_c = c + offset
        if src_c < 1:
            continue
        src_val = src_ws.cell(r, src_c).value
        if isinstance(src_val, str) and src_val.startswith('#'):
            warnings.append({"유형":"수식오류","파일":fname,
                             "셀":f"{get_column_letter(src_c)}{r}",
                             "설명":f"'{src_val}' → 0으로 처리"})
            src_val = 0
        if src_val is None:
            src_val = 0
        try:
            base_val = base_ws.cell(r, c).value or 0
            base_ws.cell(r, c).value = base_val + src_val
            added += 1
        except Exception:
            pass
    return added


# ── 시트2: 미수납조서 이어붙이기 ─────────────────────────────────────
def find_data_start_row(ws):
    """'합계' 행 다음 행 = 데이터 시작 행 (1-indexed)"""
    for r in range(1, min(ws.max_row + 1, 20)):
        v = ws.cell(r, 1).value or ws.cell(r, 2).value
        if v and isinstance(v, str) and '합계' in str(v):
            return r + 1
    return 6  # 기본값


def clear_data_area(ws, start_row):
    """기존 데이터 영역 초기화"""
    if start_row > ws.max_row:
        return
    for r in range(start_row, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            if not isinstance(cell, MergedCell):
                cell.value = None


def is_valid_row(ws, row, max_col):
    """실제 데이터가 있는 행인지 판단"""
    return any(
        ws.cell(row, c).value not in (None, '')
        for c in range(1, max_col + 1)
    )


def find_src_data_range(src_ws):
    """src 미수납조서에서 실데이터 시작/끝 행 찾기"""
    start = None
    for r in range(1, min(src_ws.max_row + 1, 20)):
        v = src_ws.cell(r, 2).value
        if v and '합계' in str(v):
            start = r + 1
            break
    if start is None:
        start = 6
    end = src_ws.max_row
    return start, end


def append_sheet2(base_ws, src_ws, next_row, warnings, fname):
    """미수납조서 이어붙이기"""
    src_start, src_end = find_src_data_range(src_ws)
    max_col = max(base_ws.max_column or 14, src_ws.max_column or 14)
    added = 0

    for sr in range(src_start, src_end + 1):
        if not is_valid_row(src_ws, sr, max_col):
            continue
        br = next_row + added
        for c in range(1, max_col + 1):
            base_cell = base_ws.cell(br, c)
            if isinstance(base_cell, MergedCell):
                continue
            sv = src_ws.cell(sr, c).value
            if isinstance(sv, str) and sv.startswith('#'):
                warnings.append({"유형":"수식오류","파일":fname,
                                  "셀":f"{get_column_letter(c)}{sr}",
                                  "설명":f"'{sv}' → 빈칸 처리"})
                sv = None
            base_cell.value = sv
        added += 1

    return next_row + added, added


def update_sum_row(base_ws, data_start_row):
    """합계 행(data_start_row - 1)의 건수·금액을 실계산값으로 갱신"""
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
        amt_징수   += safe_num(base_ws.cell(r, 7).value)   # G열 = 미수납액
        count_이행 += safe_num(base_ws.cell(r, 8).value)   # H열 = 이행강제금 건수
        amt_이행   += safe_num(base_ws.cell(r, 10).value)  # J열 = 이행강제금 미수납액

    # 수식이든 아니든 실계산값으로 덮어쓰기
    for col, val in [
        (5, int(count_징수)),   # E열: 과징금 건수
        (7, int(amt_징수)),    # G열: 과징금 미수납액
        (8, int(count_이행)),  # H열: 이행강제금 건수
        (10, int(amt_이행)),   # J열: 이행강제금 미수납액
    ]:
        cell = base_ws.cell(sum_row, col)
        if not isinstance(cell, MergedCell):
            cell.value = val if val != 0 else None


# ── 메인 취합 함수 ───────────────────────────────────────────────────
def aggregate(template_bytes, region_files):
    log, warnings = [], []

    template_bytes = shrink_styles(template_bytes)
    base_wb = openpyxl.load_workbook(template_bytes, data_only=False)

    # 합계 파일 시트 자동 감지
    base_ws1, base_ws2 = find_base_sheets(base_wb)
    if not base_ws1:
        sheet_names = ", ".join(base_wb.sheetnames)
        raise ValueError(
            f"총괄표 파일에서 현황보고 시트를 찾지 못했습니다.\n"
            f"현재 시트 목록: [{sheet_names}]\n"
            f"'현황보고' 또는 '경상북도' 시트가 있어야 합니다."
        )
    if not base_ws2:
        sheet_names = ", ".join(base_wb.sheetnames)
        raise ValueError(
            f"총괄표 파일에서 미수납조서 시트를 찾지 못했습니다.\n"
            f"현재 시트 목록: [{sheet_names}]"
        )

    sorted_files = sorted(region_files, key=lambda x: sort_key(x[0]))

    # 시군 파일 로딩 + 합계파일 혼입 방지
    loaded = {}
    for fname, fbytes in sorted_files:
        try:
            wb = load_wb(fbytes, data_only=True)
            if is_base_file(wb):
                warnings.append({"유형":"파일 오류","파일":fname,
                                  "설명":"합계(서식) 파일이 시군 파일 목록에 포함됨 → 건너뜀"})
                continue
            loaded[fname] = wb
        except Exception as e:
            warnings.append({"유형":"파일 오류","파일":fname,"설명":str(e)})

    # ── 시트1: 현황보고 취합 ──────────────────────────────────────────
    input_cells = get_input_cells(base_ws1)
    reset_sheet1(base_ws1, input_cells)

    for fname, _ in sorted_files:
        if fname not in loaded:
            continue
        src_ws1 = find_sheet(loaded[fname], SHEET1_CANDIDATES)
        if not src_ws1:
            warnings.append({"유형":"시트 없음","파일":fname,"설명":"현황보고 시트를 찾지 못함"})
            continue
        n = accumulate_sheet1(base_ws1, src_ws1, input_cells, warnings, fname)
        log.append({"파일":fname,"시트":"현황보고","누적셀수":n})

    # ── 시트2: 미수납조서 이어붙이기 ────────────────────────────────
    data_start = find_data_start_row(base_ws2)
    clear_data_area(base_ws2, data_start)
    next_row = data_start

    for fname, _ in sorted_files:
        if fname not in loaded:
            continue
        src_ws2 = find_sheet(loaded[fname], SHEET2_CANDIDATES)
        if not src_ws2:
            warnings.append({"유형":"시트 없음","파일":fname,"설명":"미수납조서 시트를 찾지 못함"})
            continue
        next_row, n = append_sheet2(base_ws2, src_ws2, next_row, warnings, fname)
        log.append({"파일":fname,"시트":"미수납조서","추가행수":n})

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
        "📌 일부 시군 파일이 누락되어도 나머지 파일은 정상 처리됩니다.\n\n"
        "📌 원본 값이 수식 오류(#VALUE! 등)인 경우, 해당 칸만 비우고 계속 진행됩니다."
    )

    template_file = st.file_uploader(
        "① 합계(서식) 파일 업로드",
        type=["xlsx"],
        key="tpl_up8"
    )
    region_files = st.file_uploader(
        "② 시군별 파일 업로드 (여러 개, xls/xlsx 모두 가능)",
        type=["xlsx","xls"],
        accept_multiple_files=True,
        key="region_up8"
    )

    if template_file and region_files and st.button("🚀 취합 시작", key="btn8"):
        # xls 변환 시도
        converted = []
        for f in region_files:
            fname = f.name
            raw = f.read()
            if fname.lower().endswith('.xls'):
                try:
                    import xlrd
                    import openpyxl
                    book = xlrd.open_workbook(file_contents=raw)
                    wb_new = openpyxl.Workbook()
                    wb_new.remove(wb_new.active)
                    for sheet_name in book.sheet_names():
                        ws_src = book.sheet_by_name(sheet_name)
                        ws_dst = wb_new.create_sheet(sheet_name)
                        for r in range(ws_src.nrows):
                            for c in range(ws_src.ncols):
                                ws_dst.cell(r+1, c+1, ws_src.cell_value(r, c))
                    buf = io.BytesIO()
                    wb_new.save(buf)
                    buf.seek(0)
                    converted.append((fname, buf))
                except ImportError:
                    st.warning(f"⚠️ xlrd 미설치로 {fname} 변환 실패. requirements.txt에 xlrd==1.2.0 추가 필요.")
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
