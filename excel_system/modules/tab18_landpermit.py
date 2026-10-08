"""
탭18(메뉴 ⑰): 토지거래허가 분기보고 취합
시군에서 받은 '토지거래계약 허가 현황' 파일들을 빈 서식에 붙여 취합본 1개로 만든다.

· 결과 파일 구성 : 맨 앞 '취합' 시트 + 올린 시군만 직제순(포항시 남구 → … → 울릉군)으로 시트 1장씩
· '취합' 합계 수식은 실제로 붙은 시군 시트만 더하도록 매번 새로 만든다
· 빈 서식(tab18_landpermit_template.xlsx)에는 숫자가 하나도 없다 (취합 / 서식 2장)
· 시군 파일은 스타일을 무시하고 셀 값만 직접 읽는다
  (다른 프로그램으로 저장돼 엑셀 라이브러리로 안 열리는 파일 대비)
· 검증 : 지역 인식 실패, 같은 지역 중복, 숫자 아닌 값(예: '해당없음') 0 처리,
         시군이 직접 입력한 합계와 세부값 불일치, 본기 > 누계, 분기 표시 불일치
"""
import io
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import streamlit as st
from openpyxl.utils import column_index_from_string, get_column_letter

TEMPLATE_PATH = Path(__file__).resolve().parent / "tab18_landpermit_template.xlsx"

# 직제순 : (지역키, 시트 이름, 공식 명칭)
ORDER = [
    ("포항남", "포항남구", "포항시 남구"), ("포항북", "포항북구", "포항시 북구"),
    ("경주", "경주", "경주시"), ("김천", "김천", "김천시"), ("안동", "안동", "안동시"),
    ("구미", "구미", "구미시"), ("영주", "영주", "영주시"), ("영천", "영천", "영천시"),
    ("상주", "상주", "상주시"), ("문경", "문경", "문경시"), ("경산", "경산", "경산시"),
    ("의성", "의성", "의성군"), ("청송", "청송", "청송군"), ("영양", "영양", "영양군"),
    ("영덕", "영덕", "영덕군"), ("청도", "청도", "청도군"), ("고령", "고령", "고령군"),
    ("성주", "성주", "성주군"), ("칠곡", "칠곡", "칠곡군"), ("예천", "예천", "예천군"),
    ("봉화", "봉화", "봉화군"), ("울진", "울진", "울진군"), ("울릉", "울릉", "울릉군"),
]
KEY_INFO = {k: (i, sheet, full) for i, (k, sheet, full) in enumerate(ORDER)}

PERIOD_LABELS = ("본기", "당월", "당분기", "금분기", "이번분기")
CUM_LABELS = ("누계",)
NOTE_COL = "Z"   # 비고 칸 (글자 그대로 복사)

# 시군이 합계 칸에 숫자를 직접 적어 보낸 경우 검산할 관계식 (서식의 수식과 같음)
CHECKS = [
    ("K", ("N", "Q", "T"), "불허가 소계 건수"), ("L", ("O", "R", "U"), "불허가 소계 필지수"),
    ("M", ("P", "S", "V"), "불허가 소계 면적"),
    ("E", ("H", "K"), "허가현황 소계 건수"), ("F", ("I", "L"), "허가현황 소계 필지수"),
    ("G", ("J", "M"), "허가현황 소계 면적"),
    ("B", ("E", "W"), "신청 건수"), ("C", ("F", "X"), "신청 필지수"), ("D", ("G", "Y"), "신청 면적"),
]

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ── 시군 파일 읽기 (스타일 무시, 첫 시트 셀 값만) ──────────────────────
def read_first_sheet(data: bytes):
    """반환: {셀주소: (값, 수식여부)}"""
    z = zipfile.ZipFile(io.BytesIO(data))
    names = z.namelist()
    shared = []
    if "xl/sharedStrings.xml" in names:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall(f"{{{_NS}}}si"):
            shared.append("".join(t.text or "" for t in si.iter(f"{{{_NS}}}t")))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    first = wb.find(f"{{{_NS}}}sheets/{{{_NS}}}sheet")
    rid = first.get(f"{{{_RNS}}}id")
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    target = next(r.get("Target") for r in rels if r.get("Id") == rid).lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target
    sheet = ET.fromstring(z.read(target))
    cells = {}
    for c in sheet.iter(f"{{{_NS}}}c"):
        ref, typ = c.get("r"), c.get("t")
        v, f = c.find(f"{{{_NS}}}v"), c.find(f"{{{_NS}}}f")
        if typ == "s" and v is not None:
            val = shared[int(v.text)]
        elif typ == "inlineStr":
            val = "".join(x.text or "" for x in c.iter(f"{{{_NS}}}t"))
        elif v is not None and v.text is not None:
            try:
                num = float(v.text)
                val = int(num) if num.is_integer() else num
            except ValueError:
                val = v.text
        else:
            val = None
        if val is not None:
            cells[ref] = (val, f is not None)
    return cells


def _to_number(v):
    """숫자로 바꿀 수 있으면 숫자, 아니면 None"""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(",", "")
    if s in ("", "-"):
        return 0
    try:
        n = float(s)
        return int(n) if n.is_integer() else n
    except ValueError:
        return None


def _find_rows(cells):
    period = cum = None
    for ref, (val, _) in cells.items():
        m = re.fullmatch(r"A(\d+)", ref)
        if not m or not isinstance(val, str):
            continue
        label = re.sub(r"\s+", "", val)
        if period is None and label in PERIOD_LABELS:
            period = int(m.group(1))
        elif cum is None and label in CUM_LABELS:
            cum = int(m.group(1))
    return period, cum


def _quarter(text):
    m = re.search(r"(\d{4})\s*년\s*(\d)\s*(?:/\s*4)?\s*분기", str(text or ""))
    return (int(m.group(1)), int(m.group(2))) if m else None


def _find_key(text):
    """글자 속에서 시군을 찾는다. 포항은 남구/북구까지 있어야 인식.
    반환: 지역키 / '포항'(구 구분 불가) / None"""
    t = re.sub(r"\s+", "", str(text or ""))
    if "포항" in t:
        if "남구" in t:
            return "포항남"
        if "북구" in t:
            return "포항북"
        return "포항"
    for key, _, _ in ORDER:
        if not key.startswith("포항") and key in t:
            return key
    return None


def _region(fname, title):
    """파일명 → 제목(A1) 순서로 시군을 찾는다.
    (파일명이 '2. 토지거래허가 분기보고_경주'처럼 번호 뒤에 시군명이 바로 오지 않아도 인식)"""
    k1 = _find_key(Path(fname).stem)
    if k1 in KEY_INFO:
        return k1
    k2 = _find_key(title)
    if k2 in KEY_INFO:
        return k2
    return None


# ── 취합본 만들기 ─────────────────────────────────────────────────────
def build(files, year, quarter):
    """files: [(파일명, bytes)] → (결과 bytes, 처리목록, 경고목록)"""
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    form = wb["서식"]
    total = wb["취합"]
    isf = lambda v: isinstance(v, str) and v.startswith("=")

    # 서식의 입력칸(수식이 아닌 칸) 목록 : (행, 열문자)
    form_period, form_cum = 7, 8
    input_cols = [get_column_letter(c) for c in range(2, 27)
                  if not isf(form.cell(form_period, c).value)]

    log, warns, used = [], [], {}
    parsed = []
    for fname, data in files:
        try:
            cells = read_first_sheet(data)
        except Exception as e:
            warns.append({"유형": "파일 읽기 실패", "파일": fname, "설명": f"엑셀 파일로 읽을 수 없음 ({e})"})
            continue
        title = cells.get("A1", ("", False))[0]
        key = _region(fname, title)
        if key is None:
            warns.append({"유형": "지역 인식 실패", "파일": fname,
                          "설명": "파일명·제목에서 시군(포항은 남구/북구까지)을 찾지 못해 제외"})
            continue
        if key in used:
            warns.append({"유형": "같은 지역 중복", "파일": fname,
                          "설명": f"{KEY_INFO[key][2]} 파일이 이미 있음({used[key]}) → 이 파일은 제외"})
            continue
        used[key] = fname
        parsed.append((KEY_INFO[key][0], key, fname, cells, title))

    parsed.sort(key=lambda x: x[0])   # 직제순
    sheet_names = []
    for _, key, fname, cells, title in parsed:
        _, sname, full = KEY_INFO[key]
        ws = wb.copy_worksheet(form)
        ws.title = sname
        sheet_names.append(sname)
        foreign = "외국인" in str(title)
        ws["A1"] = f"{'외국인 ' if foreign else ''}토지거래계약 허가 현황({full}  {year}년  {quarter}분기)"

        q = _quarter(title)
        if q and q != (year, quarter):
            warns.append({"유형": "분기 표시 다름", "파일": fname,
                          "설명": f"파일 제목은 {q[0]}년 {q[1]}분기 (선택한 분기: {year}년 {quarter}분기) — 지난 파일인지 확인"})

        src_period, src_cum = _find_rows(cells)
        if src_period is None or src_cum is None:
            warns.append({"유형": "양식 다름", "파일": fname,
                          "설명": "'본기(당월)' 또는 '누계' 줄을 찾지 못해 숫자를 넣지 못함"})
            log.append({"순서": len(sheet_names), "시군": full, "시트": sname, "파일": fname, "채운 칸": 0})
            continue

        filled, texts, got = 0, set(), {}
        for src_row, dst_row in ((src_period, form_period), (src_cum, form_cum)):
            for col in input_cols:
                raw = cells.get(f"{col}{src_row}")
                if raw is None:
                    continue
                val, _ = raw
                if col == NOTE_COL:
                    ws[f"{col}{dst_row}"] = val
                    continue
                num = _to_number(val)
                if num is None:
                    texts.add(str(val).strip())
                    num = 0
                ws[f"{col}{dst_row}"] = num
                got[(col, dst_row)] = num
                filled += 1
        log.append({"순서": len(sheet_names), "시군": full, "시트": sname, "파일": fname, "채운 칸": filled})

        if texts:
            # 한 칸에 한 글자씩 적은 경우(예: 해·당·없·음)도 원래 문구로 보여주기 위해 줄 전체 글자를 모은다
            row_txt = []
            for src_row in (src_period, src_cum):
                s = "".join(str(cells[f"{get_column_letter(c)}{src_row}"][0]).strip()
                            for c in range(2, 26)
                            if f"{get_column_letter(c)}{src_row}" in cells
                            and _to_number(cells[f"{get_column_letter(c)}{src_row}"][0]) is None)
                if s and s not in row_txt:
                    row_txt.append(s)
            warns.append({"유형": "숫자 아닌 값", "파일": fname,
                          "설명": f"숫자 칸에 '{', '.join(row_txt)}'(이)라고 적혀 있어 0으로 처리"})

        # 시군이 합계 칸에 직접 적은 숫자 검산
        for src_row, dst_row, label in ((src_period, form_period, "본기"), (src_cum, form_cum, "누계")):
            calc = {}
            def val_of(col):
                if col in calc:
                    return calc[col]
                return got.get((col, dst_row), 0)
            for tgt, parts, name in CHECKS:
                calc[tgt] = sum(val_of(p) for p in parts)
                raw = cells.get(f"{tgt}{src_row}")
                if raw and not raw[1]:
                    typed = _to_number(raw[0])
                    if typed is not None and typed != calc[tgt]:
                        warns.append({"유형": "합계 불일치", "파일": fname,
                                      "설명": f"[{label}] {name}: 파일에 적힌 값 {typed:,} ≠ 세부값 합 {calc[tgt]:,} (차이 {typed - calc[tgt]:,}) — 취합본은 세부값 기준으로 계산됨"})

        # 본기 > 누계
        over = [col for col in input_cols if col != NOTE_COL
                and got.get((col, form_period), 0) > got.get((col, form_cum), 0)]
        if over:
            warns.append({"유형": "본기 > 누계", "파일": fname,
                          "설명": f"이번 분기 값이 누계보다 큼: {', '.join(over)}열 — 누계 확인 필요"})

    # 취합 시트 : 붙은 시군 시트만 더하는 수식
    for r in (form_period, form_cum):
        for col in input_cols:
            if col == NOTE_COL:
                continue
            refs = [f"'{s}'!{col}{r}" for s in sheet_names]
            total[f"{col}{r}"] = "=" + "+".join(refs) if refs else 0
    total["A1"] = f"토지거래계약 허가 현황(경북 {year}년  {quarter}분기)"

    del wb["서식"]
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), log, warns


# ── 화면 ──────────────────────────────────────────────────────────────
def render():
    st.caption("시군에서 받은 토지거래계약 허가 현황 파일을 올리면, 빈 서식에서 취합본을 새로 만듭니다.")
    st.info("📌 결과는 맨 앞 '취합' 시트 + 올린 시군만 직제순(포항시 남구 → … → 울릉군)으로 붙습니다. "
            "파일명이나 제목에 시군명이 있으면 자동 인식합니다(포항은 남구·북구까지).")

    if not TEMPLATE_PATH.exists():
        st.error("빈 서식(tab18_landpermit_template.xlsx)을 modules 폴더에 넣어주세요.")
        return

    files = st.file_uploader("시군 파일 업로드 (여러 개)", type=["xlsx"],
                             accept_multiple_files=True, key="t18_files")

    # 파일 제목에서 연도·분기를 찾아 기본값으로
    det = {}
    for f in files or []:
        try:
            q = _quarter(read_first_sheet(f.getvalue()).get("A1", ("", False))[0])
            if q:
                det[q] = det.get(q, 0) + 1
        except Exception:
            pass
    dy, dq = max(det, key=det.get) if det else (2026, 1)
    c1, c2 = st.columns(2)
    year = c1.number_input("연도", 2020, 2100, dy, key="t18_year")
    quarter = c2.selectbox("분기", [1, 2, 3, 4], index=dq - 1, key="t18_q")

    if files and st.button("🚀 취합본 만들기", key="t18_go"):
        out, log, warns = build([(f.name, f.getvalue()) for f in files], int(year), int(quarter))
        st.session_state["t18_result"] = (out, log, warns, int(year), int(quarter))

    res = st.session_state.get("t18_result")
    if res:
        out, log, warns, y, q = res
        st.success(f"취합본을 만들었습니다. 시군 {len(log)}곳 · 맨 앞 '취합' 시트 포함 {len(log) + 1}장")
        st.download_button("📥 취합본 다운로드", out,
                           f"{y}년 {q}분기 토지거래허가 분기보고서(취합).xlsx", key="t18_dl")
        if warns:
            st.warning(f"⚠️ 확인이 필요한 항목 {len(warns)}건 (결과는 정상 생성됨)")
            st.dataframe(warns, use_container_width=True)
        else:
            st.info("특이사항 없이 정상적으로 취합되었습니다.")
        with st.expander("처리 내역 (시트 순서)"):
            st.dataframe(log, use_container_width=True)
