# -*- coding: utf-8 -*-
"""
탭15: 분기·연차 비교 시각화
이미 취합된 결과 파일 여러 개를 업로드 → 같은 지표를 기간별로 비교.

[설계 원리]
- 기간 라벨은 사용자가 자유 지정 (분기/연차 무관)
- 각 파일 → (기간,시군,지표,값) long-format → pivot
  ⇒ 파일이 2개든 12개든 코드 한 벌로 처리
- 시군명 정규화로 '포항시 남구'='포항남' 매칭 (표준목록 방식)

[입력] 취합 완료본 xlsx 여러 개 (탭2~14에서 나온 결과)
[출력] 지표별 비교표 + 추이 차트 + 엑셀 다운로드
"""
import io
import re
from datetime import datetime

import pandas as pd
import openpyxl
import streamlit as st


# ══════════════════════════════════════════════
# 시군명 정규화 (매칭 키)
# ══════════════════════════════════════════════
SIGUN_CANON = ['포항남구', '포항북구', '경주', '김천', '안동', '구미', '영주', '영천',
               '상주', '문경', '경산', '군위', '의성', '청송', '영양', '영덕', '청도',
               '고령', '성주', '칠곡', '예천', '봉화', '울진', '울릉']

# 표기 변형 → 표준키 (실무에서 발견되는 변형을 여기 추가)
SIGUN_ALIAS = {
    '포항시남구': '포항남구', '포항남': '포항남구', '포항시 남구': '포항남구',
    '포항시북구': '포항북구', '포항북': '포항북구', '포항시 북구': '포항북구',
}


def norm_sigun(raw):
    """'포항시 남구','포항남','포항 남구' → '포항남구' 로 통일.
    단순 문자제거는 '청송→송' 오류가 나므로 표준목록 매칭 방식 사용."""
    if not raw:
        return None
    s = re.sub(r'[\s\u3000]', '', str(raw))
    if s in SIGUN_ALIAS:
        return SIGUN_ALIAS[s]
    s2 = re.sub(r'(시|군)$', '', s)
    if s2 in SIGUN_ALIAS:
        return SIGUN_ALIAS[s2]
    for canon in SIGUN_CANON:
        if s2 == canon or s2 == canon.replace('구', ''):
            return canon
    return s2 or s


def num(v):
    return v if isinstance(v, (int, float)) else 0


# ══════════════════════════════════════════════
# 파일 → long-format 추출
# ══════════════════════════════════════════════
# 지표 정의: (시트명, 추출함수) — 취합 서식이 바뀌면 여기만 수정
def _extract_jongsaja(ws, period, rows):
    for r in range(5, 28):
        nm = ws.cell(r, 1).value
        if nm and nm != '계':
            sg = norm_sigun(nm)
            rows.append((period, sg, '소속공인중개사', num(ws.cell(r, 4).value)))
            rows.append((period, sg, '중개보조원', num(ws.cell(r, 5).value)))


def _extract_gaeeop(ws, period, rows):
    """각 시군은 4행 블록(중개사/중개인/법인/계). 시군명은 첫 행에만.
    '계' 행을 앵커로 찾아 총수(C)·신규(D)·폐업(E)·등록취소(F)·이전(G)을 읽고
    순증감 = 신규 - 폐업 - 등록취소 + 이전 을 산출.
    (단순 4행 고정점프는 빈 행/서식변형에 취약하므로 '계' 탐색 방식)"""
    r = 3
    maxr = ws.max_row
    while r <= maxr:
        nm = ws.cell(r, 1).value
        if nm and str(nm).strip() not in ('합계', '구분', ''):
            sg = norm_sigun(nm)
            gyr = None
            for k in range(5):                       # 앵커 아래 5행 내 '계' 탐색
                if str(ws.cell(r + k, 2).value).strip() == '계':
                    gyr = r + k
                    break
            if gyr:
                total = num(ws.cell(gyr, 3).value)
                singyu = num(ws.cell(gyr, 4).value)   # 신규등록
                pyeeop = num(ws.cell(gyr, 5).value)   # 폐업
                chwiso = num(ws.cell(gyr, 6).value)   # 등록취소
                ijeon = num(ws.cell(gyr, 7).value)    # 이전
                net = singyu - pyeeop - chwiso + ijeon
                rows.append((period, sg, '개업공인중개사', total))
                rows.append((period, sg, '신규등록', singyu))
                rows.append((period, sg, '폐업', pyeeop))
                rows.append((period, sg, '순증감', net))
                r = gyr + 1
                continue
        r += 1


def _extract_dansok(ws, period, rows):
    for r in range(7, 30):
        nm = ws.cell(r, 1).value
        if nm and nm != '계':
            rows.append((period, norm_sigun(nm), '단속업소수', num(ws.cell(r, 2).value)))


def _extract_viol(ws, period, rows, label):
    """위반유형 시트는 [근거조문|위반내용|계|시군별...] 구조.
    '계' 총계 행의 C열(3) 값이 곧 전체 건수 → 그 값을 직접 읽는다.
    (조문×시군 전체합 방식은 소계·중복행이 섞이면 이중합산되어 취약)"""
    tot = None
    for r in range(1, min(ws.max_row, 12) + 1):
        a = str(ws.cell(r, 1).value).strip() if ws.cell(r, 1).value else ''
        b = str(ws.cell(r, 2).value).strip() if ws.cell(r, 2).value else ''
        # 헤더(구분/근거조문) 다음의 '계' 총계행 찾기
        if a == '계' or b == '계':
            tot = num(ws.cell(r, 3).value)
            break
    if tot is None:                                  # '계' 라벨 없으면 헤더 다음행이 총계
        for r in range(4, min(ws.max_row, 12) + 1):
            v = ws.cell(r, 3).value
            if isinstance(v, (int, float)):
                tot = v
                break
    rows.append((period, '(전체)', f'위반-{label}', num(tot)))


SHEET_EXTRACTORS = [
    ('2.시도별 중개업 종사자수(분기)', _extract_jongsaja),
    ('3.개업공인중개사 증감내역', _extract_gaeeop),
    ('5.개업공인중개사 지도단속실적', _extract_dansok),
    ('6.고발조치', lambda ws, p, rows: _extract_viol(ws, p, rows, '고발조치')),
    ('7.등록취소', lambda ws, p, rows: _extract_viol(ws, p, rows, '등록취소')),
    ('8.업무정지', lambda ws, p, rows: _extract_viol(ws, p, rows, '업무정지')),
    ('9.과태료처분', lambda ws, p, rows: _extract_viol(ws, p, rows, '과태료')),
    ('10.자격취소, 자격정지(끝)', lambda ws, p, rows: _extract_viol(ws, p, rows, '자격취소·정지')),
]


def extract_long(file_bytes, period):
    """한 파일(바이트) → long-format rows [(기간, 시군, 지표, 값), ...]"""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    rows = []
    for sheet_name, fn in SHEET_EXTRACTORS:
        if sheet_name in wb.sheetnames:
            try:
                fn(wb[sheet_name], period, rows)
            except Exception as e:
                st.warning(f"[{period}] '{sheet_name}' 추출 중 오류: {e}")
    return rows


def guess_period(filename):
    """파일명에서 기간 라벨 추론. 실패하면 파일명 그대로."""
    name = filename.rsplit('.', 1)[0]
    # 2026_2분기, 2026년 2분기, 2026-2Q 등
    m = re.search(r'(20\d{2}).*?([1-4])\s*(?:/4)?\s*분기', name)
    if m:
        return f"{m.group(1)} {m.group(2)}Q"
    m = re.search(r'(20\d{2})[\s_\-]*[qQ]\s*([1-4])', name)
    if m:
        return f"{m.group(1)} {m.group(2)}Q"
    m = re.search(r'(20\d{2})\s*년', name)
    if m:
        return f"{m.group(1)}년"
    return name[:20]


# ══════════════════════════════════════════════
# 비교표 생성 + 엑셀 출력
# ══════════════════════════════════════════════
def build_pivot(df, indicator):
    """지표 하나 → pivot (행=시군/유형, 열=기간, +증감/증감률)"""
    sub = df[df['지표'] == indicator]
    if sub.empty:
        return None
    pv = sub.pivot_table(index='시군', columns='기간', values='값',
                         aggfunc='sum').fillna(0)
    pv = pv[sorted(pv.columns)]       # 기간 시간순
    cols = list(pv.columns)
    if len(cols) >= 2:
        pv['증감'] = pv[cols[-1]] - pv[cols[0]]
        pv['증감률(%)'] = pv.apply(
            lambda row: round((row[cols[-1]] - row[cols[0]]) / row[cols[0]] * 100, 1)
            if row[cols[0]] else 0.0, axis=1)
        pv = pv.sort_values(cols[-1], ascending=False)
    return pv


def to_excel(pivots):
    """지표별 pivot들 → 시각화 엑셀 (기존 양식과 동일 톤)"""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    NAVY, BLUE, GREY, WHITE = "1A2B44", "2E6FD0", "5B6D85", "FFFFFF"
    LINE, GREEN, RED = "D6DEEA", "1E8449", "C0392B"
    F = "맑은 고딕"
    thin = Side(style="thin", color=LINE)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for indicator, pv in pivots.items():
        if pv is None:
            continue
        safe = re.sub(r'[\\/*?:\[\]]', '', indicator)[:28]
        ws = wb.create_sheet(safe)
        ws["A1"] = f"{indicator} 기간 비교"
        ws["A1"].font = Font(name=F, size=15, bold=True, color=NAVY)

        # 헤더
        hrow = 3
        headers = ['구분'] + [str(c) for c in pv.columns]
        for j, h in enumerate(headers):
            c = ws.cell(hrow, 1 + j, h)
            c.font = Font(name=F, size=10.5, bold=True, color=WHITE)
            c.fill = PatternFill("solid", fgColor=NAVY)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border
        # 데이터
        period_cols = [c for c in pv.columns if c not in ('증감', '증감률(%)')]
        for i, (idx, row) in enumerate(pv.iterrows()):
            rr = hrow + 1 + i
            cc = ws.cell(rr, 1, idx)
            cc.font = Font(name=F, size=10); cc.border = border
            cc.alignment = Alignment(horizontal="left", vertical="center")
            for j, col in enumerate(pv.columns):
                v = row[col]
                cell = ws.cell(rr, 2 + j, round(v, 1) if isinstance(v, float) else v)
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col == '증감률(%)':
                    delta = row.get('증감', 0)
                    color = GREEN if delta > 0 else (RED if delta < 0 else GREY)
                    cell.font = Font(name=F, size=10, bold=True, color=color)
                    cell.number_format = "+0.0;-0.0;0.0"
                elif col == '증감':
                    color = GREEN if v > 0 else (RED if v < 0 else GREY)
                    cell.font = Font(name=F, size=10, bold=True, color=color)
                else:
                    cell.font = Font(name=F, size=10)
        lastrow = hrow + len(pv)
        ws.column_dimensions["A"].width = 14
        for c in range(2, len(headers) + 1):
            ws.column_dimensions[get_column_letter(c)].width = 12

        # 추이 차트 (기간 컬럼만)
        if len(period_cols) >= 2 and len(pv) <= 30:
            ch = BarChart(); ch.type = "col"
            ch.title = f"{indicator} 기간 추이"
            first = 2
            last = 1 + len(period_cols)
            data = Reference(ws, min_col=first, max_col=last,
                             min_row=hrow, max_row=lastrow)
            cats = Reference(ws, min_col=1, min_row=hrow + 1, max_row=lastrow)
            ch.add_data(data, titles_from_data=True)
            ch.set_categories(cats)
            ch.height = 10; ch.width = 24; ch.gapWidth = 60
            ws.add_chart(ch, f"A{lastrow + 3}")

        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════
# Streamlit UI
# ══════════════════════════════════════════════
def render():
    st.caption("취합 완료본 여러 개를 올리면 기간별로 비교합니다. (분기·연차 무관)")

    files = st.file_uploader(
        "취합 결과 파일 업로드 (여러 개)",
        type=["xlsx"], accept_multiple_files=True, key="cmp_files",
    )
    if not files:
        st.info("비교할 취합본을 2개 이상 올려주세요. 예: 2025_3분기.xlsx, 2026_1분기.xlsx")
        return

    # 각 파일에 기간 라벨 지정
    st.markdown("**① 각 파일의 기간 라벨 확인·수정**")
    periods = {}
    for i, f in enumerate(files):
        col1, col2 = st.columns([3, 2])
        col1.text(f"📄 {f.name}")
        default = guess_period(f.name)
        periods[i] = col2.text_input(
            "기간 라벨", value=default, key=f"cmp_period_{i}",
            label_visibility="collapsed",
        )

    if len({p.strip() for p in periods.values()}) < len(periods):
        st.warning("⚠ 기간 라벨이 중복됩니다. 서로 다르게 입력하세요.")

    if st.button("📊 비교 생성", type="primary", key="cmp_run"):
        # 전체 추출
        all_rows = []
        prog = st.progress(0.0)
        for i, f in enumerate(files):
            data = f.getvalue()
            all_rows += extract_long(data, periods[i].strip())
            prog.progress((i + 1) / len(files))
        prog.empty()

        if not all_rows:
            st.error("데이터를 추출하지 못했습니다. 취합 서식이 맞는지 확인하세요.")
            return

        df = pd.DataFrame(all_rows, columns=['기간', '시군', '지표', '값'])
        st.session_state["cmp_df"] = df

    # 결과 표시
    df = st.session_state.get("cmp_df")
    if df is None:
        return

    indicators = list(df['지표'].unique())
    st.markdown("**② 비교할 지표 선택**")
    chosen = st.multiselect(
        "지표", indicators, default=indicators[:4], key="cmp_ind",
    )
    if not chosen:
        st.info("지표를 하나 이상 선택하세요.")
        return

    pivots = {}
    for ind in chosen:
        pv = build_pivot(df, ind)
        pivots[ind] = pv
        if pv is not None:
            st.markdown(f"#### {ind}")
            st.dataframe(pv, use_container_width=True)

    # 엑셀 다운로드
    xlsx = to_excel({k: v for k, v in pivots.items() if v is not None})
    today = datetime.now().strftime("%Y%m%d")
    st.download_button(
        "📥 비교 엑셀 다운로드 (차트 포함)",
        data=xlsx,
        file_name=f"분기비교_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", key="cmp_dl",
    )
