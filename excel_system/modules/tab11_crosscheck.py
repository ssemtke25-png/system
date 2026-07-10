# -*- coding: utf-8 -*-
"""
tab11_crosscheck.py — 정합성 검산 탭 (지적재조사)

두 가지 모드:
  ① 엑셀 ↔ 엑셀 : 지구명 키로 필지수·면적·사업비 비교 + 합계 재계산 대조
  ② 한글 ↔ 엑셀 : HWPX 실시계획 문서에서 지구별 값 추출 후 엑셀과 대조

설계 원칙
  - 숫자 비교는 100% 코드로만 수행 (AI 미사용 → 환각 원천 차단)
  - "추출 안 된 지구"는 반드시 별도 표시 (거짓 안심 방지)
  - 불일치 목록은 엑셀로 다운로드 가능

호출: main.py 에서  tab11_crosscheck.render()
"""

import io
import re
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st


# ══════════════════════════════════════════════════════════════
# 공통 유틸
# ══════════════════════════════════════════════════════════════
def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _clean_num(x):
    """'53,586' '158.0' 37 등에서 숫자만 추출 → float, 실패 시 None."""
    if x is None:
        return None
    if isinstance(x, (int, float)):
        try:
            if pd.isna(x):
                return None
        except Exception:
            pass
        return float(x)
    s = str(x).replace(",", "").strip()
    m = re.search(r"-?\d+\.?\d*", s)
    return float(m.group()) if m else None


def _norm_gu(name: str) -> str:
    """지구명 정규화: 공백 제거. '문경 각서지구'->'문경각서지구' 등 매칭 안정화."""
    if not isinstance(name, str):
        return ""
    return re.sub(r"\s+", "", name).strip()


# ══════════════════════════════════════════════════════════════
# HWPX 텍스트 추출
# ══════════════════════════════════════════════════════════════
def extract_hwpx_lines(file_bytes: bytes) -> list[str]:
    lines: list[str] = []
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
        secs = [n for n in zf.namelist()
                if re.search(r"section\d+\.xml$", n, re.IGNORECASE)]
        secs.sort(key=lambda n: int(re.search(r"section(\d+)", n, re.I).group(1)))
        for sec in secs:
            try:
                root = ET.fromstring(zf.read(sec))
            except Exception:
                continue
            for p in root.iter():
                if _localname(p.tag) != "p":
                    continue
                buf = [n.text for n in p.iter()
                       if _localname(n.tag) in ("t", "char") and n.text]
                line = "".join(buf).strip()
                if line:
                    lines.append(line)
    return lines


# 지구 블록별 필드 추출 정규식
_RE_GU = re.compile(r"사업지구명\s*[:：]\s*(\S+지구)")
_RE_AREA = re.compile(r"(\d[\d,]*)\s*필지[,\s]+([\d,]+\.?\d*)\s*㎡")
_RE_COST = re.compile(r"사업비\s*추산액\s*[:：]\s*([\d,]+)\s*백만원")
_RE_OWN = re.compile(r"소유자\s*동의율\s*[:：]\s*([\d.]+)\s*%\s*\(전체\s*([\d,]+)명\s*동의대상\s*([\d,]+)명\s*동의\s*([\d,]+)명")
_RE_AREARATE = re.compile(r"면적\s*동의율\s*[:：]\s*([\d.]+)\s*%\s*\(전체\s*([\d,]+)㎡\s*동의대상\s*([\d,]+)㎡\s*동의\s*([\d,]+)㎡")


def parse_hwpx_districts(lines: list[str]) -> dict:
    """
    한글 문서에서 지구별 필드 dict 추출.
    반환: { 정규화지구명: {"원본명":.., "필지":.., "면적":.., "사업비":..,
                          "소유자동의율":.., "면적동의율":.., ...} }
    필드가 없으면 해당 키는 None.
    """
    result: dict = {}
    cur = None
    for ln in lines:
        m = _RE_GU.search(ln)
        if m:
            raw = m.group(1)
            cur = _norm_gu(raw)
            result[cur] = {"원본명": raw, "필지": None, "면적": None,
                           "사업비": None, "소유자동의율": None, "면적동의율": None,
                           "동의인원": None, "동의면적": None}
            continue
        if not cur:
            continue
        rec = result[cur]
        a = _RE_AREA.search(ln)
        if a and rec["필지"] is None:
            rec["필지"] = _clean_num(a.group(1))
            rec["면적"] = _clean_num(a.group(2))
        c = _RE_COST.search(ln)
        if c and rec["사업비"] is None:
            rec["사업비"] = _clean_num(c.group(1))  # 백만원 단위
        o = _RE_OWN.search(ln)
        if o and rec["소유자동의율"] is None:
            rec["소유자동의율"] = _clean_num(o.group(1))
            rec["동의인원"] = _clean_num(o.group(4))
        ar = _RE_AREARATE.search(ln)
        if ar and rec["면적동의율"] is None:
            rec["면적동의율"] = _clean_num(ar.group(1))
            rec["동의면적"] = _clean_num(ar.group(4))
    return result


# ══════════════════════════════════════════════════════════════
# 엑셀 지구별 데이터 추출 (유연한 헤더 탐색)
# ══════════════════════════════════════════════════════════════
def guess_gu_column(df: pd.DataFrame) -> int | None:
    """'지구'로 끝나는 문자열이 가장 많은 열을 지구명 열로 추정."""
    best_col, best_cnt = None, 0
    for c in range(df.shape[1]):
        cnt = df[c].apply(lambda v: isinstance(v, str) and v.strip().endswith("지구")).sum()
        if cnt > best_cnt:
            best_col, best_cnt = c, cnt
    return best_col


def suggest_num_columns(df: pd.DataFrame, gu_col: int) -> list[int]:
    """지구명 열 오른쪽에서 숫자가 들어있는 열 후보 반환 (열 순서 유지)."""
    cols = []
    for c in range(gu_col + 1, df.shape[1]):
        cnt = df[c].apply(lambda v: _clean_num(v) is not None).sum()
        if cnt > 0:
            cols.append(c)
    return cols


def parse_excel_districts(df: pd.DataFrame, gu_col: int,
                          col_map: dict) -> dict:
    """
    사용자가 지정한 열 매핑으로 지구별 dict 생성.
    col_map: {"필지": 열idx or None, "면적": ..., "사업비": ...}
    """
    result = {}
    for i in range(df.shape[0]):
        gu = df.iloc[i, gu_col]
        if isinstance(gu, str) and gu.strip().endswith("지구"):
            key = _norm_gu(gu)
            rec = {"원본명": gu.strip()}
            for field, col in col_map.items():
                rec[field] = _clean_num(df.iloc[i, col]) if col is not None else None
            result[key] = rec
    return result


# ══════════════════════════════════════════════════════════════
# 대조 로직
# ══════════════════════════════════════════════════════════════
def compare_dicts(left: dict, right: dict, fields: list[str],
                  left_name="A", right_name="B", tol=0.5) -> pd.DataFrame:
    """
    두 지구 dict 를 필드별로 비교해 결과 DataFrame 반환.
    상태: 일치 / 불일치 / A에만 / B에만 / 값없음
    """
    rows = []
    all_keys = sorted(set(left) | set(right))
    for key in all_keys:
        L = left.get(key)
        R = right.get(key)
        disp = (L or R).get("원본명", key)
        if L and not R:
            rows.append({"지구": disp, "항목": "-", "상태": f"{right_name}에만 없음",
                         left_name: "존재", right_name: "없음", "차이": ""})
            continue
        if R and not L:
            rows.append({"지구": disp, "항목": "-", "상태": f"{left_name}에만 없음",
                         left_name: "없음", right_name: "존재", "차이": ""})
            continue
        for f in fields:
            lv = L.get(f)
            rv = R.get(f)
            if lv is None and rv is None:
                continue
            if lv is None or rv is None:
                rows.append({"지구": disp, "항목": f, "상태": "값없음",
                             left_name: "" if lv is None else lv,
                             right_name: "" if rv is None else rv, "차이": "한쪽 값 없음"})
                continue
            if abs(lv - rv) <= tol:
                rows.append({"지구": disp, "항목": f, "상태": "일치",
                             left_name: lv, right_name: rv, "차이": 0})
            else:
                rows.append({"지구": disp, "항목": f, "상태": "불일치",
                             left_name: lv, right_name: rv, "차이": round(lv - rv, 2)})
    return pd.DataFrame(rows)


def make_download_excel(df_mismatch: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_mismatch.to_excel(w, index=False, sheet_name="불일치목록")
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════
# 결과 렌더링 공통
# ══════════════════════════════════════════════════════════════
def _render_result(df: pd.DataFrame):
    if df.empty:
        st.warning("대조할 데이터를 찾지 못했습니다. 파일 구조를 확인해 주세요.")
        return

    problem = df[df["상태"] != "일치"]
    ok_cnt = (df["상태"] == "일치").sum()
    bad_cnt = (df["상태"] == "불일치").sum()
    miss_cnt = len(problem) - bad_cnt

    c1, c2, c3 = st.columns(3)
    c1.metric("✅ 일치", int(ok_cnt))
    c2.metric("❌ 불일치", int(bad_cnt))
    c3.metric("⚠️ 누락·값없음", int(miss_cnt))

    st.divider()

    if problem.empty:
        st.success("모든 항목이 일치합니다. ✅")
    else:
        st.subheader("⚠️ 확인이 필요한 항목")
        st.dataframe(problem, use_container_width=True, hide_index=True)
        st.download_button(
            "📥 불일치 목록 엑셀 다운로드",
            data=make_download_excel(problem),
            file_name="정합성_불일치목록.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with st.expander("전체 대조 결과 보기", expanded=False):
        st.dataframe(df, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════
# 탭 본체
# ══════════════════════════════════════════════════════════════
def render():
    st.header("🔍 정합성 검산")
    st.caption("숫자 대조는 100% 코드로만 수행합니다 (AI 미사용). 추출 안 된 지구는 별도 표시됩니다.")

    mode = st.radio("검산 종류", ["① 엑셀 ↔ 엑셀", "② 한글(HWPX) ↔ 엑셀"], horizontal=True)

    # ─────────────────────────── 모드 1
    if mode.startswith("①"):
        col1, col2 = st.columns(2)
        with col1:
            f1 = st.file_uploader("엑셀 A", type=["xlsx"], key="xa")
        with col2:
            f2 = st.file_uploader("엑셀 B", type=["xlsx"], key="xb")
        if not (f1 and f2):
            st.info("비교할 엑셀 두 개를 올려주세요.")
            return

        s1 = st.selectbox("A 시트", pd.ExcelFile(f1).sheet_names, key="s1")
        s2 = st.selectbox("B 시트", pd.ExcelFile(f2).sheet_names, key="s2")
        df1 = pd.read_excel(f1, sheet_name=s1, header=None)
        df2 = pd.read_excel(f2, sheet_name=s2, header=None)

        st.markdown("**A 파일 열 지정**")
        map1, gu1 = _column_mapper(df1, "A")
        st.markdown("**B 파일 열 지정**")
        map2, gu2 = _column_mapper(df2, "B")

        if not st.button("검산 실행", type="primary"):
            return
        d1 = parse_excel_districts(df1, gu1, map1)
        d2 = parse_excel_districts(df2, gu2, map2)
        fields = [f for f in map1 if map1[f] is not None or map2.get(f) is not None]
        df = compare_dicts(d1, d2, fields, left_name="엑셀A", right_name="엑셀B")
        _render_result(df)

    # ─────────────────────────── 모드 2
    else:
        col1, col2 = st.columns(2)
        with col1:
            hf = st.file_uploader("한글 HWPX", type=["hwpx"], key="hf")
        with col2:
            xf = st.file_uploader("엑셀", type=["xlsx"], key="xf")
        if not (hf and xf):
            st.info("한글(HWPX)과 엑셀을 각각 올려주세요.")
            return

        sx = st.selectbox("엑셀 시트", pd.ExcelFile(xf).sheet_names, key="sx")
        dfe = pd.read_excel(xf, sheet_name=sx, header=None)

        st.markdown("**엑셀 열 지정** (한글에서 뽑은 값과 맞출 열을 고르세요)")
        emap, egu = _column_mapper(dfe, "E")

        if not st.button("검산 실행", type="primary"):
            return

        h = parse_hwpx_districts(extract_hwpx_lines(hf.read()))
        e = parse_excel_districts(dfe, egu, emap)

        empty_gu = [v["원본명"] for v in h.values() if v.get("필지") is None]
        if empty_gu:
            st.warning(
                "한글 문서에서 **필지/면적이 추출되지 않은 지구**입니다. "
                "실제 누락인지, 표기 형식이 달라 못 읽은 건지 원본 확인이 필요합니다:\n\n- "
                + "\n- ".join(empty_gu)
            )

        fields = [f for f in emap if emap[f] is not None]
        if not fields:
            fields = ["필지", "면적"]
        df = compare_dicts(h, e, fields, left_name="한글", right_name="엑셀")
        _render_result(df)


def _column_mapper(df: pd.DataFrame, key_prefix: str):
    """미리보기 + 지구명/필지/면적/사업비 열을 사용자가 고르게 하는 UI."""
    with st.expander("데이터 미리보기 (상위 8행)", expanded=False):
        st.dataframe(df.head(8), use_container_width=True)

    ncol = df.shape[1]
    col_labels = [f"{i}열" for i in range(ncol)]
    none_label = "(없음)"

    gu_guess = guess_gu_column(df) or 0
    num_suggest = suggest_num_columns(df, gu_guess)

    def _sel(field, default_idx, key):
        options = [none_label] + col_labels
        idx = 0 if default_idx is None else default_idx + 1
        choice = st.selectbox(field, options, index=idx, key=f"{key_prefix}_{key}")
        return None if choice == none_label else col_labels.index(choice)

    c0, c1, c2, c3 = st.columns(4)
    with c0:
        gu = _sel("지구명 열", gu_guess, "gu")
    with c1:
        pil = _sel("필지수 열", num_suggest[0] if len(num_suggest) > 0 else None, "pil")
    with c2:
        area = _sel("면적 열", num_suggest[1] if len(num_suggest) > 1 else None, "area")
    with c3:
        cost = _sel("사업비 열", None, "cost")  # 기본 (없음) — 오탐 방지

    col_map = {"필지": pil, "면적": area}
    if cost is not None:
        col_map["사업비"] = cost
    return col_map, (gu if gu is not None else gu_guess)


if __name__ == "__main__":
    render()
