# -*- coding: utf-8 -*-
"""
tab16_pressrelease.py  ―  전국 참고 기반 보도자료 AI 생성 탭
======================================================================
데이터 취합 및 AI 자동화 시스템 / 16번 탭

설계 원칙
----------
1) 팩트(숫자·지구수·필지·면적·예산·시군명 등)는 '업로드한 결과파일'에서만 추출한다.
   ─ 뉴스 클리핑의 수치는 절대 본문 팩트로 쓰지 않는다(오보 방지).
2) 뉴스 클리핑(tab14 재활용)은 아래 3가지 용도로만 쓴다.
   ─ (a) 문체·표현 참고   (b) 전국 동향 문단 자동 삽입   (c) 헤드라인 후보 제안
3) 전국 동향 문단에는 반드시 [근거 뉴스 번호]를 달고, 화면·hwpx 모두에
   "참고한 뉴스 출처표(제목/언론사/날짜/링크)"를 노출하여 사람이 검증할 수 있게 한다.
4) 출력은 (1) 화면 초안 텍스트  (2) 경북도 양식 hwpx 다운로드  둘 다 제공한다.

의존 모듈
----------
- modules/common.py       (get_gemini_model 등 공용 유틸이 있다면 재사용)
- modules/tab14_news.py    뉴스 수집 함수 (없으면 아래 collect_news_fallback 사용)
- 양식 파일: assets/보도자료_양식.hwpx  (첨부한 경북도 보도자료를 템플릿으로 저장)

requirements: streamlit, pandas, openpyxl, google-generativeai
"""

from __future__ import annotations
import io
import os
import re
import json
import zipfile
import datetime as _dt
from typing import Any

import streamlit as st
import pandas as pd

# ------------------------------------------------------------------ #
#  세션 키 프리픽스 (tab14 충돌 교훈 반영 → 모든 위젯키에 pr_ 접두)
# ------------------------------------------------------------------ #
PFX = "pr_"

# 양식 hwpx 경로 (레포에 함께 커밋)
TEMPLATE_HWPX = os.path.join(os.path.dirname(__file__), "assets", "보도자료_양식.hwpx")

GEMINI_MODEL_NAME = "gemini-2.5-flash-lite"   # 시스템 표준
GEMINI_MODEL_PRO  = "gemini-2.5-flash"        # 긴 조합은 살짝 상위 모델 옵션


# ================================================================== #
#  1. 팩트 추출  ―  결과파일(hwpx / xlsx)에서 핵심 수치·항목만 뽑는다
# ================================================================== #
def extract_facts_from_xlsx(file) -> dict[str, Any]:
    """취합 xlsx에서 팩트 후보를 넓게 긁어 딕셔너리로 반환."""
    facts: dict[str, Any] = {"source_type": "xlsx", "raw_tables": []}
    try:
        xls = pd.ExcelFile(file)
        for sh in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sh, header=None, dtype=str)
            df = df.fillna("")
            # 시트를 텍스트 블록으로 만들어 프롬프트에 통째로 넣을 수 있게 저장
            facts["raw_tables"].append({
                "sheet": sh,
                "text": df.to_csv(index=False, header=False),
            })
    except Exception as e:
        facts["error"] = f"xlsx 파싱 실패: {e}"
    return facts


def extract_facts_from_hwpx(file) -> dict[str, Any]:
    """심의결과 등 hwpx에서 미리보기 텍스트(PrvText)와 본문 텍스트를 추출."""
    facts: dict[str, Any] = {"source_type": "hwpx", "text": ""}
    try:
        data = file.read() if hasattr(file, "read") else file
        zf = zipfile.ZipFile(io.BytesIO(data))
        chunks = []
        # 1순위: 미리보기 텍스트(정제됨)
        if "Preview/PrvText.txt" in zf.namelist():
            chunks.append(zf.read("Preview/PrvText.txt").decode("utf-8", "ignore"))
        # 2순위: 본문 section*.xml 의 <hp:t> 런
        for name in zf.namelist():
            if re.match(r"Contents/section\d+\.xml", name):
                xml = zf.read(name).decode("utf-8", "ignore")
                runs = re.findall(r"<hp:t>(.*?)</hp:t>", xml, re.S)
                runs = [re.sub(r"<[^>]+>", "", r) for r in runs]
                chunks.append(" ".join(r for r in runs if r.strip()))
        facts["text"] = "\n".join(chunks)
    except Exception as e:
        facts["error"] = f"hwpx 파싱 실패: {e}"
    return facts


def build_fact_block(uploaded_files) -> str:
    """여러 결과파일을 하나의 '신뢰 팩트 블록' 텍스트로 합친다."""
    blocks = []
    for f in uploaded_files:
        name = f.name
        ext = name.lower().rsplit(".", 1)[-1]
        if ext in ("xlsx", "xlsm", "xls"):
            fx = extract_facts_from_xlsx(f)
            for t in fx.get("raw_tables", []):
                blocks.append(f"[파일:{name} / 시트:{t['sheet']}]\n{t['text']}")
            if "error" in fx:
                blocks.append(f"[파일:{name}] (오류) {fx['error']}")
        elif ext == "hwpx":
            fx = extract_facts_from_hwpx(f)
            blocks.append(f"[파일:{name}]\n{fx.get('text','')}")
        else:
            try:
                blocks.append(f"[파일:{name}]\n{f.read().decode('utf-8','ignore')}")
            except Exception:
                blocks.append(f"[파일:{name}] (읽기 실패)")
    return "\n\n".join(blocks)


# ================================================================== #
#  2. 뉴스 클리핑  ―  tab14 재활용 (없으면 fallback)
# ================================================================== #
def collect_news(keywords: list[str], days: int = 30, per_kw: int = 20,
                 do_fuzzy: bool = True, fuzzy_thr: float = 0.65) -> list[dict]:
    """
    tab14_newsclip 의 완성된 수집·정제 함수를 그대로 재사용한다.
    - search_news(keyword, cid, csec, days, max_items, exact) → (rows, err)
    - dedupe / dedupe_fuzzy 로 중복·유사 보도자료 정리
    tab14의 row 키(날짜/매체/제목/요약/링크)를 tab16 내부 형식
    (title/press/date/link/desc)으로 변환해 반환한다.
    """
    try:
        from modules import tab14_newsclip as _t14
    except Exception as e:
        st.warning(f"tab14_newsclip 모듈을 불러오지 못했습니다: {e}")
        return []

    try:
        cid = st.secrets["NAVER_CLIENT_ID"]
        csec = st.secrets["NAVER_CLIENT_SECRET"]
    except Exception:
        st.warning("NAVER API 키(secrets)가 없어 뉴스 수집을 건너뜁니다.")
        return []

    raw_rows = []
    for kw in keywords:
        try:
            rows, err = _t14.search_news(
                kw, cid, csec, days, max_items=int(per_kw), exact=True,
            )
        except Exception:
            continue
        if err:
            continue
        for r in rows:
            r["키워드"] = kw
        raw_rows.extend(rows)

    if not raw_rows:
        return []

    # tab14의 정제 로직 재사용 (완전일치 + 유사 보도자료 묶기)
    try:
        raw_rows = _t14.dedupe(raw_rows)
        if do_fuzzy and hasattr(_t14, "dedupe_fuzzy"):
            raw_rows = _t14.dedupe_fuzzy(raw_rows, fuzzy_thr)
    except Exception:
        pass

    # tab16 내부 형식으로 변환
    out = []
    for r in raw_rows:
        out.append({
            "title": r.get("제목", ""),
            "press": r.get("매체", ""),
            "date": r.get("날짜", ""),
            "link": r.get("링크", ""),
            "desc": r.get("요약", ""),
            "keyword": r.get("키워드", ""),
        })
    return out


# ================================================================== #
#  3. Gemini 조합  ―  팩트+뉴스 → 보도자료 JSON
# ================================================================== #
PROMPT_TEMPLATE = """당신은 대한민국 광역지방자치단체(경상북도) 공보 담당 주무관입니다.
아래 규칙을 반드시 지켜 '보도자료'를 작성하세요.

[절대 규칙]
1. 모든 구체 수치(지구 수, 필지 수, 면적, 예산, 시·군명, 날짜, 근거 법령)는
   반드시 <신뢰 팩트> 안에서만 가져옵니다. 뉴스에 나온 숫자는 본문 사실로 쓰지 마세요.
2. <참고 뉴스>는 (a) 문장 표현·어투 참고, (b) '전국 동향' 문단 작성 근거로만 씁니다.
3. '전국 동향' 문단을 쓸 때는 문장 끝에 근거 뉴스 번호를 [n] 형태로 표시하세요.
   근거가 없는 전국 동향 문장은 쓰지 마세요.
4. 경상북도 보도자료 문체를 따르세요: 개조식이 아닌 서술형, 문어체,
   "~밝혔다 / ~라고 말했다" 인용, 담당 과장 코멘트로 마무리.
5. 과장·추측 표현 금지. 확인 안 된 내용은 넣지 마세요.

<신뢰 팩트>
{facts}
</신뢰 팩트>

<참고 뉴스>
{news}
</참고 뉴스>

[출력 형식] — 아래 JSON만 출력. 코드펜스·설명 금지.
{{
  "headline_candidates": ["헤드라인 후보 3~5개"],
  "title": "본문에 쓸 제목",
  "subtitle": "부제(- ~ - 형식)",
  "lead": "리드 문단(핵심 팩트 압축, 1~2문장)",
  "body": ["본문 문단 배열(사업개요·현황 등, 각 문단 문자열)"],
  "national_trend": "전국 동향 문단(문장마다 [n] 근거표시 포함)",
  "trend_sources": [근거로 실제 사용한 뉴스 번호들의 정수 배열],
  "quote": "담당 과장 코멘트 문단",
  "used_facts_note": "본문에서 사용한 핵심 수치 목록(검증용)"
}}
"""


def generate_press_release(fact_block: str, news: list[dict], model_name: str) -> dict:
    import google.generativeai as genai
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    model = genai.GenerativeModel(model_name)

    news_lines = []
    for i, n in enumerate(news, 1):
        news_lines.append(
            f"[{i}] ({n.get('press','')}·{n.get('date','')}) {n.get('title','')} "
            f":: {n.get('desc','')[:120]}"
        )
    prompt = PROMPT_TEMPLATE.format(
        facts=fact_block[:16000],
        news="\n".join(news_lines) if news_lines else "(수집된 뉴스 없음)",
    )
    resp = model.generate_content(
        prompt,
        generation_config={"temperature": 0.4, "max_output_tokens": 4096},
    )
    text = (resp.text or "").strip()
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.M).strip()
    try:
        return json.loads(text)
    except Exception:
        # 방어: 첫 { ~ 마지막 } 구간만 재시도
        m = re.search(r"\{.*\}", text, re.S)
        if m:
            return json.loads(m.group(0))
        raise ValueError("Gemini JSON 파싱 실패:\n" + text[:500])


# ================================================================== #
#  4. hwpx 출력  ―  양식 템플릿의 본문 문단만 교체(raw byte 안전 방식)
# ================================================================== #
def render_plain_text(doc: dict, meta: dict) -> str:
    """화면·txt·hwpx 공통으로 쓰는 완성 보도자료 텍스트."""
    lines = []
    lines.append(f"【{meta.get('date','')}】  {meta.get('dept','건설도시국 토지정보과')}")
    lines.append(f"과장 {meta.get('gwajang','')}  담당 {meta.get('damdang','')}  "
                 f"연락처 {meta.get('tel','')}")
    lines.append("")
    lines.append(doc.get("title", ""))
    if doc.get("subtitle"):
        lines.append(doc["subtitle"])
    lines.append("")
    if doc.get("lead"):
        lines.append(doc["lead"])
        lines.append("")
    for p in doc.get("body", []):
        lines.append(p); lines.append("")
    if doc.get("national_trend"):
        lines.append(doc["national_trend"]); lines.append("")
    if doc.get("quote"):
        lines.append(doc["quote"]); lines.append("")
    return "\n".join(lines).strip()


def build_hwpx(template_path: str, doc: dict, meta: dict, sources: list[dict]) -> bytes:
    """
    양식 hwpx를 열어 section0.xml의 '본문 문단들'을 새 내용으로 교체.
    - 담당자 표/제목 등 앵커성 짧은 런은 문자열 치환
    - 리드·본문·코멘트는 문단 하나를 통째로 새 <hp:p>로 대체
    * mimetype은 ZIP_STORED로 첫 엔트리, 디렉토리 엔트리 제외 (mc.park 확립 원칙)
    """
    with open(template_path, "rb") as f:
        raw = f.read()
    zin = zipfile.ZipFile(io.BytesIO(raw))

    # section0.xml 읽기
    sec_name = next(n for n in zin.namelist()
                    if re.match(r"Contents/section\d+\.xml", n))
    sec = zin.read(sec_name).decode("utf-8")

    # --- (a) 담당자 표/헤더 앵커 치환 ---
    replacements = {
        "차 은 미": meta.get("gwajang", "차 은 미"),
        "010-3383-9093": meta.get("cell", "010-3383-9093"),
        "054-880-4055": meta.get("tel", "054-880-4055"),
        "【8. 6.(목)】": f"【{meta.get('date','')}】",
    }
    for old, new in replacements.items():
        if old in sec and new:
            sec = sec.replace(_xesc(old), _xesc(new))

    # --- (b) 제목/부제/본문 문단 교체 ---
    #  원본 대표 문단 텍스트를 앵커로 새 문장으로 치환한다.
    #  (양식이 바뀌면 아래 앵커만 갱신하면 됨)
    body_map = {
        "경북도, 제2차 지적재조사 사업지구 지정 ": doc.get("title", ""),
        " - “일석이조”의 기회를 놓치지 마세요 -": doc.get("subtitle", ""),
    }
    for old, new in body_map.items():
        if old in sec and new:
            sec = sec.replace(_xesc(old), _xesc(new))

    # 본문 문단 전체 재구성: 원본 리드~코멘트 구간을 새 문단들로 교체
    new_body = []
    if doc.get("lead"):
        new_body.append(doc["lead"])
    new_body += doc.get("body", [])
    if doc.get("national_trend"):
        new_body.append(doc["national_trend"])
    if doc.get("quote"):
        new_body.append(doc["quote"])

    sec = _replace_body_paragraphs(sec, new_body)

    # --- (c) 참고 뉴스 출처표를 문서 끝에 부록으로 추가 ---
    if sources:
        sec = _append_sources_appendix(sec, sources)

    # --- 재패키징 (안전 원칙 준수) ---
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zout:
        # mimetype 먼저, 무압축
        zout.writestr(
            zipfile.ZipInfo("mimetype"),
            zin.read("mimetype"),
            compress_type=zipfile.ZIP_STORED,
        )
        for item in zin.infolist():
            if item.filename == "mimetype":
                continue
            if item.filename.endswith("/"):      # 디렉토리 엔트리 제외
                continue
            data = sec.encode("utf-8") if item.filename == sec_name else zin.read(item.filename)
            zout.writestr(item.filename, data, compress_type=zipfile.ZIP_DEFLATED)
    return out.getvalue()


def _xesc(s: str) -> str:
    # hp:t 텍스트 노드 내부에서는 &,<,> 만 이스케이프. (따옴표는 그대로 두어야
    # 한글에서 큰따옴표가 &quot; 로 깨져 보이지 않는다.)
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _replace_body_paragraphs(sec: str, paragraphs: list[str]) -> str:
    """
    본문 리드~코멘트 사이의 <hp:p> 블록들을 새 문단으로 치환.
    구현 전략: 리드 원문 앵커가 있는 문단부터 코멘트 앵커 문단까지를
    하나의 대체 구간으로 잡고, 그 안 문단들을 새 문단 리스트로 대체한다.
    안전을 위해 '문단 껍데기(paraPr/charPr)'는 원본 첫 본문 문단 것을 재사용한다.
    """
    p_iter = list(re.finditer(r"<hp:p\b[^>]*>.*?</hp:p>", sec, re.S))
    if not p_iter:
        return sec

    LEAD_ANCHOR = "경북도는 올해 지적재조사"
    QUOTE_ANCHOR = "과장은"

    start_idx = end_idx = None
    for i, m in enumerate(p_iter):
        block = m.group(0)
        if start_idx is None and LEAD_ANCHOR in block:
            start_idx = i
        if QUOTE_ANCHOR in block:
            end_idx = i
    if start_idx is None or end_idx is None or end_idx < start_idx:
        return sec  # 앵커 실패 시 원본 유지(파괴 방지)

    # 본문 문단 껍데기 추출(스타일 유지용)
    shell = p_iter[start_idx].group(0)
    shell_open = re.match(r"<hp:p\b[^>]*>", shell).group(0)

    # 원본 run 하나의 charPr을 재활용
    run_m = re.search(r"(<hp:run\b[^>]*>).*?(</hp:run>)", shell, re.S)
    run_open = run_m.group(1) if run_m else "<hp:run>"
    def make_para(text: str) -> str:
        return f'{shell_open}{run_open}<hp:t>{_xesc(text)}</hp:t></hp:run></hp:p>'

    new_blocks = "".join(make_para(t) for t in paragraphs if t.strip())

    span_start = p_iter[start_idx].start()
    span_end = p_iter[end_idx].end()
    return sec[:span_start] + new_blocks + sec[span_end:]


def _append_sources_appendix(sec: str, sources: list[dict]) -> str:
    """문서 끝(</hp:sec> 직전 혹은 마지막 문단 뒤)에 참고 뉴스 출처표 문단 추가."""
    lines = ["※ 참고한 전국 보도자료·뉴스 출처(검증용)"]
    for i, s in enumerate(sources, 1):
        lines.append(f"[{i}] {s.get('press','')} · {s.get('date','')} · "
                     f"{s.get('title','')} — {s.get('link','')}")
    # 마지막 </hp:p> 뒤에 삽입할 새 문단들 (단순 텍스트 문단)
    appendix = "".join(
        f'<hp:p><hp:run><hp:t>{_xesc(ln)}</hp:t></hp:run></hp:p>' for ln in lines
    )
    # 문서의 '마지막' </hp:p> 뒤에 삽입 (헤더 표의 subList가 아니라 본문 끝)
    idx = sec.rfind("</hp:p>")
    if idx != -1:
        cut = idx + len("</hp:p>")
        return sec[:cut] + appendix + sec[cut:]
    return sec + appendix


# ================================================================== #
#  5. Streamlit UI
# ================================================================== #
def render_tab16():
    st.markdown("### 📰 보도자료 AI 생성 (전국 참고 조합)")
    st.caption("결과파일의 **팩트**와 tab14 **뉴스클리핑**을 조합해 보도자료 초안·hwpx를 만듭니다. "
               "수치는 업로드 파일에서만, 뉴스는 문체·동향·헤드라인 참고용으로만 씁니다.")

    col_l, col_r = st.columns([1, 1])

    with col_l:
        st.markdown("#### ① 결과파일 업로드 (팩트 소스)")
        files = st.file_uploader(
            "심의결과 hwpx · 취합 xlsx (여러 개 가능)",
            type=["hwpx", "xlsx", "xlsm", "xls", "txt"],
            accept_multiple_files=True, key=PFX + "files",
        )

        st.markdown("#### ② 담당자 정보")
        meta = {}
        c1, c2 = st.columns(2)
        meta["date"] = c1.text_input("배포일", value="", key=PFX + "date",
                                     placeholder="예: 8. 6.(목)")
        meta["dept"] = c2.text_input("담당부서", value="건설도시국 토지정보과", key=PFX + "dept")
        c3, c4 = st.columns(2)
        meta["gwajang"] = c3.text_input("과장", value="차 은 미", key=PFX + "gwajang")
        meta["damdang"] = c4.text_input("담당", value="", key=PFX + "damdang")
        c5, c6 = st.columns(2)
        meta["tel"] = c5.text_input("사무실 전화", value="054-880-4055", key=PFX + "tel")
        meta["cell"] = c6.text_input("휴대전화", value="", key=PFX + "cell")

    with col_r:
        st.markdown("#### ③ 뉴스 참고 설정")
        kw_text = st.text_area(
            "검색 키워드 (한 줄에 하나씩)",
            value="지적재조사\n지적재조사 지구지정\n지적재조사 사업",
            height=110, key=PFX + "kw",
        )
        cc1, cc2 = st.columns(2)
        days = cc1.selectbox("수집 기간", [7, 14, 30, 60, 90],
                             index=2, format_func=lambda d: f"최근 {d}일", key=PFX + "days")
        per_kw = cc2.number_input("키워드당 최대", 50, 1000, 200, step=50, key=PFX + "perkw")
        use_trend = st.checkbox("전국 동향 문단 자동 삽입", value=True, key=PFX + "trend")
        use_head = st.checkbox("헤드라인 후보 3~5개 제안", value=True, key=PFX + "head")
        do_fuzzy = st.checkbox("유사 보도자료 묶기", value=True, key=PFX + "fuzzy",
                               help="같은 보도자료를 매체마다 제목만 바꿔 낸 경우 하나만 남깁니다.")
        model_choice = st.radio("모델", [GEMINI_MODEL_NAME, GEMINI_MODEL_PRO],
                                horizontal=True, key=PFX + "model")

    st.divider()
    go = st.button("📰 보도자료 생성", type="primary", key=PFX + "go")

    if go:
        if not files:
            st.error("결과파일을 먼저 업로드하세요. (수치는 이 파일에서만 추출됩니다)")
            st.stop()

        with st.spinner("① 결과파일에서 팩트 추출 중…"):
            fact_block = build_fact_block(files)
        if not fact_block.strip():
            st.error("업로드 파일에서 텍스트를 추출하지 못했습니다.")
            st.stop()

        keywords = [k.strip() for k in kw_text.splitlines() if k.strip()]
        news = []
        if use_trend or use_head:
            with st.spinner("② 전국 뉴스 클리핑 수집 중…"):
                news = collect_news(keywords, days=int(days), per_kw=int(per_kw),
                                    do_fuzzy=bool(do_fuzzy))
            st.success(f"뉴스 {len(news)}건 수집")

        with st.spinner("③ Gemini가 보도자료 조합 중…"):
            try:
                doc = generate_press_release(fact_block, news, model_choice)
            except Exception as e:
                st.error(f"생성 실패: {e}")
                st.stop()

        # --- 근거 뉴스만 추려 출처표 구성 ---
        trend_idx = doc.get("trend_sources", []) if use_trend else []
        used_sources = [news[i - 1] for i in trend_idx if 1 <= i <= len(news)]

        st.session_state[PFX + "result"] = {
            "doc": doc, "meta": meta,
            "sources": used_sources, "all_news": news,
        }

    # ------------------------------------------------------------ #
    #  결과 표시
    # ------------------------------------------------------------ #
    res = st.session_state.get(PFX + "result")
    if res:
        doc, meta = res["doc"], res["meta"]
        used_sources, all_news = res["sources"], res["all_news"]

        st.markdown("## 📄 생성 결과")

        if doc.get("headline_candidates"):
            st.markdown("#### 헤드라인 후보")
            for i, h in enumerate(doc["headline_candidates"], 1):
                st.markdown(f"- **{i}.** {h}")

        st.markdown("#### 보도자료 초안")
        plain = render_plain_text(doc, meta)
        st.text_area("본문", plain, height=380, key=PFX + "plain")

        # 검증용: 사용한 핵심 수치
        if doc.get("used_facts_note"):
            with st.expander("🔎 본문에 사용된 핵심 수치(검증용)"):
                st.write(doc["used_facts_note"])

        # 전국 동향 근거 뉴스 출처표 (검증 필수)
        if used_sources:
            st.markdown("#### 🔗 전국 동향 문단이 참고한 뉴스 (검증)")
            df = pd.DataFrame([{
                "번호": i + 1, "언론사": s.get("press", ""),
                "날짜": s.get("date", ""), "제목": s.get("title", ""),
                "링크": s.get("link", ""),
            } for i, s in enumerate(used_sources)])
            st.dataframe(df, use_container_width=True,
                         column_config={"링크": st.column_config.LinkColumn()})
        elif res["doc"].get("national_trend"):
            st.warning("전국 동향 문단은 있으나 근거 뉴스 번호가 비어 있습니다. "
                       "동향 문단을 삭제하거나 뉴스를 다시 수집하세요.")

        # 참고한 전체 뉴스 목록(펼침)
        if all_news:
            with st.expander(f"📰 수집된 전체 뉴스 {len(all_news)}건 보기"):
                st.dataframe(pd.DataFrame(all_news), use_container_width=True)

        # ------- 다운로드 -------
        st.markdown("#### ⬇️ 다운로드")
        d1, d2 = st.columns(2)
        d1.download_button(
            "📝 텍스트(txt) 저장", plain.encode("utf-8"),
            file_name="보도자료_초안.txt", key=PFX + "dl_txt",
        )
        try:
            if os.path.exists(TEMPLATE_HWPX):
                hwpx_bytes = build_hwpx(TEMPLATE_HWPX, doc, meta, used_sources)
                d2.download_button(
                    "📄 hwpx(경북도 양식) 저장", hwpx_bytes,
                    file_name="보도자료.hwpx", key=PFX + "dl_hwpx",
                    mime="application/octet-stream",
                )
            else:
                d2.info("양식 파일(assets/보도자료_양식.hwpx)을 레포에 넣으면 hwpx도 생성됩니다.")
        except Exception as e:
            d2.error(f"hwpx 생성 오류: {e}")


# main.py 에서:  from modules.tab16_pressrelease import render_tab16
#               with tabs[15]: render_tab16()
if __name__ == "__main__":
    render_tab16()
