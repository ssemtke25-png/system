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
3) 전국 동향 문단에는 반드시 [근거 뉴스 번호]를 달고, 초안 하단에
   "참고한 뉴스 출처표(제목/언론사/날짜/링크)"를 붙여 사람이 검증할 수 있게 한다.
4) 출력은 화면 초안 텍스트 + txt 저장. (붙여넣기로 한글 양식에 옮겨 씀)

의존 모듈
----------
- modules/tab14_newsclip.py   뉴스 수집·정제 함수 재사용

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


def extract_facts_from_pdf(file) -> dict[str, Any]:
    """PDF에서 텍스트·표를 추출한다.

    - pdfplumber가 있으면 표까지 추출(통계 PDF에 유리)
    - 없으면 pypdf로 텍스트만 추출
    - 텍스트가 거의 안 나오면 '스캔 PDF(이미지)'로 판단해 안내한다.
    """
    facts: dict[str, Any] = {"source_type": "pdf", "text": "", "scanned": False}
    data = file.read() if hasattr(file, "read") else file
    chunks = []

    # 1순위: pdfplumber (텍스트 + 표)
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                t = page.extract_text() or ""
                if t.strip():
                    chunks.append(t)
                for tbl in (page.extract_tables() or []):
                    for row in tbl:
                        cells = [c for c in row if c]
                        if cells:
                            chunks.append(" | ".join(str(c) for c in cells))
        facts["text"] = "\n".join(chunks)
    except Exception:
        # 2순위: pypdf (텍스트만)
        try:
            import pypdf
            reader = pypdf.PdfReader(io.BytesIO(data))
            for page in reader.pages:
                t = page.extract_text() or ""
                if t.strip():
                    chunks.append(t)
            facts["text"] = "\n".join(chunks)
        except Exception as e:
            facts["error"] = f"pdf 파싱 실패: {e}"
            return facts

    # 스캔 PDF 판정: 추출 텍스트가 극히 적으면 이미지 PDF로 간주
    if len(facts["text"].strip()) < 20:
        facts["scanned"] = True

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
        elif ext == "pdf":
            fx = extract_facts_from_pdf(f)
            if fx.get("scanned"):
                st.warning(f"⚠️ '{name}' 은(는) 스캔(이미지) PDF로 보여 텍스트를 "
                           f"읽지 못했습니다. 텍스트 PDF·xlsx·hwpx로 넣거나, "
                           f"핵심 수치를 직접 입력해 주세요.")
                blocks.append(f"[파일:{name}] (스캔 PDF — 텍스트 추출 불가)")
            elif "error" in fx:
                blocks.append(f"[파일:{name}] (오류) {fx['error']}")
            else:
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

    # 관련성 필터: 검색 키워드가 '제목'에 있는 기사만 남긴다.
    # (본문에 스쳐 지나간 무관 기사 — 연예·스포츠·사건 등 — 를 사전 차단)
    # tab14의 apply_title_match를 재사용하되, 없으면 자체 로직으로 처리.
    try:
        if hasattr(_t14, "apply_title_match"):
            raw_rows, _ = _t14.apply_title_match(raw_rows)
        else:
            def _sq(s):
                return re.sub(r"[\s\u3000·・…‥\-–—_/\\|]", "", s or "").lower()
            raw_rows = [r for r in raw_rows
                        if _sq(r.get("키워드", "")) in _sq(r.get("제목", ""))]
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

[주체 판단 규칙 — 가장 먼저 적용]
- <신뢰 팩트>가 경상북도(또는 경북 산하 시·군)의 사업·실적·계획에 관한 것이면,
  경상북도를 주어로 하는 일반적인 도(道) 보도자료로 작성합니다.
- <신뢰 팩트>가 국토교통부·타 시도·전국 단위 등 경북이 주체가 아닌 내용이면,
  억지로 "경상북도가 ~한다"로 바꾸지 말고, 자료에 나온 실제 주체를 그대로
  주어로 쓰거나 중립적인 관점으로 작성합니다.
  · 예: 국토부 전국 경진대회 자료 → "국토교통부가 ~를 개최한다"가 주어.
    경북 참가 사실이 자료에 '명시'되어 있을 때만 경북 참가를 언급합니다.
- 자료에 없는 경북의 참가·수상·기대·포부는 절대 지어내지 마세요.

[절대 규칙]
1. 모든 구체 수치(지구 수, 필지 수, 면적, 예산, 시·군명, 날짜, 근거 법령)와
   '누가 무엇을 했는지'는 반드시 <신뢰 팩트> 안에서만 가져옵니다.
   뉴스에 나온 숫자는 본문 사실로 쓰지 마세요.
2. <참고 뉴스>는 (a) 문장 표현·어투 참고, (b) '전국 동향' 문단 작성 근거로만 씁니다.
3. '전국 동향' 문단 규칙(엄격히 지킬 것):
   - <신뢰 팩트>의 주제와 '직접 관련된' 뉴스만 근거로 씁니다.
     주제와 무관한 뉴스(연예·스포츠·사건사고 등)는 절대 참고하지 마세요.
   - 관련 뉴스가 2건 미만이면 전국 동향 문단을 아예 쓰지 마세요(national_trend를 빈 문자열 "").
   - 각 문장 끝에는 근거 뉴스 번호를 [n] 형태로 표시하되, 한 문장에 최대 2개까지만.
   - 전체 동향 문단은 2~3문장을 넘기지 마세요. 번호를 여러 개 나열하지 마세요.
   - trend_sources에는 실제로 문장 근거로 쓴 번호만(최대 4개) 넣습니다.
4. 보도자료 문체를 따르세요: 개조식이 아닌 서술형, 문어체,
   "~밝혔다 / ~라고 말했다" 인용, 담당자 코멘트로 마무리.
   단, 코멘트의 인물은 자료에 그 발언·직위가 있을 때만 넣고, 없으면 코멘트를 비웁니다.
5. 과장·추측 표현 금지. 확인 안 된 내용은 넣지 마세요.
   자료로 뒷받침되지 않는 문장은 아예 쓰지 않는 것이 원칙입니다.


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
        doc = json.loads(text)
    except Exception:
        # 방어: 첫 { ~ 마지막 } 구간만 재시도
        m = re.search(r"\{.*\}", text, re.S)
        if not m:
            raise ValueError("Gemini JSON 파싱 실패:\n" + text[:500])
        doc = json.loads(m.group(0))

    return _sanitize_trend(doc)


def _sanitize_trend(doc: dict) -> dict:
    """전국 동향 문단의 '번호 폭탄'을 후처리로 차단한다.

    - 한 문장(또는 한 곳)에 근거 번호가 3개 이상 몰려 있으면 → 무관 뉴스를
      긁어 붙인 신호로 보고, 동향 문단 전체를 폐기한다.
    - [n] 표기가 비정상적으로 많으면(4개 초과) 역시 폐기한다.
    - trend_sources도 최대 4개로 제한한다.
    """
    trend = doc.get("national_trend", "") or ""

    # [1, 2, 3, ... 267] 처럼 대괄호 안에 번호가 여러 개 나열된 경우 탐지
    bracket_lists = re.findall(r"\[([\d,\s]+)\]", trend)
    max_in_bracket = 0
    for b in bracket_lists:
        nums = [x for x in re.split(r"[,\s]+", b) if x.strip().isdigit()]
        max_in_bracket = max(max_in_bracket, len(nums))

    # 전체 [n] 개수
    total_refs = len(re.findall(r"\[\d+\]", trend)) + sum(
        len([x for x in re.split(r"[,\s]+", b) if x.strip().isdigit()])
        for b in bracket_lists
    )

    # 한 곳에 3개 이상 몰렸거나, 전체 참조가 4개 초과면 → 억지 동향으로 판단, 폐기
    if max_in_bracket >= 3 or total_refs > 4:
        doc["national_trend"] = ""
        doc["trend_sources"] = []
        return doc

    # trend_sources 최대 4개로 제한
    ts = doc.get("trend_sources", []) or []
    if isinstance(ts, list):
        doc["trend_sources"] = ts[:4]

    return doc


# ================================================================== #
#  4. 텍스트 초안 생성
# ================================================================== #
def render_plain_text(doc: dict, meta: dict, sources: list[dict] | None = None) -> str:
    """화면 표시·txt 저장·복사용 완성 보도자료 텍스트."""
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

    # 참고 뉴스 출처표(검증용) — 붙여넣기 후 이 부분은 삭제하고 쓰면 됨
    if sources:
        lines.append("─" * 30)
        lines.append("※ 전국 동향 문단이 참고한 뉴스 출처(검증용, 배포 전 삭제)")
        for i, s in enumerate(sources, 1):
            lines.append(f"[{i}] {s.get('press','')} · {s.get('date','')} · "
                         f"{s.get('title','')} — {s.get('link','')}")

    return "\n".join(lines).strip()


# ================================================================== #
#  5. Streamlit UI
# ================================================================== #
def render_tab16():
    st.markdown("### 📰 보도자료 AI 생성 (전국 참고 조합)")
    st.caption("결과파일의 **팩트**와 tab14 **뉴스클리핑**을 조합해 보도자료 초안을 만듭니다. "
               "수치는 업로드 파일에서만, 뉴스는 문체·동향·헤드라인 참고용으로만 씁니다. "
               "자료가 경북 사업이면 경북 주어로, 국토부·전국 자료면 중립적으로 작성됩니다.")

    col_l, col_r = st.columns([1, 1])

    with col_l:
        st.markdown("#### ① 결과파일 업로드 (팩트 소스)")
        files = st.file_uploader(
            "심의결과 hwpx · 취합 xlsx · PDF (여러 개 가능)",
            type=["hwpx", "xlsx", "xlsm", "xls", "pdf", "txt"],
            accept_multiple_files=True, key=PFX + "files",
        )

        # 업로드 파일이 바뀌면(추가·삭제·교체) 이전 생성 결과를 폐기한다.
        # 파일 지문(이름+크기)을 세션에 저장해두고 달라지면 result를 비운다.
        # → 파일만 바꾸고 재생성 안 했을 때 옛 보도자료가 남는 캐시 문제 방지.
        cur_sig = tuple(sorted((f.name, f.size) for f in files)) if files else ()
        if st.session_state.get(PFX + "filesig") != cur_sig:
            st.session_state[PFX + "filesig"] = cur_sig
            st.session_state.pop(PFX + "result", None)

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
    #  결과 표시 (업로드 파일이 남아 있을 때만)
    # ------------------------------------------------------------ #
    res = st.session_state.get(PFX + "result")
    if res and files:
        doc, meta = res["doc"], res["meta"]
        used_sources, all_news = res["sources"], res["all_news"]

        st.markdown("## 📄 생성 결과")

        if doc.get("headline_candidates"):
            st.markdown("#### 헤드라인 후보")
            for i, h in enumerate(doc["headline_candidates"], 1):
                st.markdown(f"- **{i}.** {h}")

        st.markdown("#### 보도자료 초안")
        plain = render_plain_text(doc, meta, used_sources)
        st.text_area("본문", plain, height=380)

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
        st.markdown("#### ⬇️ 저장")
        st.caption("위 초안을 복사해 한글 보도자료 양식에 붙여넣으면 됩니다. "
                   "txt 파일로도 받을 수 있어요.")
        st.download_button(
            "📝 텍스트(txt) 저장", plain.encode("utf-8"),
            file_name="보도자료_초안.txt", key=PFX + "dl_txt",
        )


# main.py 에서:  from modules.tab16_pressrelease import render_tab16
#               with tabs[15]: render_tab16()
if __name__ == "__main__":
    render_tab16()
