"""
tab7_hwpx_ai.py
행사 계획서 기반 AI 문서 자동생성 탭
- HWPX 업로드 → 요약 → 7가지 생성 기능
"""

import streamlit as st
import google.generativeai as genai
import zipfile
import io
import re
import json
import requests
from pathlib import Path
from datetime import datetime

# ─────────────────────────────────────────────
# 초기 설정
# ─────────────────────────────────────────────
GEMINI_MODEL = "gemini-3.5-flash"
MALGUN_GOTHIC = "맑은 고딕"

render = None  # main.py 호환용 (아래에서 render_tab7로 덮어씀)


def get_gemini_model():
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    return genai.GenerativeModel(GEMINI_MODEL)


def gemini_text(prompt: str) -> str:
    """Gemini 호출 → 텍스트 반환"""
    return get_gemini_model().generate_content(prompt).text.strip()


# ─────────────────────────────────────────────
# HWPX 텍스트 추출
# ─────────────────────────────────────────────
def extract_hwpx_text(hwpx_file) -> str:
    text_parts = []
    try:
        with zipfile.ZipFile(io.BytesIO(hwpx_file.read()), "r") as zf:
            xml_files = [n for n in zf.namelist() if n.endswith(".xml")]
            content_xmls = sorted([f for f in xml_files if "Contents" in f or "content" in f.lower()])
            target = content_xmls if content_xmls else xml_files
            for name in target:
                try:
                    data = zf.read(name).decode("utf-8", errors="ignore")
                    clean = re.sub(r"<[^>]+>", " ", data)
                    clean = re.sub(r"\s+", " ", clean).strip()
                    if len(clean) > 50:
                        text_parts.append(clean)
                except Exception:
                    continue
    except Exception as e:
        return f"[파일 읽기 오류: {e}]"
    full = " ".join(text_parts)
    return full[:8000] if len(full) > 8000 else full


def summarize_plan(raw_text: str) -> str:
    prompt = f"""너는 20년 차 베테랑 공무원이다.
아래 행사 계획서 원문을 분석하여 핵심 정보를 JSON 형식으로 정리하라.
반드시 JSON만 출력하고 마크다운 코드블록 없이 순수 JSON만 반환하라.
값이 없는 항목은 빈 문자열("")로 표기하라.

원문:
{raw_text}

출력 JSON 형식:
{{
  "행사명": "",
  "주최기관": "",
  "주관기관": "",
  "일시": "",
  "장소": "",
  "대상": "",
  "참석예정인원": "",
  "행사목적": "",
  "주요프로그램": [""],
  "예산": "",
  "담당부서": "",
  "담당자": "",
  "기타특이사항": ""
}}"""
    return gemini_text(prompt)


# ─────────────────────────────────────────────
# DOCX 생성 유틸
# ─────────────────────────────────────────────
def _paragraphs_to_docx(paragraphs: list, out_name: str) -> bytes | None:
    try:
        from docx import Document
        from docx.shared import Pt, RGBColor, Cm

        doc = Document()
        doc.core_properties.author = "AI 문서 자동생성 시스템"
        doc.core_properties.title = out_name.replace(".docx", "")

        style = doc.styles["Normal"]
        style.font.name = MALGUN_GOTHIC
        style.font.size = Pt(11)

        for h_style, sz, color in [
            ("Heading 1", 18, (31, 56, 100)),
            ("Heading 2", 14, (46, 116, 181)),
        ]:
            s = doc.styles[h_style]
            s.font.name = MALGUN_GOTHIC
            s.font.size = Pt(sz)
            s.font.color.rgb = RGBColor(*color)

        for section in doc.sections:
            section.top_margin = Cm(2.5)
            section.bottom_margin = Cm(2.5)
            section.left_margin = Cm(3.0)
            section.right_margin = Cm(2.5)

        for p in paragraphs:
            text = p.get("text", "")
            h = p.get("heading")
            bold = p.get("bold", False)
            if h == 1:
                doc.add_heading(text, level=1)
            elif h == 2:
                doc.add_heading(text, level=2)
            else:
                para = doc.add_paragraph()
                run = para.add_run(text)
                run.font.name = MALGUN_GOTHIC
                run.font.size = Pt(11)
                if bold:
                    run.bold = True

        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()
    except ImportError:
        st.error("python-docx 패키지가 필요합니다: pip install python-docx")
        return None


def text_to_display(content: str, sess_key: str, label: str, height: int = 400):
    """텍스트를 세션에 저장하고 text_area로 표시"""
    st.session_state[sess_key] = content
    _show_text_area(sess_key, label, height)


def _show_text_area(sess_key: str, label: str, height: int = 400):
    if st.session_state.get(sess_key):
        st.caption("📋 텍스트 박스 클릭 → Ctrl+A → Ctrl+C → 한글/워드에 붙여넣기")
        st.text_area(
            label=label,
            value=st.session_state[sess_key],
            height=height,
            key=f"ta_{sess_key}",
        )


# ─────────────────────────────────────────────
# 1. 문서 4종
# ─────────────────────────────────────────────
DOC_TYPES = {
    "요약보고서": {
        "prompt": """너는 20년 차 베테랑 공무원이다.
아래 행사 계획서 요약을 분석하여 바쁜 상사 보고용 핵심 요약본을 작성하라.

[작성 원칙]
- '감성적인 수식어'는 모두 배제한다
- '신뢰감 있고 건조한 문체(행정 공문서 체계)'로 작성한다
- 개조식(글머리 기호 사용)으로 작성한다
- A4 1장 이내 분량

[반드시 포함할 항목 - 개조식]
○ 행사 개요 (행사명, 일시, 장소, 주최/주관)
○ 추진 목적
○ 주요 내용 (프로그램 구성)
○ 참석 대상 및 규모
○ 소요 예산
○ 기대 효과
○ 향후 계획

마크다운 헤딩(#) 없이 순수 텍스트와 ○ 기호로만 작성하라.""",
    },
    "국장인사말": {
        "prompt": """너는 20년 차 베테랑 공무원이다.
아래 행사 계획서 요약을 바탕으로 국장님용 인사말씀을 작성하라.

[작성 원칙]
- 20년 차 공무원의 시각에서 작성한다
- 행사와 우리 기관 정책 방향의 연계성을 강조한다
- 내빈에게는 정중하되 건조하고 신뢰감 있는 문체로 작성한다
- 감성적인 문구는 제거하고 행정적인 책임감과 비전 위주로 작성한다
- 서두에는 자연스러운 인삿말을 넣되, 따뜻한 어투도 허용한다
- A4 1장 이내 (400~600자)

[구성]
1. 서두 인삿말 (참석자 환영 및 감사)
2. 행사의 행정적 의미와 정책 연계성
3. 기관의 추진 방향 및 비전 제시
4. 마무리 당부 말씀

마크다운 없이 순수 텍스트로 작성하라.""",
    },
    "과장인사말": {
        "prompt": """너는 20년 차 베테랑 공무원이다.
아래 행사 계획서 요약을 바탕으로 과장님용 인사말씀을 작성하라.

[작성 원칙]
- 실무 총괄 과장 버전이다
- 행사 준비 상황에 대한 언급을 포함한다
- 참가자들에 대한 실무적인 당부와 협조 요청을 담는다
- 행사의 원활한 운영을 위한 안내 사항을 포함한다
- 서두에는 자연스러운 인삿말을 넣되, 따뜻한 어투도 허용한다
- A4 1장 이내 (300~500자)

[구성]
1. 서두 인삿말
2. 행사 준비 과정 및 참가자 노고 치하
3. 행사 목적 및 주요 내용 안내
4. 참가자 당부 및 협조 요청
5. 마무리

마크다운 없이 순수 텍스트로 작성하라.""",
    },
    "보도자료": {
        "prompt": """너는 20년 차 베테랑 공무원이다.
아래 행사 계획서 요약을 바탕으로 공식 언론 배포용 보도자료를 작성하라.

[작성 원칙]
- 문체: '~했다', '~한다'로 끝나는 명료한 기사체
- 불필요한 감상적 표현이나 형용사는 모두 배제한다
- 행정 보도자료의 공식 형식을 엄수한다

[반드시 지킬 구조]
1) 상단 정보란:
   작성일자: (행사 계획서의 일시 기준)
   담당부서: (계획서의 담당부서)
   담당자: (계획서의 담당자)
   연락처: 000-0000-0000

2) 제목: 본문의 핵심 성과를 포함한 한 문장
   예) "○○부, △△ 행사 개최…□□ 역량 강화 나선다"

3) 부제목: 대시(-)를 활용하여 핵심 포인트 1~2개 요약
   예) - △△ 전문인력 ○○명 참가, □□ 성과 기대

4) 본문:
   - 첫 문단: 육하원칙에 따라 행사를 명확히 기술
   - 이후 문단: 추진 배경 → 행사 목적 → 구체적 기대효과 순으로 서술

5) 기관장 인용구:
   ○○○ ○○○○은 "인용구 내용"이라고 말했다.

6) 문의처:
   ○○부 ○○과 담당자명 (전화번호)

마크다운 없이 순수 텍스트로 작성하라.""",
    },
}


def _render_one_doc(doc_name: str, cfg: dict, summary_str: str):
    sess_key = f"doc4_text_{doc_name}"
    st.markdown(f"**{doc_name}**")

    if st.button(f"✍️ {doc_name} 생성", key=f"doc4_btn_{doc_name}"):
        with st.spinner(f"{doc_name} 생성 중..."):
            full_prompt = cfg["prompt"] + f"\n\n행사 계획서 요약:\n{summary_str}"
            content = gemini_text(full_prompt)
        st.session_state[sess_key] = content

    _show_text_area(sess_key, doc_name, height=420)


def render_doc4():
    st.subheader("📄 문서 4종 자동생성")
    st.caption("생성 후 텍스트 박스에서 Ctrl+A → Ctrl+C 로 복사하여 한글/워드에 붙여넣기 하세요.")
    summary = st.session_state.get("plan_summary_dict", {})
    summary_str = json.dumps(summary, ensure_ascii=False, indent=2)

    doc_list = list(DOC_TYPES.items())
    col1, col2 = st.columns(2)
    with col1:
        _render_one_doc(doc_list[0][0], doc_list[0][1], summary_str)
        st.markdown("---")
        _render_one_doc(doc_list[2][0], doc_list[2][1], summary_str)
    with col2:
        _render_one_doc(doc_list[1][0], doc_list[1][1], summary_str)
        st.markdown("---")
        _render_one_doc(doc_list[3][0], doc_list[3][1], summary_str)


# ─────────────────────────────────────────────
# 2. 사회자 멘트
# ─────────────────────────────────────────────
MC_DEFAULT = [
    "개회선언",
    "국민의례",
    "내빈소개",
    "기관장인사말",
    "축사",
    "주요프로그램 소개",
    "폐회선언",
]


def render_mc():
    st.subheader("🎤 사회자 멘트 자동생성")
    summary = st.session_state.get("plan_summary_dict", {})
    summary_str = json.dumps(summary, ensure_ascii=False, indent=2)

    # 사용자가 순서 편집 가능
    with st.expander("⚙️ 순서 편집 (항목 수정/추가 가능)", expanded=False):
        default_order = "\n".join(MC_DEFAULT)
        edited = st.text_area(
            "순서 (한 줄에 하나씩)",
            value=default_order,
            height=200,
            key="mc_order_edit",
        )
        mc_order = [line.strip() for line in edited.split("\n") if line.strip()]

    st.info("순서: " + " → ".join(mc_order))

    tone = st.selectbox(
        "멘트 톤",
        ["격식체 (공식 행사)", "친근체 (소규모/내부 행사)", "방송체 (대규모/공개 행사)"],
        key="mc_tone",
    )

    if st.button("✍️ 사회자 멘트 생성", key="mc_gen"):
        prompt = f"""너는 20년 차 베테랑 공무원이자 공공기관 행사 전문 사회자다.
아래 행사 정보와 순서에 맞는 사회자 멘트를 작성하라.

[톤]: {tone}
[순서]: {', '.join(mc_order)}

[행사 계획서 요약]:
{summary_str}

[작성 원칙]
- 각 순서는 ## 헤딩으로 구분한다
- 멘트는 실제 현장에서 바로 읽을 수 있는 완성된 문장으로 작성한다
- 각 순서별 100~200자 내외
- 격식체: '~하겠습니다', '~드립니다' 등 공식적 존댓말
- 친근체: 자연스럽고 따뜻하되 품위 유지
- 방송체: 명확한 발음과 강약이 느껴지는 리듬감 있는 문장
- 다음 순서로 자연스럽게 넘어가는 전환 문구를 반드시 포함한다
- 내빈 소개 순서에서는 직책 → 성함 순으로 소개하는 형식을 유지한다"""

        with st.spinner("사회자 멘트 생성 중..."):
            content = gemini_text(prompt)
        st.session_state["mc_content"] = content

    _show_text_area("mc_content", "사회자 멘트", height=550)


# ─────────────────────────────────────────────
# 3. 현수막 문구
# ─────────────────────────────────────────────
def render_banner():
    st.subheader("🪧 현수막 문구 시안")
    summary = st.session_state.get("plan_summary_dict", {})
    summary_str = json.dumps(summary, ensure_ascii=False, indent=2)

    col1, col2 = st.columns(2)
    with col1:
        n = st.slider("시안 수", 3, 7, 5, key="banner_n")
    with col2:
        style = st.selectbox(
            "문구 스타일",
            ["공식/격식형", "친근/따뜻형", "역동/강조형", "혼합 (다양하게)"],
            key="banner_style",
        )

    if st.button("✍️ 현수막 문구 생성", key="banner_gen"):
        prompt = f"""너는 20년 차 베테랑 공무원이자 공공기관 홍보 전문가다.
아래 행사 계획서를 바탕으로 현수막 문구 시안 {n}개를 작성하라.

[스타일]: {style}
[행사 계획서 요약]: {summary_str}

[작성 원칙]
- 각 시안은 15~25자 이내로 간결하게 작성한다
- 행사 목적과 핵심 메시지가 명확히 담겨야 한다
- 감상적이거나 모호한 표현은 배제한다
- 번호. 문구 형식으로 출력한다
- 각 문구 아래 한 줄 사용 설명 추가 (어떤 상황/위치에 적합한지)

[스타일별 기준]
- 공식/격식형: 기관명·사업명 중심, 권위 있고 전문적인 어조
- 친근/따뜻형: 대상자 중심, 공감과 참여를 유도하는 어조
- 역동/강조형: 핵심 키워드 강조, 임팩트 있고 기억에 남는 문구
- 혼합형: 위 3가지 스타일을 골고루 섞어 다양하게 제시"""

        with st.spinner("현수막 문구 생성 중..."):
            content = gemini_text(prompt)
        st.session_state["banner_content"] = content

    _show_text_area("banner_content", "현수막 문구 시안", height=350)


# ─────────────────────────────────────────────
# 4. 결과보고서 초안
# ─────────────────────────────────────────────
def render_result_report():
    st.subheader("📋 결과보고서 초안")
    summary = st.session_state.get("plan_summary_dict", {})
    summary_str = json.dumps(summary, ensure_ascii=False, indent=2)

    col1, col2 = st.columns(2)
    with col1:
        actual_attendance = st.text_input("실제 참석 인원", placeholder="예: 150명", key="result_attendance")
        satisfaction = st.text_input("만족도 조사 결과", placeholder="예: 4.2/5.0 (85% 응답)", key="result_satisfaction")
    with col2:
        result_note = st.text_area(
            "주요 성과 및 특이사항",
            placeholder="예) 전년 대비 참석률 20% 증가\n예산 집행률 98%\n언론 보도 3건",
            height=105,
            key="result_note",
        )

    if st.button("✍️ 결과보고서 초안 생성", key="result_gen"):
        prompt = f"""너는 20년 차 베테랑 공무원이다.
아래 행사 계획서 요약과 결과 정보를 바탕으로 행사 결과보고서 초안을 작성하라.

[행사 결과 정보]
- 실제 참석 인원: {actual_attendance or '미기재'}
- 만족도 조사: {satisfaction or '미기재'}
- 주요 성과 및 특이사항: {result_note or '없음'}

[행사 계획서 요약]
{summary_str}

[작성 원칙]
- '감성적인 수식어'는 모두 배제한다
- 신뢰감 있고 건조한 행정 공문서 문체로 작성한다
- 수치와 사실 위주로 명확하게 기술한다
- 계획 대비 실적을 비교하는 형식을 포함한다

[작성 형식 - 개조식 중심]
1. 행사 개요
   ○ 행사명 / 일시 / 장소 / 주최·주관 / 참석 인원

2. 추진 결과
   ○ 참석 현황 (계획 대비 실적)
   ○ 프로그램별 진행 결과
   ○ 만족도 조사 결과

3. 예산 집행 현황
   ○ 편성 예산 / 집행액 / 집행률

4. 주요 성과 및 시사점
   ○ 성과 요약
   ○ 미흡 사항 및 개선 방향

5. 향후 계획
   ○ 후속 조치 사항

마크다운 없이 순수 텍스트와 ○ 기호로만 작성하라."""

        with st.spinner("결과보고서 초안 생성 중..."):
            content = gemini_text(prompt)
        st.session_state["result_content"] = content

    _show_text_area("result_content", "결과보고서 초안", height=550)


# ─────────────────────────────────────────────
# 5. PPT 자동생성
# ─────────────────────────────────────────────
PPT_PROMPT_BASE = """너는 20년 차 베테랑 공무원이자 공공기관 발표자료 전문가다.
아래 행사 계획서 요약을 바탕으로 현장에서 바로 사용할 수 있는 발표자료 슬라이드를 구성하라.

[작성 원칙]
- 감성적 수식어 배제, 신뢰감 있고 건조한 행정 문체 사용
- 수치와 사실 위주로 간결하게 작성
- 각 bullet은 30자 이내, 핵심 키워드 중심

[슬라이드 구성 원칙]
1. 첫 슬라이드: 반드시 "title" 레이아웃 (행사명 + 일시/장소/주최 부제목)
2. 두 번째 슬라이드: "content" 레이아웃으로 목차 구성 (01. 02. 형식)
3. 중간 섹션 전환 시: "section" 레이아웃 사용
4. 마지막 슬라이드: "closing" 레이아웃
5. "table" 레이아웃: 일정표·예산표·역할분담표 등 표가 필요할 때

[필수 슬라이드 순서]
표지(title) → 목차(content) → 행사 개요(content) → 추진 배경 및 목적(content) → 주요 프로그램(table) → 세부 내용(content, 계획서 내용 반영) → 기대 효과(content) → 클로징(closing)

[출력 형식 - 반드시 순수 JSON 배열만, 마크다운 코드블록 없이]
각 슬라이드: slide_num, layout(title/content/section/table/closing), title, subtitle(표지·클로징만), bullets(content용), headers(table 컬럼명), rows(table 데이터 2차원 배열)"""


def _build_ppt_prompt(slide_count: int, summary_str: str) -> str:
    return (
        PPT_PROMPT_BASE
        + f"\n\n슬라이드 수: {slide_count}개 (표지·클로징 포함)"
        + f"\n\n행사 계획서 요약:\n{summary_str}"
    )


THEMES = {
    "네이비 (공공기관 정장)": {
        "bg":      (0x12, 0x1A, 0x2E),
        "bg2":     (0x1A, 0x27, 0x44),
        "accent":  (0x00, 0xB4, 0xD8),
        "accent2": (0xC8, 0xF5, 0x00),
        "card":    (0x1E, 0x2D, 0x4A),
        "text":    (0xFF, 0xFF, 0xFF),
        "subtext": (0xA0, 0xBE, 0xD8),
    },
    "다크그린 (환경/생태)": {
        "bg":      (0x0A, 0x1A, 0x0F),
        "bg2":     (0x12, 0x2A, 0x18),
        "accent":  (0x00, 0xE5, 0x96),
        "accent2": (0xC8, 0xF5, 0x00),
        "card":    (0x14, 0x2D, 0x1C),
        "text":    (0xFF, 0xFF, 0xFF),
        "subtext": (0x90, 0xC8, 0xA0),
    },
    "다크버건디 (품격)": {
        "bg":      (0x1A, 0x08, 0x0E),
        "bg2":     (0x2A, 0x10, 0x16),
        "accent":  (0xFF, 0x6B, 0x6B),
        "accent2": (0xFF, 0xD9, 0x3D),
        "card":    (0x2E, 0x12, 0x18),
        "text":    (0xFF, 0xFF, 0xFF),
        "subtext": (0xD4, 0x9A, 0xA0),
    },
    "다크차콜 (모던)": {
        "bg":      (0x0F, 0x0F, 0x0F),
        "bg2":     (0x1A, 0x1A, 0x1A),
        "accent":  (0x00, 0xD4, 0xFF),
        "accent2": (0xFF, 0xC8, 0x00),
        "card":    (0x22, 0x22, 0x22),
        "text":    (0xFF, 0xFF, 0xFF),
        "subtext": (0xA0, 0xA0, 0xA0),
    },
}


def render_ppt():
    st.subheader("📊 PPT 자동생성")
    st.caption("행사 계획서 기반으로 현장 사용 가능한 발표자료를 생성합니다.")
    summary = st.session_state.get("plan_summary_dict", {})
    summary_str = json.dumps(summary, ensure_ascii=False, indent=2)

    col1, col2 = st.columns([1, 2])
    with col1:
        slide_count = st.slider("슬라이드 수", 8, 20, 12, key="ppt_slide_count")
    with col2:
        theme = st.selectbox(
            "색상 테마",
            list(THEMES.keys()),
            key="ppt_theme",
        )

    if st.button("🖥️ PPT 생성", key="ppt_gen", type="primary"):
        prompt = _build_ppt_prompt(slide_count, summary_str)

        with st.spinner("AI가 슬라이드 내용을 구성 중... (10~20초 소요)"):
            raw = gemini_text(prompt)

        try:
            clean = re.sub(r"```json|```", "", raw).strip()
            slides_data = json.loads(clean)
        except Exception as e:
            st.error(f"슬라이드 데이터 파싱 오류: {e}")
            with st.expander("AI 원본 응답 확인"):
                st.text(raw[:1000])
            return

        with st.spinner("PPT 파일 생성 중..."):
            pptx_bytes = _build_pptx(slides_data, summary, theme)

        if pptx_bytes:
            event_name = summary.get("행사명", "행사") if summary else "행사"
            st.session_state["ppt_bytes"] = pptx_bytes
            st.session_state["ppt_name"] = f"{event_name}_발표자료.pptx"
            st.session_state["ppt_count"] = len(slides_data)

    if st.session_state.get("ppt_bytes"):
        st.success(f"✅ {st.session_state.get('ppt_count', '')}개 슬라이드 생성 완료!")
        st.download_button(
            "⬇️ PPT 다운로드 (.pptx)",
            data=st.session_state["ppt_bytes"],
            file_name=st.session_state.get("ppt_name", "발표자료.pptx"),
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
        )


def _build_pptx(slides_data: list, summary: dict, theme: str = "네이비 (공공기관 정장)") -> bytes | None:
    """다크 배경 고품질 PPT 생성 (PDF 스타일)"""
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt, Emu
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN

        pal = THEMES.get(theme, THEMES["네이비 (공공기관 정장)"])

        def rgb(key):
            return RGBColor(*pal[key])

        def WHITE():
            return RGBColor(0xFF, 0xFF, 0xFF)

        prs = Presentation()
        prs.slide_width  = Inches(13.33)
        prs.slide_height = Inches(7.5)
        W = prs.slide_width
        H = prs.slide_height
        blank = prs.slide_layouts[6]

        def rect(slide, x, y, w, h, color_key=None, rgb_val=None):
            s = slide.shapes.add_shape(1, x, y, w, h)
            s.fill.solid()
            s.fill.fore_color.rgb = rgb(color_key) if color_key else RGBColor(*rgb_val)
            s.line.fill.background()
            return s

        def txt(slide, text, x, y, w, h, size, bold=False,
                color_key=None, rgb_val=None, align=PP_ALIGN.LEFT, italic=False, wrap=True):
            tb = slide.shapes.add_textbox(x, y, w, h)
            tf = tb.text_frame
            tf.word_wrap = wrap
            p = tf.paragraphs[0]
            p.alignment = align
            r = p.add_run()
            r.text = str(text)
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.name = MALGUN_GOTHIC
            if color_key:
                r.font.color.rgb = rgb(color_key)
            else:
                r.font.color.rgb = RGBColor(*(rgb_val or (255,255,255)))
            return tb

        def page_num(slide, n):
            tb = slide.shapes.add_textbox(Inches(12.6), Inches(7.1), Inches(0.6), Inches(0.3))
            p = tb.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT
            r = p.add_run()
            r.text = str(n)
            r.font.size = Pt(10)
            r.font.name = MALGUN_GOTHIC
            r.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

        for idx, si in enumerate(slides_data):
            slide    = prs.slides.add_slide(blank)
            layout   = si.get("layout", "content")
            title_t  = si.get("title", "")
            subtitle = si.get("subtitle", "")
            bullets  = si.get("bullets", [])
            headers  = si.get("headers", [])
            rows     = si.get("rows", [])
            snum     = idx + 1

            # ── 전체 공통 배경 ──────────────────────
            rect(slide, 0, 0, W, H, "bg")

            # ══ title ══════════════════════════════
            if layout == "title":
                # 우측 그라데이션 효과 대신 약간 밝은 사각형
                rect(slide, int(W*0.55), 0, int(W*0.45), H, "bg2")
                # 좌측 세로 액센트 바
                rect(slide, 0, 0, Inches(0.12), H, "accent2")
                # 하단 얇은 라인
                rect(slide, Inches(0.12), int(H*0.88), W, Inches(0.04), "accent")

                # 주최기관 (상단 좌)
                org = (summary or {}).get("주최기관", "")
                if org:
                    txt(slide, org,
                        Inches(0.5), Inches(0.35), Inches(8), Inches(0.5),
                        size=13, color_key="subtext")

                # 메인 제목 (좌측 중앙)
                txt(slide, title_t,
                    Inches(0.5), Inches(1.6), Inches(7.5), Inches(3.0),
                    size=40, bold=True, color_key="text")

                # accent 구분선
                rect(slide, Inches(0.5), Inches(4.7), Inches(4), Inches(0.05), "accent")

                # 부제목 (일시·장소)
                if subtitle:
                    txt(slide, subtitle,
                        Inches(0.5), Inches(4.85), Inches(7.5), Inches(1.0),
                        size=15, color_key="subtext")

            # ══ closing ════════════════════════════
            elif layout == "closing":
                rect(slide, int(W*0.55), 0, int(W*0.45), H, "bg2")
                rect(slide, 0, 0, Inches(0.12), H, "accent2")
                rect(slide, Inches(0.12), int(H*0.88), W, Inches(0.04), "accent")

                txt(slide, title_t,
                    Inches(0.5), Inches(2.0), Inches(7.5), Inches(2.0),
                    size=46, bold=True, color_key="text")

                rect(slide, Inches(0.5), Inches(4.2), Inches(3.5), Inches(0.05), "accent")

                if subtitle:
                    txt(slide, subtitle,
                        Inches(0.5), Inches(4.4), Inches(7.5), Inches(0.7),
                        size=18, color_key="subtext")

                # 담당부서
                dept = ""
                if summary:
                    d = summary.get("담당부서","")
                    p = summary.get("담당자","")
                    dept = f"{d}  {p}".strip()
                if dept:
                    txt(slide, dept,
                        Inches(0.5), Inches(6.9), Inches(8), Inches(0.4),
                        size=11, color_key="subtext", italic=True)

            # ══ section ════════════════════════════
            elif layout == "section":
                # 우측 패널
                rect(slide, int(W*0.6), 0, int(W*0.4), H, "bg2")
                # 좌측 두꺼운 액센트 바
                rect(slide, 0, 0, Inches(0.25), H, "accent")
                # 중앙 수평선
                rect(slide, Inches(0.5), int(H*0.62), Inches(7.5), Inches(0.04), "accent2")

                # 큰 섹션 번호
                txt(slide, f"{snum-1:02d}",
                    Inches(0.4), Inches(1.0), Inches(7.5), Inches(2.5),
                    size=110, bold=True, color_key="accent",
                    align=PP_ALIGN.LEFT)

                # 섹션 제목
                txt(slide, title_t,
                    Inches(0.5), Inches(3.8), Inches(8.5), Inches(1.5),
                    size=32, bold=True, color_key="text")

            # ══ table ══════════════════════════════
            elif layout == "table" and headers and rows:
                # 상단 헤더 패널
                rect(slide, 0, 0, W, Inches(1.3), "bg2")
                rect(slide, 0, 0, Inches(0.12), H, "accent")
                rect(slide, Inches(0.12), Inches(1.25), W, Inches(0.05), "accent")

                txt(slide, title_t,
                    Inches(0.3), Inches(0.22), Inches(12.5), Inches(0.8),
                    size=26, bold=True, color_key="text")

                col_count = len(headers)
                row_count = len(rows) + 1
                table = slide.shapes.add_table(
                    row_count, col_count,
                    Inches(0.3), Inches(1.4), Inches(12.7), Inches(5.7)
                ).table

                col_w = Inches(12.7) // col_count
                for c in range(col_count):
                    table.columns[c].width = col_w

                # 헤더 행
                for c, hdr in enumerate(headers):
                    cell = table.cell(0, c)
                    cell.text = str(hdr)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = rgb("accent")
                    p2 = cell.text_frame.paragraphs[0]
                    p2.alignment = PP_ALIGN.CENTER
                    run = p2.runs[0] if p2.runs else p2.add_run()
                    run.font.bold = True
                    run.font.size = Pt(14)
                    run.font.color.rgb = WHITE()
                    run.font.name = MALGUN_GOTHIC

                # 데이터 행
                for r, row_data in enumerate(rows):
                    row_bg = pal["card"] if r % 2 == 0 else pal["bg2"]
                    for c, val in enumerate(row_data[:col_count]):
                        cell = table.cell(r+1, c)
                        cell.text = str(val)
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(*row_bg)
                        p2 = cell.text_frame.paragraphs[0]
                        p2.alignment = PP_ALIGN.CENTER
                        run = p2.runs[0] if p2.runs else p2.add_run()
                        run.font.size = Pt(13)
                        run.font.color.rgb = rgb("text")
                        run.font.name = MALGUN_GOTHIC

                page_num(slide, snum)

            # ══ content (기본) ═════════════════════
            else:
                # 상단 패널
                rect(slide, 0, 0, W, Inches(1.3), "bg2")
                rect(slide, 0, 0, Inches(0.12), H, "accent")
                rect(slide, Inches(0.12), Inches(1.25), W, Inches(0.05), "accent2")

                txt(slide, title_t,
                    Inches(0.3), Inches(0.22), Inches(12.5), Inches(0.8),
                    size=26, bold=True, color_key="text")

                if bullets:
                    is_toc = any(str(b).startswith(("01","02","1.","2.")) for b in bullets)

                    if is_toc:
                        # 목차: 카드 스타일로 배치
                        per_row = 2
                        card_w = Inches(5.8)
                        card_h = Inches(1.4)
                        for i, b in enumerate(bullets[:6]):
                            cx = Inches(0.5) + (i % per_row) * Inches(6.7)
                            cy = Inches(1.5) + (i // per_row) * Inches(1.6)
                            rect(slide, cx, cy, card_w, card_h, "card")
                            # 좌측 포인트 바
                            rect(slide, cx, cy, Inches(0.08), card_h, "accent")
                            txt(slide, str(b),
                                cx + Inches(0.2), cy + Inches(0.3),
                                card_w - Inches(0.3), Inches(0.8),
                                size=18, bold=True, color_key="text")
                    else:
                        # 일반 불릿
                        tb = slide.shapes.add_textbox(
                            Inches(0.4), Inches(1.45), Inches(12.5), Inches(5.7)
                        )
                        tf = tb.text_frame
                        tf.word_wrap = True
                        for i, b in enumerate(bullets):
                            p2 = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                            p2.space_before = Pt(8)
                            p2.space_after  = Pt(4)
                            r2 = p2.add_run()
                            r2.text = f"  ▸  {b}"
                            r2.font.size = Pt(18)
                            r2.font.bold = False
                            r2.font.color.rgb = rgb("text")
                            r2.font.name = MALGUN_GOTHIC

                page_num(slide, snum)

        buf = io.BytesIO()
        prs.save(buf)
        return buf.getvalue()

    except ImportError:
        st.error("python-pptx 패키지가 필요합니다: pip install python-pptx")
        return None
    except Exception as e:
        st.error(f"PPT 생성 오류: {e}")
        import traceback
        st.text(traceback.format_exc())
        return None


def render_namecard():
    st.subheader("🪪 명찰 자동생성")
    summary = st.session_state.get("plan_summary_dict", {})
    event_name = summary.get("행사명", "행사") if summary else "행사"

    st.info("엑셀 파일에 **이름**, **소속** 컬럼이 있어야 합니다.")

    col1, col2 = st.columns(2)
    with col1:
        excel_file = st.file_uploader("명단 엑셀 업로드", type=["xlsx", "xls"], key="namecard_excel")
    with col2:
        schedule_text = st.text_area(
            "뒷면 일정표 내용",
            placeholder="09:00 등록\n10:00 개회식\n11:00 주요프로그램\n12:00 폐회",
            height=150,
            key="namecard_schedule",
        )

    custom_event = st.text_input("행사명 (비워두면 계획서에서 자동)", value="", placeholder=event_name, key="namecard_event")
    final_event = custom_event.strip() or event_name

    if excel_file and st.button("🪪 명찰 생성", key="namecard_gen"):
        try:
            from openpyxl import load_workbook
            wb = load_workbook(io.BytesIO(excel_file.read()), data_only=True)
            ws = wb.active

            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
            name_idx = next(
                (i for i, h in enumerate(headers) if any(k in h for k in ["이름", "성명", "name", "Name"])),
                None,
            )
            org_idx = next(
                (i for i, h in enumerate(headers) if any(k in h for k in ["소속", "기관", "org", "Org", "dept", "Dept"])),
                None,
            )

            if name_idx is None:
                st.error(f"'이름' 또는 '성명' 컬럼을 찾을 수 없습니다. 현재 컬럼: {headers}")
                return
            if org_idx is None:
                st.warning("'소속' 컬럼을 찾을 수 없습니다. 소속 없이 생성합니다.")

            persons = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                name = str(row[name_idx]).strip() if row[name_idx] is not None else ""
                org = str(row[org_idx]).strip() if org_idx is not None and row[org_idx] is not None else ""
                if name and name.lower() != "none":
                    persons.append({"name": name, "org": org})

            if not persons:
                st.error("명단에 유효한 데이터가 없습니다.")
                return

            st.info(f"총 {len(persons)}명 처리 중...")
            docx_bytes = _build_namecard_docx(persons, final_event, schedule_text)
            if docx_bytes:
                st.session_state["namecard_bytes"] = docx_bytes
                st.session_state["namecard_count"] = len(persons)

        except Exception as e:
            st.error(f"명단 처리 오류: {e}")
            import traceback
            st.text(traceback.format_exc())

    if st.session_state.get("namecard_bytes"):
        count = st.session_state.get("namecard_count", 0)
        st.success(f"✅ {count}명 명찰 생성 완료!")
        st.download_button(
            f"⬇️ 명찰 다운로드 ({count}명)",
            data=st.session_state["namecard_bytes"],
            file_name="명찰.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


def _build_namecard_docx(persons: list, event_name: str, schedule_text: str) -> bytes | None:
    try:
        from docx import Document
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        doc = Document()
        for section in doc.sections:
            section.page_width = Cm(21)
            section.page_height = Cm(29.7)
            section.top_margin = Cm(1.5)
            section.bottom_margin = Cm(1.5)
            section.left_margin = Cm(2.0)
            section.right_margin = Cm(2.0)

        def add_dashed_line():
            hr = doc.add_paragraph()
            hr.paragraph_format.space_before = Pt(0)
            hr.paragraph_format.space_after = Pt(0)
            pPr = hr._p.get_or_add_pPr()
            pBdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "dashed")
            bottom.set(qn("w:sz"), "6")
            bottom.set(qn("w:space"), "1")
            bottom.set(qn("w:color"), "888888")
            pBdr.append(bottom)
            pPr.append(pBdr)

        def add_card_front(name: str, org: str, first: bool):
            if not first:
                add_dashed_line()

            p_event = doc.add_paragraph()
            p_event.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_event.paragraph_format.space_before = Pt(28)
            p_event.paragraph_format.space_after = Pt(8)
            r = p_event.add_run(event_name)
            r.font.name = MALGUN_GOTHIC
            r.font.size = Pt(15)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

            if org:
                p_org = doc.add_paragraph()
                p_org.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_org.paragraph_format.space_before = Pt(4)
                p_org.paragraph_format.space_after = Pt(4)
                r_org = p_org.add_run(org)
                r_org.font.name = MALGUN_GOTHIC
                r_org.font.size = Pt(17)
                r_org.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

            p_name = doc.add_paragraph()
            p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_name.paragraph_format.space_before = Pt(8)
            p_name.paragraph_format.space_after = Pt(28)
            r_name = p_name.add_run(name)
            r_name.font.name = MALGUN_GOTHIC
            r_name.font.size = Pt(38)
            r_name.font.bold = True
            r_name.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)

        def add_card_back(first: bool):
            if not first:
                add_dashed_line()

            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_title.paragraph_format.space_before = Pt(28)
            p_title.paragraph_format.space_after = Pt(10)
            r = p_title.add_run("행사 일정")
            r.font.name = MALGUN_GOTHIC
            r.font.size = Pt(15)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

            lines = schedule_text.strip().split("\n") if schedule_text.strip() else ["일정 미정"]
            for line in lines:
                p_s = doc.add_paragraph()
                p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_s.paragraph_format.space_before = Pt(3)
                p_s.paragraph_format.space_after = Pt(3)
                r_s = p_s.add_run(line.strip())
                r_s.font.name = MALGUN_GOTHIC
                r_s.font.size = Pt(13)

            p_end = doc.add_paragraph()
            p_end.paragraph_format.space_after = Pt(28)

        for i in range(0, len(persons), 2):
            batch = persons[i: i + 2]

            add_card_front(batch[0]["name"], batch[0]["org"], True)
            if len(batch) > 1:
                add_card_front(batch[1]["name"], batch[1]["org"], False)

            doc.add_page_break()

            add_card_back(True)
            if len(batch) > 1:
                add_card_back(False)

            if i + 2 < len(persons):
                doc.add_page_break()

        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    except ImportError:
        st.error("python-docx 패키지가 필요합니다: pip install python-docx")
        return None
    except Exception as e:
        st.error(f"명찰 생성 오류: {e}")
        import traceback
        st.text(traceback.format_exc())
        return None


# ─────────────────────────────────────────────
# 7. 행사장 약도
# ─────────────────────────────────────────────
def render_map():
    st.subheader("🗺️ 행사장 약도")
    summary = st.session_state.get("plan_summary_dict", {})
    default_place = summary.get("장소", "") if summary else ""

    place_name = st.text_input(
        "장소명 또는 주소 입력",
        value=default_place,
        placeholder="예: 부산 벡스코  또는  부산광역시 해운대구 APEC로 55",
        key="map_place",
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        zoom = st.slider("지도 확대 수준", 1, 14, 4, key="map_zoom",
                         help="숫자가 클수록 더 확대됩니다")
    with col2:
        map_w = st.number_input("가로 (px)", value=800, min_value=200, max_value=1200, key="map_w")
    with col3:
        map_h = st.number_input("세로 (px)", value=600, min_value=200, max_value=900, key="map_h")

    if st.button("🗺️ 약도 생성", key="map_gen") and place_name:
        try:
            kakao_key = st.secrets["KAKAO_API_KEY"]
            hdrs = {"Authorization": f"KakaoAK {kakao_key}"}

            def kw_search(q):
                r = requests.get(
                    "https://dapi.kakao.com/v2/local/search/keyword.json",
                    headers=hdrs, params={"query": q}
                )
                docs = r.json().get("documents", [])
                return docs[0] if docs else None

            def addr_search(q):
                r = requests.get(
                    "https://dapi.kakao.com/v2/local/search/address.json",
                    headers=hdrs, params={"query": q}
                )
                docs = r.json().get("documents", [])
                if not docs:
                    return None
                d = docs[0]
                ra = d.get("road_address") or {}
                a  = d.get("address") or {}
                return {
                    "x": d.get("x"),
                    "y": d.get("y"),
                    "place_name": q,
                    "road_address_name": ra.get("address_name") or a.get("address_name", q),
                }

            # 괄호 안 주소 추출: "장소명 (도로명주소)" 패턴 처리
            paren = re.search(r"\(([^)]+)\)", place_name)
            paren_addr = paren.group(1).strip() if paren else None
            clean_name = re.sub(r"\s*\([^)]*\)", "", place_name).strip()

            doc0 = None

            # 1차: 괄호 안 주소로 주소검색
            if paren_addr:
                st.info(f"🔍 괄호 내 주소 감지: **{paren_addr}**")
                doc0 = addr_search(paren_addr)

            # 2차: 장소명으로 키워드 검색
            if not doc0:
                doc0 = kw_search(clean_name)

            # 3차: 원본 전체로 키워드 검색
            if not doc0:
                doc0 = kw_search(place_name)

            # 4차: AI 주소 보정
            if not doc0:
                with st.spinner("AI가 주소를 보정 중..."):
                    fixed = gemini_text(
                        f"다음 장소명의 정확한 도로명 주소만 한 줄로 출력하세요.\n"
                        f"예) 부산광역시 동구 중앙대로 206\n장소명: {clean_name}"
                    )
                st.info(f"🔍 AI 보정: **{fixed}**")
                doc0 = kw_search(fixed) or addr_search(fixed)

            if not doc0 or not doc0.get("x"):
                st.error("위치를 찾을 수 없습니다. 도로명 주소를 직접 입력해보세요.")
                st.caption("예) 부산광역시 동구 중앙대로 206")
                return

            lng = float(doc0["x"])
            lat = float(doc0["y"])
            found_name = doc0.get("place_name", clean_name)
            found_addr = doc0.get("road_address_name", "")
            st.success(f"📍 찾은 장소: **{found_name}** ({found_addr})")

            # Static Map API (dapi.kakao.com 정식 엔드포인트)
            map_resp = requests.get(
                "https://dapi.kakao.com/v2/maps/staticmap",
                headers=hdrs,
                params={
                    "center": f"{lng},{lat}",
                    "level": zoom,
                    "w": int(map_w),
                    "h": int(map_h),
                    "markers": f"color:red|{lng},{lat}",
                }
            )

            ct = map_resp.headers.get("Content-Type", "")
            if map_resp.status_code == 200 and "image" in ct:
                st.session_state["map_img"] = map_resp.content
                st.session_state["map_name"] = found_name
            else:
                # 대체: OpenStreetMap 정적 지도 (무료, 인증 불필요)
                osm_url = (
                    f"https://staticmap.openstreetmap.de/staticmap.php"
                    f"?center={lat},{lng}&zoom={zoom+2}"
                    f"&size={int(map_w)}x{int(map_h)}"
                    f"&markers={lat},{lng},red-pushpin"
                )
                osm_resp = requests.get(osm_url, timeout=10)
                if osm_resp.status_code == 200 and "image" in osm_resp.headers.get("Content-Type", ""):
                    st.session_state["map_img"] = osm_resp.content
                    st.session_state["map_name"] = found_name
                    st.caption("※ OpenStreetMap 기반 지도")
                else:
                    from urllib.parse import quote
                    st.warning(f"지도 이미지 생성 실패. 카카오맵 링크를 제공합니다.")
                    st.markdown(f"[🗺️ 카카오맵에서 보기](https://map.kakao.com/?q={quote(found_addr or clean_name)})")
                    st.caption(f"좌표: 위도 {lat:.6f}, 경도 {lng:.6f}")

        except KeyError:
            st.error("KAKAO_API_KEY가 st.secrets에 설정되어 있지 않습니다.")
        except Exception as e:
            st.error(f"지도 생성 오류: {e}")
            import traceback
            st.text(traceback.format_exc())

    if st.session_state.get("map_img"):
        found_name = st.session_state.get("map_name", "행사장")
        st.image(st.session_state["map_img"], caption=f"📍 {found_name}", use_container_width=True)
        st.download_button(
            "⬇️ 약도 이미지 다운로드",
            data=st.session_state["map_img"],
            file_name=f"행사장약도_{found_name}.png",
            mime="image/png",
        )


# ─────────────────────────────────────────────
# 메인 탭 렌더러
# ─────────────────────────────────────────────
def render_tab7():
    st.title("🤖 AI 문서 자동생성")
    st.caption("HWPX 계획서를 업로드하면 각종 문서를 자동으로 생성합니다.")

    st.markdown("---")
    st.subheader("📁 행사 계획서 업로드")

    hwpx_file = st.file_uploader(
        "HWPX 파일 업로드 (.hwpx / .hwp)",
        type=["hwpx", "hwp"],
        key="hwpx_upload",
    )

    if hwpx_file:
        file_key = f"{hwpx_file.name}_{hwpx_file.size}"
        already_summarized = (
            st.session_state.get("hwpx_file_key") == file_key
            and "plan_summary_raw" in st.session_state
        )

        if not already_summarized:
            st.session_state["hwpx_file_key"] = file_key
            with st.spinner("계획서 텍스트 추출 및 요약 중..."):
                raw_text = extract_hwpx_text(hwpx_file)
                if len(raw_text) < 100:
                    st.error("계획서에서 텍스트를 충분히 추출하지 못했습니다.")
                else:
                    summary_raw = summarize_plan(raw_text)
                    st.session_state["plan_summary_raw"] = summary_raw
                    try:
                        clean = re.sub(r"```json|```", "", summary_raw).strip()
                        st.session_state["plan_summary_dict"] = json.loads(clean)
                    except Exception:
                        st.session_state["plan_summary_dict"] = {}

        if "plan_summary_dict" in st.session_state:
            summary = st.session_state["plan_summary_dict"]
            with st.expander("📋 계획서 요약 확인 (모든 기능에 재사용됨)", expanded=True):
                if summary:
                    cols = st.columns(2)
                    keys = list(summary.keys())
                    half = len(keys) // 2
                    for i, k in enumerate(keys):
                        v = summary[k]
                        if isinstance(v, list):
                            v = ", ".join(str(x) for x in v)
                        with cols[0 if i < half else 1]:
                            st.markdown(f"**{k}:** {v}")
                else:
                    st.text(st.session_state.get("plan_summary_raw", "요약 없음"))
    else:
        st.info("HWPX 계획서를 업로드하면 아래 기능들이 활성화됩니다. 업로드 없이도 수동으로 사용 가능합니다.")
        if "plan_summary_dict" not in st.session_state:
            st.session_state["plan_summary_dict"] = {}

    st.markdown("---")
    tabs = st.tabs([
        "📄 문서 4종",
        "🎤 사회자 멘트",
        "🪧 현수막",
        "📋 결과보고서",
        "📊 PPT",
        "🪪 명찰",
        "🗺️ 행사장 약도",
    ])

    with tabs[0]: render_doc4()
    with tabs[1]: render_mc()
    with tabs[2]: render_banner()
    with tabs[3]: render_result_report()
    with tabs[4]: render_ppt()
    with tabs[5]: render_namecard()
    with tabs[6]: render_map()


# main.py 호환용 별칭
render = render_tab7

if __name__ == "__main__":
    st.set_page_config(page_title="AI 문서 자동생성", page_icon="🤖", layout="wide")
    render_tab7()
