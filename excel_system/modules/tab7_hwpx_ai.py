"""
tab7_hwpx_ai.py  ·  행사 계획서 기반 AI 문서 자동생성
리팩토링: 2026-07
"""

import streamlit as st
import google.generativeai as genai
import zipfile, io, re, json, requests
from pathlib import Path

# ── 상수 ──────────────────────────────────────────────────────────────
GEMINI_MODEL  = "gemini-2.5-flash-lite"
FONT          = "맑은 고딕"

# PPT 테마 (bg/bg2/card/accent/accent2/text/subtext)
THEMES = {
    "네이비 (공공기관 정장)": {
        "bg":      (0x12, 0x1A, 0x2E), "bg2":    (0x1F, 0x38, 0x64),
        "card":    (0x16, 0x2A, 0x4E), "accent":  (0x2E, 0x74, 0xB5),
        "accent2": (0x70, 0xAD, 0x47), "text":    (0xE8, 0xF0, 0xFE),
        "subtext": (0x8A, 0xAD, 0xD4),
    },
    "다크블루 (격식)": {
        "bg":      (0x0D, 0x1B, 0x2A), "bg2":    (0x1B, 0x4F, 0x72),
        "card":    (0x15, 0x35, 0x50), "accent":  (0x1B, 0x4F, 0x72),
        "accent2": (0xF3, 0x9C, 0x12), "text":    (0xEC, 0xF0, 0xF1),
        "subtext": (0xAB, 0xC4, 0xD8),
    },
    "그린 (환경/생태)": {
        "bg":      (0x0A, 0x1A, 0x0F), "bg2":    (0x1E, 0x4D, 0x2B),
        "card":    (0x14, 0x35, 0x1D), "accent":  (0x27, 0xAE, 0x60),
        "accent2": (0xF3, 0x9C, 0x12), "text":    (0xE8, 0xF8, 0xEA),
        "subtext": (0xA9, 0xD9, 0xB4),
    },
    "버건디 (품격)": {
        "bg":      (0x1A, 0x08, 0x0C), "bg2":    (0x6B, 0x21, 0x2C),
        "card":    (0x2E, 0x10, 0x15), "accent":  (0xC0, 0x39, 0x4B),
        "accent2": (0xE8, 0xB4, 0x6A), "text":    (0xFD, 0xF0, 0xF1),
        "subtext": (0xF5, 0xC6, 0xCC),
    },
}

# 보도자료 내장 예시 (참고파일 없을 때 폴백)
_PRESS_EXAMPLE = """
【보 도 자 료】
담당부서: 공간정보제도과  담당자: 홍길동  연락처: 044-000-0000

○○부, 지적재조사 담당자 역량강화 교육 실시
- 전국 시·군·구 담당 공무원 150명 대상, 최신 측량기술 집중 교육

국토교통부(장관 ○○○)는 10일 정부세종청사에서 전국 시·군·구 지적재조사 담당 공무원 150명을 대상으로 역량강화 교육을 실시했다.

이번 교육은 지적재조사사업의 현장 추진력을 높이기 위해 마련됐다. 드론 측량, 3D 공간정보 활용 등 최신 기술을 중심으로 진행됐으며, 우수사례 발표와 현장 실습도 병행됐다.

○○○ 국토교통부 지적재조사단장은 "이번 교육이 현장 담당자들의 실무 역량을 높이고 사업 추진에 실질적인 도움이 되길 바란다"고 말했다.

문의: 국토교통부 공간정보제도과 (044-000-0000)
"""

# ── AI 공통 ───────────────────────────────────────────────────────────
def get_model():
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    return genai.GenerativeModel(GEMINI_MODEL)

def ai(prompt: str) -> str:
    return get_model().generate_content(prompt).text.strip()

# ── HWPX 처리 ─────────────────────────────────────────────────────────
def hwpx_to_text(f) -> str:
    parts = []
    try:
        with zipfile.ZipFile(io.BytesIO(f.read()), "r") as zf:
            targets = sorted([n for n in zf.namelist()
                              if n.endswith(".xml") and ("Contents" in n or "content" in n.lower())])
            if not targets:
                targets = [n for n in zf.namelist() if n.endswith(".xml")]
            for name in targets:
                try:
                    raw = zf.read(name).decode("utf-8", errors="ignore")
                    clean = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", raw)).strip()
                    if len(clean) > 50:
                        parts.append(clean)
                except Exception:
                    pass
    except Exception as e:
        return f"[읽기 오류: {e}]"
    full = " ".join(parts)
    return full[:8000]

def summarize_hwpx(raw: str) -> str:
    return ai(f"""너는 20년 차 베테랑 공무원이다.
아래 행사 계획서 원문을 분석해 핵심 정보를 순수 JSON으로만 반환하라 (마크다운 없이).

원문: {raw}

JSON 형식:
{{"행사명":"","주최기관":"","주관기관":"","일시":"","장소":"","대상":"",
"참석예정인원":"","행사목적":"","주요프로그램":[""],
"예산":"","담당부서":"","담당자":"","기타특이사항":""}}""")

# ── 공통 UI ───────────────────────────────────────────────────────────
def show_textarea(sess_key: str, label: str, height: int = 420):
    if st.session_state.get(sess_key):
        st.caption("📋 텍스트 박스 클릭 → Ctrl+A → Ctrl+C → 한글/워드에 붙여넣기")
        st.text_area(label, value=st.session_state[sess_key],
                     height=height, key=f"ta_{sess_key}")

def gen_button(label: str, key: str, prompt: str, sess_key: str, height: int = 420):
    if st.button(label, key=key):
        with st.spinner("생성 중..."):
            st.session_state[sess_key] = ai(prompt)
    show_textarea(sess_key, label.replace("✍️ ", ""), height)

# ── 파일 텍스트 추출 (보도자료 참고용) ──────────────────────────────────
def extract_file_text(f) -> str:
    """txt / hwpx / pdf 에서 텍스트 추출 (최대 2,000자 — 보도자료 약 2장 분량)"""
    name = f.name.lower()
    try:
        if name.endswith(".txt"):
            return f.read().decode("utf-8", errors="ignore")[:2000]
        elif name.endswith((".hwpx", ".hwp")):
            return hwpx_to_text(f)[:2000]
        elif name.endswith(".pdf"):
            try:
                import fitz
                doc = fitz.open(stream=f.read(), filetype="pdf")
                return "\n".join(p.get_text() for p in doc)[:2000]
            except ImportError:
                return f.read().decode("utf-8", errors="ignore")[:2000]
    except Exception as e:
        return f"[추출 실패: {e}]"
    return ""

# ── 프롬프트 빌더 ─────────────────────────────────────────────────────
def _summary_str() -> str:
    return json.dumps(st.session_state.get("plan_summary_dict", {}),
                      ensure_ascii=False, indent=2)

def _prompt_report(s):
    return f"""너는 20년 차 베테랑 공무원이다. 신뢰감 있고 건조한 행정 공문서 문체로 작성하라.
감성적 수식어 배제. 개조식(○ 기호), 마크다운 없이 순수 텍스트.

[포함 항목] ○ 행사 개요 ○ 추진 목적 ○ 주요 내용 ○ 참석 대상·규모 ○ 소요 예산 ○ 기대 효과 ○ 향후 계획

행사 계획서 요약: {s}"""

def _prompt_director(s):
    return f"""너는 20년 차 베테랑 공무원이다. 국장급 인사말씀을 작성하라.
- 행사와 기관 정책 방향의 연계성 강조
- 내빈에게 정중하되 건조하고 신뢰감 있는 문체
- 서두 인삿말 포함 (따뜻한 어투 허용), 400~600자, 마크다운 없이

[구성] 1.서두 인삿말 2.행사의 행정적 의미·정책 연계 3.기관 비전 제시 4.마무리 당부

행사 계획서 요약: {s}"""

def _prompt_manager(s):
    return f"""너는 20년 차 베테랑 공무원이다. 과장급 인사말씀을 작성하라.
- 실무 총괄 과장 버전, 행사 준비 상황 언급 포함
- 참가자 실무적 당부·협조 요청
- 서두 인삿말 포함 (따뜻한 어투 허용), 300~500자, 마크다운 없이

[구성] 1.서두 2.준비 과정·노고 치하 3.목적·주요 내용 안내 4.당부·협조 요청 5.마무리

행사 계획서 요약: {s}"""

def _prompt_press(s, ref_texts):
    if ref_texts:
        examples = "\n\n--- 참고자료 구분 ---\n\n".join(
            f"[참고자료 {i+1}]\n{t}" for i, t in enumerate(ref_texts))
        style = f"""[1단계 - 문체 학습용 (내용 말고 스타일만 참고)]
{examples}

[2단계 - 실제 작성 주제 (아래 내용으로 보도자료를 써라)]"""
    else:
        style = f"""[참고 예시 - 문체·형식만 참고]
{_PRESS_EXAMPLE}
--- 예시 끝 ---

[실제 작성 주제]"""

    return f"""너는 20년 차 베테랑 공무원이자 언론홍보 전문가다.

{style}
행사 계획서 요약 (← 이 내용으로 보도자료를 작성할 것):
{s}

[필수 구조]
1) 담당부서 / 담당자 / 연락처
2) 제목: 위 행사의 핵심 성과 한 문장
3) 부제: 대시(-) 핵심 포인트 1~2개
4) 본문: 육하원칙으로 행사 기술 → 추진 배경 → 목적 → 기대효과
5) 기관장 인용구: ○○○은 "..."이라고 말했다.
6) 문의처

[원칙] 참고자료 내용은 절대 쓰지 않는다 / '~했다','~한다' 기사체 / 마크다운 없이"""

def _prompt_mc(s, order, tone):
    return f"""너는 20년 차 베테랑 공무원이자 공공기관 전문 사회자다.
[톤]: {tone}  [순서]: {', '.join(order)}

[원칙]
- 각 순서는 ## 헤딩으로 구분
- 현장에서 바로 읽을 수 있는 완성된 문장, 각 100~200자
- 다음 순서로 자연스럽게 넘어가는 전환 문구 포함
- 내빈소개: 직책 → 성함 순

행사 계획서 요약: {s}"""

def _prompt_banner(s, n, style):
    return f"""너는 20년 차 베테랑 공무원이자 공공기관 홍보 전문가다.
현수막 문구 시안 {n}개를 작성하라.
[스타일]: {style}
[원칙] 15~25자 이내 / 번호. 문구 형식 / 각 문구 아래 사용 설명 한 줄 / 모호한 표현 배제
행사 계획서 요약: {s}"""

def _prompt_result(s, attendance, satisfaction, note):
    return f"""너는 20년 차 베테랑 공무원이다. 행사 결과보고서 초안을 작성하라.
감성 배제, 수치·사실 위주, 개조식(○), 마크다운 없이.

[결과 정보] 참석인원:{attendance or '미기재'} / 만족도:{satisfaction or '미기재'} / 특이사항:{note or '없음'}

[형식]
1.행사 개요 (○ 행사명/일시/장소/주최·주관/참석인원)
2.추진 결과 (○ 참석현황 계획대비 / 프로그램별 결과 / 만족도)
3.예산 집행 (○ 편성/집행/집행률)
4.주요 성과 및 시사점 (○ 성과 / 미흡사항·개선방향)
5.향후 계획 (○ 후속조치)

행사 계획서 요약: {s}"""

def _prompt_ppt(s, n):
    return (
        """너는 20년 차 베테랑 공무원이자 공공기관 발표자료 전문가다.
감성 배제, 수치·사실 위주, 담당자가 바로 발표할 수 있는 수준으로 작성하라.
제목만 있는 슬라이드 금지 — 모든 슬라이드에 body(본문 2~3문장)를 반드시 채워라.

[레이아웃 7종]
title: 표지 (title, subtitle)
closing: 마무리 (title, subtitle)
section: 챕터 구분 (title만)
content: 제목+본문+불릿 (title, body, bullets[])
two_column: 좌우비교 (title, body, left_title, left_bullets[], right_title, right_bullets[])
highlight: 숫자강조 (title, body, stat_number, stat_label, bullets[])
table: 표 (title, body, headers[], rows[][])

[필수 순서]
표지(title) → 목차(content, 01.형식) → 행사개요(highlight) → 추진배경(two_column) → 주요프로그램(table) → 세부내용(content) → 기대효과(content) → 클로징(closing)

순수 JSON 배열만 반환 (마크다운 없이)."""
        + f"\n\n슬라이드 수: {n}개\n행사 계획서 요약:\n{s}"
    )

# ── 1. 문서 4종 ───────────────────────────────────────────────────────
def render_doc4():
    st.subheader("📄 문서 4종 자동생성")
    st.caption("생성 후 Ctrl+A → Ctrl+C → 한글/워드에 붙여넣기")
    s = _summary_str()

    # 보도자료 참고파일 업로드
    with st.expander("📎 보도자료 참고파일 업로드 (선택 · 품질 향상)", expanded=False):
        st.caption("우리 기관 실제 보도자료를 올리면 문체·형식을 그대로 따라씁니다.")
        ref_files = st.file_uploader(
            "참고 보도자료 (txt / hwpx / pdf, 최대 3개)",
            type=["txt", "hwpx", "hwp", "pdf"],
            accept_multiple_files=True,
            key="press_ref_files",
        )
        if ref_files:
            cnt = min(len(ref_files), 3)
            st.success(f"✅ {cnt}개 반영됨 (3개 초과 시 앞 3개만 사용)")

    # 참고 텍스트 캐싱
    fkey = str([f.name for f in ref_files] if ref_files else [])
    if ref_files and st.session_state.get("_press_fkey") != fkey:
        with st.spinner("참고 보도자료 추출 중..."):
            st.session_state["press_ref_texts"] = [
                t for f in ref_files[:3]  # 최대 3개
                if (t := extract_file_text(f)) and len(t) > 50
            ]
        st.session_state["_press_fkey"] = fkey
    elif not ref_files:
        st.session_state["press_ref_texts"] = []

    ref_texts = st.session_state.get("press_ref_texts", [])

    col1, col2 = st.columns(2)
    with col1:
        gen_button("✍️ 요약보고서 생성", "btn_report",  _prompt_report(s),   "doc_report")
        st.markdown("---")
        gen_button("✍️ 과장인사말 생성", "btn_manager", _prompt_manager(s),  "doc_manager")
    with col2:
        gen_button("✍️ 국장인사말 생성", "btn_director","btn_director" and _prompt_director(s), "doc_director")
        st.markdown("---")
        gen_button("✍️ 보도자료 생성",   "btn_press",   _prompt_press(s, ref_texts), "doc_press", height=500)

# ── 2. 사회자 멘트 ────────────────────────────────────────────────────
MC_DEFAULT = ["개회선언","국민의례","내빈소개","기관장인사말","축사","주요프로그램 소개","폐회선언"]

def render_mc():
    st.subheader("🎤 사회자 멘트 자동생성")
    s = _summary_str()
    with st.expander("⚙️ 순서 편집", expanded=False):
        edited = st.text_area("순서 (한 줄에 하나)", "\n".join(MC_DEFAULT),
                              height=200, key="mc_order_edit")
        order = [l.strip() for l in edited.split("\n") if l.strip()]
    st.info("순서: " + " → ".join(order))
    tone = st.selectbox("멘트 톤", ["격식체 (공식 행사)","친근체 (소규모/내부 행사)","방송체 (대규모/공개 행사)"])
    gen_button("✍️ 사회자 멘트 생성", "btn_mc", _prompt_mc(s, order, tone), "doc_mc", height=550)

# ── 3. 현수막 ─────────────────────────────────────────────────────────
def render_banner():
    st.subheader("🪧 현수막 문구 시안")
    s = _summary_str()
    c1, c2 = st.columns(2)
    n     = c1.slider("시안 수", 3, 7, 5)
    style = c2.selectbox("스타일", ["공식/격식형","친근/따뜻형","역동/강조형","혼합 (다양하게)"])
    gen_button("✍️ 현수막 문구 생성", "btn_banner", _prompt_banner(s, n, style), "doc_banner", height=350)

# ── 4. 결과보고서 ─────────────────────────────────────────────────────
def render_result():
    st.subheader("📋 결과보고서 초안")
    s = _summary_str()
    c1, c2 = st.columns(2)
    attendance   = c1.text_input("실제 참석 인원",    placeholder="예: 150명")
    satisfaction = c1.text_input("만족도 조사 결과",  placeholder="예: 4.2/5.0")
    note         = c2.text_area("주요 성과·특이사항", placeholder="예) 전년 대비 참석률 20% 증가", height=105)
    gen_button("✍️ 결과보고서 초안 생성", "btn_result",
               _prompt_result(s, attendance, satisfaction, note), "doc_result", height=550)

# ── 5. PPT ────────────────────────────────────────────────────────────
def render_ppt():
    st.subheader("📊 PPT 자동생성")
    st.caption("행사 계획서 기반으로 현장 사용 가능한 발표자료를 생성합니다.")
    s = _summary_str()
    c1, c2 = st.columns([1, 2])
    n     = c1.slider("슬라이드 수", 8, 20, 12)
    theme = c2.selectbox("색상 테마", list(THEMES.keys()))

    if st.button("🖥️ PPT 생성", type="primary"):
        with st.spinner("AI가 슬라이드 구성 중... (10~20초)"):
            raw = ai(_prompt_ppt(s, n))
        try:
            slides = json.loads(re.sub(r"```json|```", "", raw).strip())
        except Exception as e:
            st.error(f"파싱 오류: {e}")
            with st.expander("AI 원본"):
                st.text(raw[:1000])
            return
        with st.spinner("PPT 파일 생성 중..."):
            data = _build_pptx(slides, st.session_state.get("plan_summary_dict", {}), theme)
        if data:
            name = st.session_state.get("plan_summary_dict", {}).get("행사명", "행사")
            st.session_state.update({"ppt_bytes": data, "ppt_name": f"{name}_발표자료.pptx",
                                     "ppt_count": len(slides)})

    if st.session_state.get("ppt_bytes"):
        st.success(f"✅ {st.session_state['ppt_count']}개 슬라이드 생성 완료!")
        st.download_button("⬇️ PPT 다운로드", data=st.session_state["ppt_bytes"],
                           file_name=st.session_state["ppt_name"],
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                           type="primary")

def _build_pptx(slides_data, summary, theme="네이비 (공공기관 정장)"):
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN

        pal = THEMES.get(theme, THEMES["네이비 (공공기관 정장)"])
        def C(k): return RGBColor(*pal[k])
        def W():  return RGBColor(0xFF,0xFF,0xFF)

        prs = Presentation()
        prs.slide_width  = Inches(13.33)
        prs.slide_height = Inches(7.5)
        W_  = prs.slide_width
        H_  = prs.slide_height
        BL  = prs.slide_layouts[6]

        def R(sl, x,y,w,h, k):
            s = sl.shapes.add_shape(1,x,y,w,h)
            s.fill.solid(); s.fill.fore_color.rgb = C(k); s.line.fill.background()
        def T(sl, text, x,y,w,h, sz, bold=False, k="text", al=PP_ALIGN.LEFT, it=False):
            tb = sl.shapes.add_textbox(x,y,w,h)
            tf = tb.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.alignment = al
            r = p.add_run(); r.text = str(text)
            r.font.size=Pt(sz); r.font.bold=bold; r.font.italic=it
            r.font.name=FONT; r.font.color.rgb=C(k)
        def BODY(sl, text, x,y,w,h):
            if text: T(sl,text,x,y,w,h,13,k="subtext")
        def BULLETS(sl, items, x,y,w,h, sz=17):
            tb = sl.shapes.add_textbox(x,y,w,h)
            tf = tb.text_frame; tf.word_wrap=True
            for i,b in enumerate(items):
                p2 = tf.paragraphs[0] if i==0 else tf.add_paragraph()
                p2.space_before=Pt(7); p2.space_after=Pt(3)
                r2=p2.add_run(); r2.text=f"  ▸  {b}"
                r2.font.size=Pt(sz); r2.font.name=FONT; r2.font.color.rgb=C("text")
        def HDR(sl, title, body="", hh=Inches(1.5)):
            R(sl,0,0,W_,hh,"bg2")
            R(sl,0,0,Inches(0.12),H_,"accent")
            R(sl,Inches(0.12),hh-Inches(0.05),W_,Inches(0.05),"accent2")
            T(sl,title,Inches(0.3),Inches(0.15),Inches(12.5),Inches(0.75),26,True)
            BODY(sl,body,Inches(0.3),Inches(0.9),Inches(12.5),Inches(0.5))
        def PN(sl,n):
            tb=sl.shapes.add_textbox(Inches(12.6),Inches(7.1),Inches(0.6),Inches(0.3))
            p=tb.text_frame.paragraphs[0]; p.alignment=PP_ALIGN.RIGHT
            r=p.add_run(); r.text=str(n); r.font.size=Pt(10)
            r.font.name=FONT; r.font.color.rgb=RGBColor(0x55,0x55,0x55)

        for idx,si in enumerate(slides_data):
            sl   = prs.slides.add_slide(BL)
            lay  = si.get("layout","content")
            ttl  = si.get("title","")
            sub  = si.get("subtitle","")
            body = si.get("body","")
            buls = si.get("bullets",[])
            hdrs = si.get("headers",[])
            rows = si.get("rows",[])
            sn   = idx+1

            R(sl,0,0,W_,H_,"bg")  # 공통 배경

            if lay == "title":
                R(sl,int(W_*0.55),0,int(W_*0.45),H_,"bg2")
                R(sl,0,0,Inches(0.12),H_,"accent2")
                R(sl,Inches(0.12),int(H_*0.88),W_,Inches(0.04),"accent")
                org=(summary or {}).get("주최기관","")
                if org: T(sl,org,Inches(0.5),Inches(0.35),Inches(8),Inches(0.5),13,k="subtext")
                T(sl,ttl,Inches(0.5),Inches(1.6),Inches(7.5),Inches(3.0),40,True)
                R(sl,Inches(0.5),Inches(4.7),Inches(4),Inches(0.05),"accent")
                if sub: T(sl,sub,Inches(0.5),Inches(4.85),Inches(7.5),Inches(1.0),15,k="subtext")

            elif lay == "closing":
                R(sl,int(W_*0.55),0,int(W_*0.45),H_,"bg2")
                R(sl,0,0,Inches(0.12),H_,"accent2")
                R(sl,Inches(0.12),int(H_*0.88),W_,Inches(0.04),"accent")
                T(sl,ttl,Inches(0.5),Inches(2.0),Inches(7.5),Inches(2.0),46,True)
                R(sl,Inches(0.5),Inches(4.2),Inches(3.5),Inches(0.05),"accent")
                if sub: T(sl,sub,Inches(0.5),Inches(4.4),Inches(7.5),Inches(0.7),18,k="subtext")
                dept=" ".join(filter(None,[(summary or {}).get("담당부서",""),
                                           (summary or {}).get("담당자","")])).strip()
                if dept: T(sl,dept,Inches(0.5),Inches(6.9),Inches(8),Inches(0.4),11,k="subtext",it=True)

            elif lay == "section":
                R(sl,int(W_*0.6),0,int(W_*0.4),H_,"bg2")
                R(sl,0,0,Inches(0.25),H_,"accent")
                R(sl,Inches(0.5),int(H_*0.62),Inches(7.5),Inches(0.04),"accent2")
                T(sl,f"{sn-1:02d}",Inches(0.4),Inches(0.8),Inches(7.5),Inches(2.5),110,True,k="accent")
                T(sl,ttl,Inches(0.5),Inches(3.8),Inches(8.5),Inches(1.5),32,True)

            elif lay == "highlight":
                stat_n = si.get("stat_number","")
                stat_l = si.get("stat_label","")
                HDR(sl,ttl,body)
                R(sl,Inches(0.3),Inches(1.65),Inches(4.5),Inches(5.2),"card")
                R(sl,Inches(0.3),Inches(1.65),Inches(0.1),Inches(5.2),"accent2")
                T(sl,stat_n,Inches(0.5),Inches(2.3),Inches(4.1),Inches(2.5),72,True,k="accent2",al=PP_ALIGN.CENTER)
                T(sl,stat_l,Inches(0.5),Inches(5.0),Inches(4.1),Inches(0.6),16,k="subtext",al=PP_ALIGN.CENTER)
                if buls: BULLETS(sl,buls,Inches(5.1),Inches(1.7),Inches(7.8),Inches(5.2),16)
                PN(sl,sn)

            elif lay == "two_column":
                lt=si.get("left_title","현황");  lb=si.get("left_bullets",[])
                rt=si.get("right_title","목표"); rb=si.get("right_bullets",[])
                HDR(sl,ttl,body)
                cw=Inches(5.9); cy=Inches(1.65); ch=Inches(5.2); lx=Inches(0.3); rx=lx+cw+Inches(0.3)
                for x,title,bullets,ak in [(lx,lt,lb,"accent"),(rx,rt,rb,"accent2")]:
                    R(sl,x,cy,cw,ch,"card")
                    R(sl,x,cy,Inches(0.1),ch,ak)
                    T(sl,title,x+Inches(0.2),cy+Inches(0.15),cw-Inches(0.3),Inches(0.55),18,True,k=ak)
                    if bullets: BULLETS(sl,bullets,x+Inches(0.1),cy+Inches(0.8),cw-Inches(0.2),Inches(4.2),15)
                PN(sl,sn)

            elif lay=="table" and hdrs and rows:
                HDR(sl,ttl,body,Inches(1.4))
                cc=len(hdrs); rc=len(rows)+1
                tbl=sl.shapes.add_table(rc,cc,Inches(0.3),Inches(1.5),Inches(12.7),Inches(5.7)).table
                cw_=Inches(12.7)//cc
                for c in range(cc): tbl.columns[c].width=cw_
                for c,h in enumerate(hdrs):
                    cell=tbl.cell(0,c); cell.text=str(h)
                    cell.fill.solid(); cell.fill.fore_color.rgb=C("accent")
                    p2=cell.text_frame.paragraphs[0]; p2.alignment=PP_ALIGN.CENTER
                    r2=p2.runs[0] if p2.runs else p2.add_run()
                    r2.font.bold=True; r2.font.size=Pt(14); r2.font.color.rgb=W(); r2.font.name=FONT
                for ri,row in enumerate(rows):
                    rbg=pal["card"] if ri%2==0 else pal["bg2"]
                    for c,v in enumerate(row[:cc]):
                        cell=tbl.cell(ri+1,c); cell.text=str(v)
                        cell.fill.solid(); cell.fill.fore_color.rgb=RGBColor(*rbg)
                        p2=cell.text_frame.paragraphs[0]; p2.alignment=PP_ALIGN.CENTER
                        r2=p2.runs[0] if p2.runs else p2.add_run()
                        r2.font.size=Pt(13); r2.font.color.rgb=C("text"); r2.font.name=FONT
                PN(sl,sn)

            else:  # content
                is_toc=any(str(b).startswith(("01","02","1.","2.")) for b in buls)
                hh=Inches(1.5) if body else Inches(1.3)
                HDR(sl,ttl,body,hh)
                cy=hh+Inches(0.1); ch=H_-hh-Inches(0.4)
                if is_toc:
                    for i,b in enumerate(buls[:6]):
                        cx=Inches(0.5)+(i%2)*Inches(6.7); cy2=cy+(i//2)*Inches(1.55)
                        R(sl,cx,cy2,Inches(5.8),Inches(1.35),"card")
                        R(sl,cx,cy2,Inches(0.08),Inches(1.35),"accent")
                        T(sl,str(b),cx+Inches(0.2),cy2+Inches(0.35),Inches(5.5),Inches(0.7),18,True)
                elif buls:
                    BULLETS(sl,buls,Inches(0.4),cy,Inches(12.5),ch,17)
                PN(sl,sn)

        buf=io.BytesIO(); prs.save(buf); return buf.getvalue()
    except ImportError:
        st.error("pip install python-pptx"); return None
    except Exception as e:
        st.error(f"PPT 오류: {e}")
        import traceback; st.text(traceback.format_exc()); return None

# ── 6. 명찰 ──────────────────────────────────────────────────────────
def render_namecard():
    st.subheader("🪪 명찰 자동생성")
    summary = st.session_state.get("plan_summary_dict", {})
    event_name = summary.get("행사명","행사") if summary else "행사"
    st.info("엑셀 파일에 **이름**, **소속** 컬럼이 있어야 합니다.")
    c1,c2=st.columns(2)
    excel_file=c1.file_uploader("명단 엑셀 업로드",type=["xlsx","xls"],key="nc_excel")
    schedule=c2.text_area("뒷면 일정표",placeholder="09:00 등록\n10:00 개회식",height=150)
    ev=st.text_input("행사명 (비워두면 자동)",placeholder=event_name) or event_name

    if excel_file and st.button("🪪 명찰 생성"):
        try:
            from openpyxl import load_workbook
            wb=load_workbook(io.BytesIO(excel_file.read()),data_only=True); ws=wb.active
            hdrs=[str(c.value).strip() if c.value else "" for c in ws[1]]
            ni=next((i for i,h in enumerate(hdrs) if any(k in h for k in ["이름","성명","name","Name"])),None)
            oi=next((i for i,h in enumerate(hdrs) if any(k in h for k in ["소속","기관","org","Org","dept"])),None)
            if ni is None: st.error(f"이름 컬럼 없음. 현재: {hdrs}"); return
            persons=[{"name":str(r[ni]).strip(),"org":str(r[oi]).strip() if oi and r[oi] else ""}
                     for r in ws.iter_rows(min_row=2,values_only=True)
                     if r[ni] and str(r[ni]).lower()!="none"]
            if not persons: st.error("유효한 데이터 없음"); return
            st.info(f"총 {len(persons)}명 처리 중...")
            data=_build_namecard(persons,ev,schedule)
            if data:
                st.session_state["nc_bytes"]=data
                st.session_state["nc_count"]=len(persons)
        except Exception as e:
            st.error(f"오류: {e}")
            import traceback; st.text(traceback.format_exc())

    if st.session_state.get("nc_bytes"):
        st.success(f"✅ {st.session_state['nc_count']}명 생성 완료!")
        st.download_button(f"⬇️ 명찰 다운로드",data=st.session_state["nc_bytes"],
                           file_name="명찰.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

def _build_namecard(persons, event_name, schedule_text):
    try:
        from docx import Document
        from docx.shared import Pt,Cm,RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        doc=Document()
        for sec in doc.sections:
            sec.page_width=Cm(21); sec.page_height=Cm(29.7)
            sec.top_margin=sec.bottom_margin=Cm(1.5)
            sec.left_margin=sec.right_margin=Cm(2.0)

        def dashed(d):
            hr=d.add_paragraph(); hr.paragraph_format.space_before=hr.paragraph_format.space_after=Pt(0)
            pPr=hr._p.get_or_add_pPr(); pBdr=OxmlElement("w:pBdr")
            b=OxmlElement("w:bottom")
            b.set(qn("w:val"),"dashed"); b.set(qn("w:sz"),"6")
            b.set(qn("w:space"),"1"); b.set(qn("w:color"),"888888")
            pBdr.append(b); pPr.append(pBdr)

        def para(d,text,size,bold=False,color=(0x1A,0x1A,0x2E),before=4,after=4,align=WD_ALIGN_PARAGRAPH.CENTER):
            p=d.add_paragraph(); p.alignment=align
            p.paragraph_format.space_before=Pt(before); p.paragraph_format.space_after=Pt(after)
            r=p.add_run(text); r.font.name=FONT; r.font.size=Pt(size)
            r.font.bold=bold; r.font.color.rgb=RGBColor(*color)

        def front(name,org,first):
            if not first: dashed(doc)
            para(doc,event_name,15,True,(0x1F,0x38,0x64),28,8)
            if org: para(doc,org,17,False,(0x2E,0x74,0xB5),4,4)
            para(doc,name,38,True,(0x1A,0x1A,0x2E),8,28)

        def back(first):
            if not first: dashed(doc)
            para(doc,"행사 일정",15,True,(0x1F,0x38,0x64),28,10)
            for line in (schedule_text or "일정 미정").strip().split("\n"):
                para(doc,line.strip(),13,False,(0x1A,0x1A,0x2E),3,3)
            para(doc,"",13,False,(0x1A,0x1A,0x2E),4,28)

        for i in range(0,len(persons),2):
            batch=persons[i:i+2]
            front(batch[0]["name"],batch[0]["org"],True)
            if len(batch)>1: front(batch[1]["name"],batch[1]["org"],False)
            doc.add_page_break()
            back(True)
            if len(batch)>1: back(False)
            if i+2<len(persons): doc.add_page_break()

        buf=io.BytesIO(); doc.save(buf); return buf.getvalue()
    except ImportError:
        st.error("pip install python-docx"); return None
    except Exception as e:
        st.error(f"명찰 오류: {e}")
        import traceback; st.text(traceback.format_exc()); return None

# ── 7. 행사장 약도 ────────────────────────────────────────────────────
def render_map():
    st.subheader("🗺️ 행사장 약도")
    summary=st.session_state.get("plan_summary_dict",{})
    default=summary.get("장소","") if summary else ""
    place=st.text_input("장소명 또는 주소",value=default,
                        placeholder="예: 부산 벡스코  또는  부산광역시 해운대구 APEC로 55")
    c1,c2,c3=st.columns(3)
    zoom=c1.slider("확대 수준",1,14,4,help="숫자가 클수록 확대")
    mw=c2.number_input("가로(px)",200,1200,800)
    mh=c3.number_input("세로(px)",200,900,600)

    if st.button("🗺️ 약도 생성") and place:
        try:
            kkey=st.secrets["KAKAO_API_KEY"]
            hdrs={"Authorization":f"KakaoAK {kkey}"}

            def kw(q):
                r=requests.get("https://dapi.kakao.com/v2/local/search/keyword.json",
                               headers=hdrs,params={"query":q})
                d=r.json().get("documents",[])
                return d[0] if d else None

            def addr(q):
                r=requests.get("https://dapi.kakao.com/v2/local/search/address.json",
                               headers=hdrs,params={"query":q})
                d=r.json().get("documents",[])
                return d[0] if d else None

            # 괄호 안 주소 추출
            pm=re.search(r"\(([^)]+)\)",place)
            paddr=pm.group(1).strip() if pm else None
            clean=re.sub(r"\s*\([^)]*\)","",place).strip()

            doc0=None
            if paddr:
                st.info(f"🔍 주소 감지: **{paddr}**")
                a=addr(paddr)
                if a:
                    an=((a.get("road_address") or {}).get("address_name")
                        or (a.get("address") or {}).get("address_name",paddr))
                    doc0={"x":a["x"],"y":a["y"],"place_name":clean,"road_address_name":an}

            if not doc0: doc0=kw(clean)
            if not doc0: doc0=kw(place)
            if not doc0:
                with st.spinner("AI가 주소 보정 중..."):
                    fixed=ai(f"다음 장소명의 정확한 도로명 주소를 한 줄로만 출력하세요.\n장소명: {clean}")
                st.info(f"🔍 AI 보정: **{fixed}**")
                doc0=kw(fixed)
                if not doc0:
                    a=addr(fixed)
                    if a:
                        an=((a.get("road_address") or {}).get("address_name")
                            or (a.get("address") or {}).get("address_name",fixed))
                        doc0={"x":a["x"],"y":a["y"],"place_name":clean,"road_address_name":an}

            if not doc0 or not doc0.get("x"):
                st.error("위치를 찾을 수 없습니다. 도로명 주소를 직접 입력해보세요.")
                return

            lng,lat=float(doc0["x"]),float(doc0["y"])
            fn=doc0.get("place_name",clean)
            fa=doc0.get("road_address_name") or doc0.get("address_name","")
            st.success(f"📍 찾은 장소: **{fn}** ({fa})")

            # Static Map
            mr=requests.get("https://dapi.kakao.com/v2/maps/staticmap",headers=hdrs,
                            params={"center":f"{lng},{lat}","level":zoom,
                                    "w":int(mw),"h":int(mh),"markers":f"color:red|{lng},{lat}"})
            ct=mr.headers.get("Content-Type","")
            if mr.status_code==200 and "image" in ct:
                st.session_state.update({"map_img":mr.content,"map_name":fn})
            else:
                # 대체: OpenStreetMap
                ourl=(f"https://staticmap.openstreetmap.de/staticmap.php"
                      f"?center={lat},{lng}&zoom={zoom+2}&size={int(mw)}x{int(mh)}"
                      f"&markers={lat},{lng},red-pushpin")
                or_=requests.get(ourl,timeout=10)
                if or_.status_code==200 and "image" in or_.headers.get("Content-Type",""):
                    st.session_state.update({"map_img":or_.content,"map_name":fn})
                    st.caption("※ OpenStreetMap 기반")
                else:
                    from urllib.parse import quote
                    st.warning("지도 이미지 생성 실패.")
                    st.markdown(f"[🗺️ 카카오맵에서 보기](https://map.kakao.com/?q={quote(fa or clean)})")
        except KeyError:
            st.error("KAKAO_API_KEY가 secrets에 없습니다.")
        except Exception as e:
            st.error(f"오류: {e}")
            import traceback; st.text(traceback.format_exc())

    if st.session_state.get("map_img"):
        fn=st.session_state.get("map_name","행사장")
        st.image(st.session_state["map_img"],caption=f"📍 {fn}",use_container_width=True)
        st.download_button("⬇️ 약도 이미지 다운로드",data=st.session_state["map_img"],
                           file_name=f"행사장약도_{fn}.png",mime="image/png")

# ── 메인 ──────────────────────────────────────────────────────────────
def render_tab7():
    st.title("🤖 AI 문서 자동생성")
    st.caption("HWPX 계획서를 업로드하면 각종 문서를 자동으로 생성합니다.")
    st.markdown("---")
    st.subheader("📁 행사 계획서 업로드")

    hwpx=st.file_uploader("HWPX 파일 업로드 (.hwpx / .hwp)",type=["hwpx","hwp"])
    if hwpx:
        fkey=f"{hwpx.name}_{hwpx.size}"
        if (st.session_state.get("hwpx_fkey")!=fkey or "plan_summary_raw" not in st.session_state):
            st.session_state["hwpx_fkey"]=fkey
            with st.spinner("텍스트 추출 및 요약 중..."):
                raw=hwpx_to_text(hwpx)
                if len(raw)<100:
                    st.error("텍스트를 충분히 추출하지 못했습니다."); return
                summary_raw=summarize_hwpx(raw)
                st.session_state["plan_summary_raw"]=summary_raw
                try:
                    st.session_state["plan_summary_dict"]=json.loads(
                        re.sub(r"```json|```","",summary_raw).strip())
                except Exception:
                    st.session_state["plan_summary_dict"]={}

        summary=st.session_state.get("plan_summary_dict",{})
        with st.expander("📋 계획서 요약 (모든 기능에 재사용됨)",expanded=True):
            if summary:
                c1,c2=st.columns(2); keys=list(summary.keys()); half=len(keys)//2
                for i,k in enumerate(keys):
                    v=summary[k]
                    if isinstance(v,list): v=", ".join(str(x) for x in v)
                    (c1 if i<half else c2).markdown(f"**{k}:** {v}")
            else:
                st.text(st.session_state.get("plan_summary_raw",""))
    else:
        st.info("HWPX 업로드 없이도 수동으로 사용 가능합니다.")
        if "plan_summary_dict" not in st.session_state:
            st.session_state["plan_summary_dict"]={}

    st.markdown("---")
    tabs=st.tabs(["📄 문서 4종","🎤 사회자 멘트","🪧 현수막","📋 결과보고서","📊 PPT","🪪 명찰","🗺️ 행사장 약도"])
    with tabs[0]: render_doc4()
    with tabs[1]: render_mc()
    with tabs[2]: render_banner()
    with tabs[3]: render_result()
    with tabs[4]: render_ppt()
    with tabs[5]: render_namecard()
    with tabs[6]: render_map()

render = render_tab7  # main.py 호환

if __name__ == "__main__":
    st.set_page_config(page_title="AI 문서 자동생성", page_icon="🤖", layout="wide")
    render_tab7()
