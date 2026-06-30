import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import google.generativeai as genai
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 🚨 라이브러리 인식 오류 철통 방어
try:
    from pptx import Presentation
except ImportError:
    st.error("🚨 `python-pptx` 라이브러리를 찾을 수 없습니다! 터미널에서 `pip install python-pptx`를 실행해 주세요.")
    Presentation = None

# HWPX 내부에서 순수 텍스트만 뽑아내는 함수 (최대 글자 수 제한)
def extract_text_from_hwpx(uploaded_file, max_chars=4000):
    text_list = []
    total_chars = 0
    try:
        uploaded_file.seek(0)
        with zipfile.ZipFile(uploaded_file, 'r') as zf:
            for item in zf.namelist():
                if item.lower().startswith('contents/section') and item.lower().endswith('.xml'):
                    xml_data = zf.read(item)
                    root = ET.fromstring(xml_data)
                    for elem in root.iter():
                        if elem.tag.endswith('}t') and elem.text:
                            text_list.append(elem.text)
                            total_chars += len(elem.text)
                            if total_chars >= max_chars:
                                return "\n".join(text_list)
        return "\n".join(text_list)
    except Exception as e:
        st.error(f"HWPX 파일을 읽는 중 오류가 발생했습니다: {e}")
        return None

# ✅ AI 응답 텍스트를 .docx 파일로 변환하는 함수
def create_docx_file(title, content_text):
    doc = Document()

    # 제목 스타일
    title_para = doc.add_heading(title, level=1)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()  # 빈 줄

    # 본문: 줄 단위로 추가 (## 제목 처리 포함)
    for line in content_text.split('\n'):
        line = line.strip()
        if not line:
            doc.add_paragraph()
            continue
        if line.startswith('## '):
            doc.add_heading(line.replace('## ', ''), level=2)
        elif line.startswith('# '):
            doc.add_heading(line.replace('# ', ''), level=1)
        elif line.startswith('- ') or line.startswith('* '):
            para = doc.add_paragraph(line[2:], style='List Bullet')
        else:
            para = doc.add_paragraph(line)
            para.runs[0].font.size = Pt(11) if para.runs else None

    docx_stream = io.BytesIO()
    doc.save(docx_stream)
    docx_stream.seek(0)
    return docx_stream

# PPT 파일 생성 함수
def create_ppt_file(parsed_slides):
    if Presentation is None:
        return None

    prs = Presentation()

    if not parsed_slides:
        parsed_slides = [{"title": "요약 생성 실패", "content": "AI가 내용 구조화에 실패했습니다."}]

    for slide_data in parsed_slides:
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)

        title_shape = slide.shapes.title
        if title_shape:
            title_shape.text = slide_data.get("title", "제목 없음")

        body_shape = None
        try:
            body_shape = slide.placeholders[1]
        except KeyError:
            for shape in slide.placeholders:
                if shape != title_shape:
                    body_shape = shape
                    break

        if body_shape:
            body_shape.text = slide_data.get("content", "")

    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# PPT 응답 텍스트 파싱 함수
def parse_ppt_response(raw_text):
    raw_text = raw_text.replace("**[SLIDE]**", "[SLIDE]").replace("[slide]", "[SLIDE]")
    raw_text = raw_text.replace("**[TITLE]**", "[TITLE]").replace("[Title]", "[TITLE]").replace("[title]", "[TITLE]")
    raw_text = raw_text.replace("**[CONTENT]**", "[CONTENT]").replace("[Content]", "[CONTENT]").replace("[content]", "[CONTENT]")

    slide_blocks = raw_text.split("[SLIDE]")
    slides_data = []

    for block in slide_blocks:
        if not block.strip():
            continue

        title = "제목 없음"
        content_lines = []
        is_content = False

        for line in block.strip().split('\n'):
            if line.startswith("[TITLE]"):
                title = line.replace("[TITLE]", "").strip()
            elif line.startswith("[CONTENT]"):
                is_content = True
            elif is_content:
                content_lines.append(line.strip())

        if title != "제목 없음" or content_lines:
            slides_data.append({
                "title": title,
                "content": "\n".join(content_lines).strip()
            })

    return slides_data

# 탭 7 메인 화면 렌더링 함수
def render():
    st.header("🤖 AI 행사 문서 & 파워포인트(PPT) 자동 작성기")
    st.info("행사 계획서(HWPX)를 올리면, AI가 요약본/인사말/보도자료/사회자멘트를 작성하고 **실제 PPT 발표 자료까지 알아서 디자인하여 제공합니다.**")

    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key)
    except:
        st.error("보안 금고(secrets.toml)에 GEMINI_API_KEY가 없습니다. 먼저 API 키를 등록해 주세요!")
        return

    uploaded_file = st.file_uploader("행사 계획서 (HWPX) 업로드", type=["hwpx"], key="hwpx_ai_up")

    if uploaded_file and st.button("🚀 천하무적 자동 생성 (문서 4종 + 사회자멘트 + PPT 다운로드)"):
        if Presentation is None:
            st.error("🚨 PPT 생성 라이브러리가 없습니다. 터미널에서 `pip install python-pptx`를 먼저 실행해 주세요!")
            return

        # 1단계: HWPX 텍스트 추출 (최대 4000자)
        with st.spinner("📂 HWPX 파일에서 핵심 데이터를 뽑아내고 있습니다..."):
            hwpx_text = extract_text_from_hwpx(uploaded_file, max_chars=4000)

        if not hwpx_text:
            return

        model = genai.GenerativeModel("gemini-2.5-flash-lite")

        # 2단계: 요약본 생성 (이후 모든 프롬프트에서 재사용 → 토큰 절약)
        with st.spinner("📋 AI가 원문을 압축 요약하는 중입니다... (토큰 절약 모드)"):
            try:
                prompt_summary = f"""
다음 행사 계획서 원문에서 핵심 정보만 추출해줘.
반드시 아래 항목만 간결하게 정리해. 다른 말은 쓰지 마.

- 행사명:
- 일시:
- 장소:
- 주최/주관:
- 참석 대상 및 규모:
- 행사 목적:
- 주요 프로그램 (번호 목록으로):
- 예산 또는 특이사항 (있으면):

원문:
{hwpx_text}
"""
                summary = model.generate_content(prompt_summary).text
            except Exception as e:
                st.error(f"요약 중 오류 발생: {e}")
                return

        # =============================================
        # 3단계: 문서 4종 + 사회자 멘트 + PPT 동시 생성
        # =============================================
        prompt_doc = f"""
다음 행사 요약 정보를 바탕으로 4가지 공식 문서를 작성해줘.

1. 1페이지 요약보고서 (핵심 개요 중심)
2. 인사말씀 (국장님용 - 비전/정책 중심)
3. 인사말씀 (과장님용 - 실무/격려 중심, A4 1장 이내)
4. 공식 보도자료 (언론 배포용 - 우측상단 정보, 제목, 부제목 등 양식 준수)

행사 요약:
{summary}
"""

        prompt_mc = f"""
다음 행사 요약 정보를 바탕으로 행사 전체 사회자 진행 멘트를 작성해줘.

형식은 아래처럼 순서별로 작성해:
- 각 순서 앞에 [순서명] 표시
- 실제 사회자가 말하는 멘트를 자연스럽게 작성
- 시작("안녕하십니까")부터 마무리("감사합니다") 까지 전체 포함
- 중간 전환 멘트, 박수 유도, 다음 순서 안내 포함

행사 요약:
{summary}
"""

        prompt_ppt = f"""
다음 행사 요약 정보를 바탕으로 PPT 슬라이드 자료를 기획해줘.
슬라이드는 4~6장 내외로 만들고, 반드시 아래 형식만 사용해. 다른 말은 절대 쓰지 마.

[SLIDE]
[TITLE] 슬라이드 제목
[CONTENT]
- 내용1
- 내용2

행사 요약:
{summary}
"""

        # 문서 4종 + docx 다운로드
        st.subheader("📄 생성된 문서 4종")
        with st.spinner("📝 AI가 문서 4종을 작성 중입니다..."):
            try:
                doc_response = model.generate_content(prompt_doc)
                doc_text = doc_response.text
                st.success("✅ 문서 4종 완료!")

                with st.expander("📄 생성된 문서 확인하기", expanded=True):
                    try:
                        st.markdown(doc_text)
                    except ValueError:
                        st.error("🚨 AI가 문서를 생성하지 못했습니다.")

                # ✅ docx 다운로드 버튼
                docx_file = create_docx_file("AI 자동생성 행사 문서 4종", doc_text)
                st.download_button(
                    label="📥 문서 4종 DOCX 다운로드",
                    data=docx_file.getvalue(),
                    file_name="AI_행사문서_4종.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_doc_btn"
                )
            except Exception as e:
                st.error(f"문서 생성 중 오류 발생: {e}")

        st.divider()

        # 사회자 멘트 + docx 다운로드
        st.subheader("🎤 사회자 진행 멘트")
        with st.spinner("🎤 AI가 사회자 멘트를 작성 중입니다..."):
            try:
                mc_response = model.generate_content(prompt_mc)
                mc_text = mc_response.text
                st.success("✅ 사회자 멘트 완료!")

                with st.expander("🎤 사회자 멘트 확인하기", expanded=True):
                    try:
                        st.markdown(mc_text)
                    except ValueError:
                        st.error("🚨 AI가 사회자 멘트를 생성하지 못했습니다.")

                # ✅ docx 다운로드 버튼
                mc_docx_file = create_docx_file("AI 자동생성 사회자 진행 멘트", mc_text)
                st.download_button(
                    label="📥 사회자 멘트 DOCX 다운로드",
                    data=mc_docx_file.getvalue(),
                    file_name="AI_사회자멘트.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_mc_btn"
                )
            except Exception as e:
                st.error(f"사회자 멘트 생성 중 오류 발생: {e}")

        st.divider()

        # PPT 생성 + 다운로드
        st.subheader("📊 파워포인트(PPT) 자동 생성")
        with st.spinner("📊 AI가 PPT 슬라이드를 조립 중입니다..."):
            try:
                ppt_response = model.generate_content(prompt_ppt)

                try:
                    raw_text = ppt_response.text
                except ValueError:
                    raw_text = "[SLIDE]\n[TITLE] 오류 발생\n[CONTENT]\nAI가 내용을 생성하지 못했습니다."

                slides_data = parse_ppt_response(raw_text)
                ppt_file = create_ppt_file(slides_data)

                if ppt_file:
                    st.success("✅ PPT 파일 완료!")
                    st.download_button(
                        label="📥 자동 생성된 PPT 다운로드 (.pptx)",
                        data=ppt_file.getvalue(),
                        file_name="AI_자동생성_발표자료.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key="download_ppt_btn"
                    )

                with st.expander("💡 AI가 구성한 슬라이드 기획안 원본 보기"):
                    st.text(raw_text)

            except Exception as e:
                st.error(f"PPT 생성 중 오류 발생: {e}")
