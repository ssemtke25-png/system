import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import google.generativeai as genai
import io
from pptx import Presentation

# HWPX 내부에서 순수 텍스트만 쏙 뽑아내는 함수 (대소문자 무시 기능 탑재)
def extract_text_from_hwpx(uploaded_file):
    text_list = []
    try:
        with zipfile.ZipFile(uploaded_file, 'r') as zf:
            for item in zf.namelist():
                # Section 대소문자 문제 완벽 해결
                if item.lower().startswith('contents/section') and item.lower().endswith('.xml'):
                    xml_data = zf.read(item)
                    root = ET.fromstring(xml_data)
                    for elem in root.iter():
                        if elem.tag.endswith('}t') and elem.text:
                            text_list.append(elem.text)
        return "\n".join(text_list)
    except Exception as e:
        st.error(f"HWPX 파일을 읽는 중 오류가 발생했습니다: {e}")
        return None

# AI가 짜준 슬라이드 구조를 진짜 PPT 파일로 변환하는 천하무적 함수
def create_ppt_file(parsed_slides):
    prs = Presentation()
    for slide_data in parsed_slides:
        # 슬라이드 레이아웃 1번: 제목 + 내용 (가장 표준적인 레이아웃)
        slide_layout = prs.slide_layouts[1] 
        slide = prs.slides.add_slide(slide_layout)
        
        title_shape = slide.shapes.title
        body_shape = slide.placeholders[1]
        
        # 제목과 본문 넣기
        title_shape.text = slide_data.get("title", "제목 없음")
        body_shape.text = slide_data.get("content", "")
        
    # 완성된 PPT를 메모리에 저장
    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# 탭 7 메인 화면 렌더링 함수
def render():
    st.header("🤖 AI 행사 문서 & 파워포인트(PPT) 자동 작성기")
    st.info("행사 계획서(HWPX)를 올리면, AI가 요약본/인사말/보도자료를 작성하고 **실제 PPT 발표 자료까지 알아서 디자인하여 제공합니다.**")

    # 1. API 키 숨기기 (스트림릿 금고에서 꺼내오기)
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key)
    except:
        st.error("보안 금고(secrets.toml)에 GEMINI_API_KEY가 없습니다. 먼저 API 키를 등록해 주세요!")
        return

    uploaded_file = st.file_uploader("행사 계획서 (HWPX) 업로드", type=["hwpx"], key="hwpx_ai_up")
    
    if uploaded_file and st.button("🚀 천하무적 자동 생성 (문서 4종 + PPT 다운로드)"):
        with st.spinner("HWPX 파일에서 핵심 데이터를 뽑아내고 있습니다..."):
            hwpx_text = extract_text_from_hwpx(uploaded_file)
            
        if not hwpx_text:
            return
            
        # ==========================================
        # 프롬프트 1: 보고서, 인사말, 보도자료용
        # ==========================================
        prompt_doc = f"""
        다음 행사 계획서를 분석해서 4가지 공식 문서를 작성해줘.
        
        1. 1페이지 요약보고서 (핵심 개요 중심)
        2. 인사말씀 (국장님용 - 비전/정책 중심)
        3. 인사말씀 (과장님용 - 실무/격려 중심, A4 1장 이내)
        4. 공식 보도자료 (언론 배포용 - 우측상단 정보, 제목, 부제목 등 양식 준수)
        
        원문 데이터:
        {hwpx_text}
        """
        
        # ==========================================
        # 프롬프트 2: PPT 자동 생성을 위한 구조화된 텍스트
        # ==========================================
        prompt_ppt = f"""
        다음 행사 계획서 내용을 바탕으로 파워포인트(PPT) 슬라이드 자료를 기획해줘.
        슬라이드는 4~6장 내외로 만들고, 파이썬 코드가 인식할 수 있게 **반드시** 아래 형식을 지켜서 출력해.
        다른 인사말이나 군더더기 설명은 절대 쓰지 마.

        [SLIDE]
        [TITLE] 슬라이드 제목 (예: 행사 개요)
        [CONTENT]
        - 일시 및 장소: ...
        - 참석 대상: ...

        [SLIDE]
        [TITLE] 주요 프로그램
        [CONTENT]
        - 1부: ...
        - 2부: ...
        
        원문 데이터:
        {hwpx_text}
        """
        
        model = genai.GenerativeModel("gemini-1.5-flash")
        
        # 화면을 반으로 쪼개서 왼쪽은 일반 문서, 오른쪽은 PPT 생성기 배치
        col1, col2 = st.columns(2)
        
        with col1:
            with st.spinner("📝 AI가 행사 문서(요약, 인사말, 보도자료)를 치열하게 작성 중입니다..."):
                try:
                    doc_response = model.generate_content(prompt_doc)
                    st.success("✅ 문서 초안 작성 완료!")
                    with st.expander("📄 생성된 문서 확인하기", expanded=True):
                        st.markdown(doc_response.text)
                except Exception as e:
                    st.error(f"문서 생성 중 오류 발생: {e}")
                
        with col2:
            with st.spinner("📊 AI가 기획안을 분석하여 PPT 슬라이드를 조립 중입니다..."):
                try:
                    ppt_response = model.generate_content(prompt_ppt)
                    
                    # AI가 보내준 텍스트를 파싱(해석)해서 슬라이드 데이터로 분리
                    slides_data = []
                    raw_text = ppt_response.text
                    slide_blocks = raw_text.split("[SLIDE]")
                    
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
                                
                        if title or content_lines:
                            slides_data.append({
                                "title": title,
                                "content": "\n".join(content_lines).strip()
                            })
                    
                    # 🚀 대망의 PPT 파일 제작!
                    ppt_file = create_ppt_file(slides_data)
                    st.success("✅ PPT 파워포인트 파일 디자인 완료!")
                    
                    # 다운로드 버튼 생성
                    st.download_button(
                        label="📥 자동 생성된 PPT 다운로드 (.pptx)",
                        data=ppt_file,
                        file_name="AI_자동생성_발표자료.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                    
                    with st.expander("💡 AI가 구성한 슬라이드 기획안 미리보기"):
                        st.text(ppt_response.text)
                        
                except Exception as e:
                    st.error(f"PPT 생성 중 오류 발생: {e}")
