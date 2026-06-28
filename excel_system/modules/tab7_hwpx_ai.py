import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import google.generativeai as genai

# 🌟 1. HWPX 내부에서 순수 텍스트만 쏙 뽑아내는 함수
def extract_text_from_hwpx(uploaded_file):
    text_list = []
    try:
        # HWPX는 사실 zip 파일과 같습니다. 압축을 풀어서 속을 들여다봅니다.
        with zipfile.ZipFile(uploaded_file, 'r') as zf:
            for item in zf.namelist():
                # 본문 내용이 담긴 XML 파일들만 찾습니다.
                if item.startswith('Contents/section') and item.endswith('.xml'):
                    xml_data = zf.read(item)
                    root = ET.fromstring(xml_data)
                    
                    # 모든 태그를 뒤져서 글자가 있는 부분(보통 <hp:t>)을 찾아 합칩니다.
                    for elem in root.iter():
                        if elem.tag.endswith('}t') and elem.text:
                            text_list.append(elem.text)
        return "\n".join(text_list)
    except Exception as e:
        st.error(f"HWPX 파일을 읽는 중 오류가 발생했습니다: {e}")
        return None

# 🌟 2. 탭 7 메인 화면 렌더링 함수
def render():
    st.header("🤖 AI 행사 인사말 & 보도자료 작성기")
    st.info("행사 계획서(HWPX)를 올리면, AI가 내용을 분석해 초안을 작성해 줍니다.")

    # 사이드바나 상단에 API 키 입력란 배치 (보안을 위해 password 타입으로)
    api_key = st.text_input("🔑 구글 Gemini API Key를 입력하세요", type="password")

    uploaded_file = st.file_uploader("행사 계획서 파일 업로드 (.hwpx)", type=["hwpx"])

    if uploaded_file is not None and api_key:
        if st.button("✨ AI 초안 생성 시작"):
            with st.spinner("HWPX 파일에서 텍스트를 추출하는 중..."):
                doc_text = extract_text_from_hwpx(uploaded_file)
            
            if doc_text:
                st.success("텍스트 추출 성공! AI가 문서를 작성 중입니다...")
                
                try:
                    # 🌟 3. 구글 Gemini AI 세팅 및 명령 전달
                    genai.configure(api_key=api_key)
                    # 최신 빠르고 똑똑한 모델 선택
                    model = genai.GenerativeModel('gemini-1.5-flash') 
                    
                    prompt = f"""
                    다음은 우리 기관의 행사 계획서 내용입니다. 
                    이 내용을 꼼꼼히 분석하여 다음 2가지 문서를 전문적인 공공기관 어조로 작성해 주세요.
                    
                    [행사 계획서 내용 시작]
                    {doc_text}
                    [행사 계획서 내용 끝]

                    ---
                    **요청 사항 1: 행사 인사말씀 (Opening Remarks)**
                    - 행사를 주관하는 대표자의 인사말을 작성해 주세요.
                    - 내빈에 대한 감사, 행사의 취지와 의미, 향후 비전 및 당부의 말씀이 포함되게 해주세요.
                    - 분량은 A4 1장 내외로, 너무 길지 않고 희망찬 어조로 작성해 주세요.

                    **요청 사항 2: 공식 보도자료 (Press Release)**
                    - 언론사에 배포할 공식 보도자료를 작성해 주세요.
                    - 매력적인 [제목], 육하원칙에 맞춘 [본문], 행사의 [기대효과]가 명확히 드러나게 해주세요.
                    - 객관적이고 신뢰감 있는 기사체로 작성해 주세요.
                    """

                    with st.spinner("AI가 인사말과 보도자료를 열심히 쓰고 있습니다... (약 10~20초 소요)"):
                        response = model.generate_content(prompt)
                    
                    # 🌟 4. 결과 출력
                    st.divider()
                    st.subheader("🎉 AI 작성 결과")
                    
                    # 텍스트 복사가 쉽게 st.text_area를 활용하거나 markdown 출력
                    st.markdown(response.text)
                    
                except Exception as e:
                    st.error(f"AI 생성 중 오류가 발생했습니다. API 키를 확인해 주세요. 상세 오류: {e}")
            else:
                st.warning("파일에서 읽어올 텍스트가 없습니다. 파일 내용을 확인해 주세요.")
    
    elif uploaded_file and not api_key:
        st.warning("먼저 구글 Gemini API Key를 입력해 주세요!")