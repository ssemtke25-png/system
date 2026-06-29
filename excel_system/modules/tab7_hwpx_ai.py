import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import google.generativeai as genai

# HWPX 내부에서 순수 텍스트만 쏙 뽑아내는 함수
def extract_text_from_hwpx(uploaded_file):
    text_list = []
    try:
        with zipfile.ZipFile(uploaded_file, 'r') as zf:
            for item in zf.namelist():
                if item.startswith('Contents/section') and item.endswith('.xml'):
                    xml_data = zf.read(item)
                    root = ET.fromstring(xml_data)
                    for elem in root.iter():
                        if elem.tag.endswith('}t') and elem.text:
                            text_list.append(elem.text)
        return "\n".join(text_list)
    except Exception as e:
        st.error(f"HWPX 파일을 읽는 중 오류가 발생했습니다: {e}")
        return None

# 탭 7 메인 화면 렌더링 함수
def render():
    st.header("🤖 AI 행사 문서 자동 작성기")
    st.info("행사 계획서(HWPX)를 올리면, AI가 내용 분석 후 요약본, 인사말(국장/과장), 보도자료를 작성해 줍니다.")

    # 🌟 1. API 키 숨기기 (스트림릿 금고에서 몰래 꺼내오기)
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except KeyError:
        st.error("⚠️ 관리자 설정 오류: 스트림릿 Secrets에 API 키가 설정되지 않았습니다.")
        return

    # API 입력창 삭제됨! 바로 파일 업로드로 넘어갑니다.
    uploaded_file = st.file_uploader("행사 계획서 파일 업로드 (.hwpx)", type=["hwpx"])

    if uploaded_file is not None:
        if st.button("✨ AI 문서 생성 시작"):
            with st.spinner("HWPX 파일에서 텍스트를 추출하는 중..."):
                doc_text = extract_text_from_hwpx(uploaded_file)
            
            if doc_text:
                st.success("텍스트 추출 성공! AI가 문서를 작성 중입니다...")
                
                try:
                    # AI 세팅
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel('gemini-2.5-flash') 
                    
                    # 🌟 2. 마법의 프롬프트 (요약/국장/과장/보도자료 분리)
                    prompt = f"""
                    다음은 우리 기관의 행사 계획서 내용입니다. 
                    이 내용을 꼼꼼히 분석하여 아래의 4가지 항목을 명확하게 구분해서 작성해 주세요.
                    
                    [행사 계획서 내용 시작]
                    {doc_text}
                    [행사 계획서 내용 끝]

                    ---
                    **1. 행사 핵심 요약본 (1장 짜리)**
                    - 바쁜 상사 보고를 위해 행사 개요(일시, 장소, 목적, 주요내용, 기대효과 등)를 개조식(글머리 기호 사용)으로 깔끔하게 1장 분량으로 요약해 주세요.

                    **2. 인사말씀 (국장님용)**
                    - 행사 전체를 총괄하는 국장님 격의 인사말입니다.
                    - 거시적인 비전, 기관의 정책 방향과의 연계성, 참석 내빈에 대한 정중한 감사 인사가 돋보이게 작성해 주세요. (A4 1장 내외 분량)

                    **3. 인사말씀 (과장님용)**
                    - 실무를 총괄하는 과장님 격의 인사말입니다.
                    - 행사 준비 노고 치하, 실무적인 기대효과, 참가자들에 대한 친근하고 따뜻한 격려 중심으로 작성해 주세요. (A4 반장 분량)

                    **4. 공식 보도자료 (Press Release)**
                    - 언론사에 배포할 공식 보도자료입니다.
                    - 매력적인 [제목], 육하원칙에 맞춘 [본문], 행사의 [기대효과]가 명확히 드러나게, 객관적이고 신뢰감 있는 기사체로 작성해 주세요.
                    """

                    with st.spinner("AI가 요약본, 국장님/과장님 인사말, 보도자료를 동시에 쓰고 있습니다... (약 10~20초 소요)"):
                        response = model.generate_content(prompt)
                    
                    # 결과 출력
                    st.divider()
                    st.subheader("🎉 AI 작성 결과")
                    st.markdown(response.text)
                    
                except Exception as e:
                    st.error(f"AI 생성 중 오류가 발생했습니다. 상세 오류: {e}")
            else:
                st.warning("파일에서 읽어올 텍스트가 없습니다. 파일 내용을 확인해 주세요.")
