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
                from .parser import extract_text_from_hwpx
                doc_text = extract_text_from_hwpx(uploaded_file)
            
            if doc_text:
                st.success("텍스트 추출 성공! AI가 문서를 작성 중입니다...")
                
                try:
                    # AI 세팅
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel('gemini-2.5-flash')
                    
                    # 🌟 2. 마법의 프롬프트 (요약/국장/과장/보도자료 분리)
                    # 🌟 [고도화된 프롬프트] 역할 부여, 감성 제거, 요약 강화
                    prompt = f"""
                    너는 20년 차 베테랑 공무원이다. 공공기관의 행사 계획서를 분석하여, 
                    아래의 4가지 보고서 형식으로 작성하되 '감성적인 수식어'는 모두 배제하고 
                    '신뢰감 있고 건조한 문체(행정 공문서 체계)'로 작성해라.

                    [행사 계획서 내용 시작]
                    {doc_text}
                    [행사 계획서 내용 끝]

                    ---
                    **1. 행사 핵심 요약본 (1장 개조식)**
                    - 바쁜 상사 보고를 위해 행사 개요(개요, 일시, 장소, 목적, 주요내용, 기대효과, 향후계획 등)를 개조식(글머리 기호 사용)으로 깔끔하게 1장 이내로 요약, 작성하라

                    **2. 인사말씀 (국장님용 - 비전/정책 중심)**
                    - 20년 차 공무원의 시각에서 작성해라.
                    - 행사와 우리 기관 정책 방향의 연계성을 강조하고, 내빈에게는 정중하되 건조하고 신뢰감 있는 문체로 작성해라.
                    - 감성적인 문구는 제거하고 행정적인 책임감과 비전 위주로 작성해라.
                    - 서두에 인삿말이 있어야 한다. 감성적인 어투도 좋다 (A4 1장 이내)

                    **3. 인사말씀 (과장님용 - 실무/격려 중심)**
                    - 실무 총괄 과장님 버전이다.
                    - 행사 준비 상황에 대한 언급, 참가자들에 대한 실무적인 당부, 행사의 원활한 운영을 위한 협조 요청 위주로 작성해라. 
                    - 서두에 인삿말이 있어야 한다. 감성적인 어투도 좋다 (A4 1장 이내)

                    **4. 공식 보도자료 (언론 배포용)**
                    - 형식: 아래 보도자료 구조를 반드시 엄수할 것.
                     1) [우측 상단]:【작성일자】, [담당부서], [작성자(과장/사무관/주무관)], [연락처]를 명시할 것.
                     2) [제목]: 본문의 핵심 성과를 포함하여 한 문장으로 작성할 것 (예: "경북도, 세계 마약퇴치의 날 맞아 청소년 마약 예방 캠페인 개최").
                     3) [부제목]: 대시(-)를 활용하여 성과 및 기대효과를 1~2개 핵심 포인트로 요약할 것.
                     4) [본문]: 첫 문단에서 육하원칙에 따라 행사를 명확히 기술할 것. 이후 문단에서는 행사의 목적, 추진 배경, 구체적인 기대효과 순으로 논리적으로 서술할 것.
                     5) [끝맺음]: 기관장 또는 부지사의 인용구(따옴표 활용)를 넣어 행사의 의미를 강조할 것.
                   - 문체: '~했다', '~한다'로 끝나는 명료한 개조식 기사체로 작성하고, 불필요한 감상적 표현이나 형용사는 모두 배제할 것.
                   - 참고: 제공된 첨부파일들의 행정 보도자료 스타일을 그대로 복제할 것.
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
