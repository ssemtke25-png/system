import streamlit as st

def render():
    st.info("공공행정 AI Assistant 개발 중입니다.")

doc = extract_text_from_hwpx()

prompt = build_main_prompt()

result = generate_document()

st.write(result)
