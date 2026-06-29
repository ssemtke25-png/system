"""
HWPX 파일에서 텍스트를 추출하는 모듈
"""

import zipfile
import xml.etree.ElementTree as ET
import streamlit as st


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
