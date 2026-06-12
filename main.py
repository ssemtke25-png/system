import streamlit as st
import pandas as pd
import os
import json
import re  # 이 줄이 맨 위에 꼭 있어야 합니다!
from bs4 import BeautifulSoup

# ==========================================
# [1. 웹 페이지 기본 설정]
# ==========================================
st.set_page_config(page_title="지적재조사 통합 업무지원 시스템", page_icon="🔍", layout="wide")

DATA_DIR = "data"
EXCEL_PATH = f"{DATA_DIR}/data.xlsx"
EVENT_PATH = f"{DATA_DIR}/calendar_events.json"
LAW_HTML_PATH = f"{DATA_DIR}/지적재조사에 관한 특별법(인용조문 3단비교).html"
REG_DIR = f"{DATA_DIR}/규정"

# ==========================================
# [2. 데이터 파싱 함수 (규정 분석 기능 강화)]
# ==========================================
@st.cache_data
def load_all_data():
    law_db = []
    if os.path.exists(LAW_HTML_PATH):
        try:
            with open(LAW_HTML_PATH, "r", encoding="utf-8") as f:
                soup = BeautifulSoup(f, "html.parser")
            for tr in soup.select("tr"):
                tds = tr.select("td")
                if len(tds) >= 3 and (jo_span := tds[0].select_one("span.bl")):
                    law_db.append({
                        "조문": jo_span.get_text(strip=True),
                        "법률": tds[0].get_text("\n", strip=True),
                        "시행령": tds[1].get_text("\n", strip=True),
                        "시행규칙": tds[2].get_text("\n", strip=True)
                    })
        except: pass

    def safe_load_sheet(sheet_name):
        if not os.path.exists(EXCEL_PATH): return pd.DataFrame(columns=["제목", "내용", "수정여부"])
        try:
            return pd.read_excel(EXCEL_PATH, sheet_name=sheet_name).fillna("").astype(str)
        except:
            return pd.DataFrame(columns=["제목", "내용", "수정여부"])

    df_qna = safe_load_sheet("질의회신")
    df_case = safe_load_sheet("판례검색")

    reg_db = {}
    if os.path.exists(REG_DIR):
        for f in os.listdir(REG_DIR):
            if f.endswith(".html"):
                file_path = os.path.join(REG_DIR, f)
                html_content = ""
                try:
                    with open(file_path, "r", encoding="utf-8") as file:
                        html_content = file.read()
                except UnicodeDecodeError:
                    try:
                        with open(file_path, "r", encoding="cp949") as file:
                            html_content = file.read()
                    except:
                        continue
                
                if html_content:
                    soup = BeautifulSoup(html_content, "html.parser")
                    reg_list = []
                    for tr in soup.select("tr"):
                        tds = tr.select("td")
                        if tds and (jo_span := tr.select_one("span.bl")):
                            reg_list.append({"조문": jo_span.get_text(strip=True), "내용": "\n".join([td.get_text("\n", strip=True) for td in tds])})
                    
                    # [핵심 수정] 표 형태가 아닐 때 글을 읽어내는 PC버전의 분석기 부활!
                    if not reg_list:
                        text_lines = soup.get_text(separator="\n").split("\n")
                        current_jo, current_content = "전체 내용", []
                        for line in text_lines:
                            line = line.strip()
                            if not line: continue
                            if re.match(r'^제\s*\d+\s*조', line) or re.match(r'^\[별표', line):
                                if current_content: 
                                    reg_list.append({"조문": current_jo, "내용": "\n".join(current_content)})
                                current_jo, current_content = line, [line]
                            else: 
                                current_content.append(line)
                        if current_content: 
                            reg_list.append({"조문": current_jo, "내용": "\n".join(current_content)})

                    if reg_list: 
                        reg_db[f.replace(".html", "")] = reg_list

    return df_qna, df_case, law_db, reg_db

df_qna, df_case, law_db, reg_db = load_all_data()
# 3. 카테고리별 데이터 출력 로직
if mode in ["📑 질의회신", "🧑‍⚖️ 판례"]:
    target_df = df_qna if mode == "📑 질의회신" else df_case
    if keyword:
        if only_title:
            res = target_df[target_df['제목'].str.contains(keyword, case=False, na=False)]
        else:
            res = target_df[target_df['제목'].str.contains(keyword, case=False, na=False) | 
                            target_df['내용'].str.contains(keyword, case=False, na=False)]
    else:
        res = target_df 
        
    st.caption(f"총 {len(res)}건의 자료가 있습니다.")
    
    for idx, row in res.iterrows(): 
        icon = "🟢" if str(row.get("수정여부")).strip().upper() == "Y" else "📑"
        with st.expander(f"{icon} {row['제목']}"):
            content = row['내용'].replace("\n", "<br>")
            content = highlight_text(content, keyword) if keyword else content
            st.markdown(content, unsafe_allow_html=True)
            
            # [추가된 복사 기능]
            st.info("💡 내용 공유하기 (아래 상자 우측 상단의 📋 아이콘 클릭)")
            st.code(f"[{row['제목']}]\n{row['내용']}", language="text")

elif mode == "⚖️ 법령":
    if keyword:
        count = 0
        for item in law_db:
            match = (keyword in item['조문']) if only_title else (keyword in item['조문'] or keyword in item['법률'] or keyword in item['시행령'] or keyword in item['시행규칙'])
            if match:
                count += 1
                with st.expander(f"⚖️ [법령] {item['조문']}"):
                    st.markdown(f"**📜 [법률]**<br>{highlight_text(item['법률'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
                    st.markdown("---")
                    st.markdown(f"**⚙️ [시행령]**<br>{highlight_text(item['시행령'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
                    st.markdown("---")
                    st.markdown(f"**📝 [시행규칙]**<br>{highlight_text(item['시행규칙'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
                    
                    # [추가된 복사 기능]
                    st.info("💡 조문 3단 비교 복사하기 (우측 상단 📋 클릭)")
                    copy_text = f"[{item['조문']}]\n\n[법률]\n{item['법률']}\n\n[시행령]\n{item['시행령']}\n\n[시행규칙]\n{item['시행규칙']}"
                    st.code(copy_text, language="text")
        st.caption(f"총 {count}건의 조문이 검색되었습니다.")
        
    else:
        st.caption(f"총 {len(law_db)}개의 전체 조문입니다.")
        for item in law_db:
            with st.expander(f"⚖️ [법령] {item['조문']}"):
                st.markdown(f"**📜 [법률]**<br>{item['법률'].replace(chr(10), '<br>')}", unsafe_allow_html=True)
                st.markdown("---")
                st.markdown(f"**⚙️ [시행령]**<br>{item['시행령'].replace(chr(10), '<br>')}", unsafe_allow_html=True)
                st.markdown("---")
                st.markdown(f"**📝 [시행규칙]**<br>{item['시행규칙'].replace(chr(10), '<br>')}", unsafe_allow_html=True)
                
                # [추가된 복사 기능]
                st.info("💡 조문 3단 비교 복사하기 (우측 상단 📋 클릭)")
                copy_text = f"[{item['조문']}]\n\n[법률]\n{item['법률']}\n\n[시행령]\n{item['시행령']}\n\n[시행규칙]\n{item['시행규칙']}"
                st.code(copy_text, language="text")

elif mode in ["🏢 업무규정", "📐 측량규정"]:
    is_survey = (mode == "📐 측량규정")
    
    count = 0
    for reg_name, reg_data in reg_db.items():
        if (is_survey and "측량" in reg_name) or (not is_survey and "측량" not in reg_name):
            display_name = reg_name.replace("규정_", "")
            
            for item in reg_data:
                match = True
                if keyword:
                    match = (keyword in item['조문']) if only_title else (keyword in item['조문'] or keyword in item['내용'])
                
                if match:
                    count += 1
                    with st.expander(f"📖 [{display_name}] {item['조문']}"):
                        content = item['내용'].replace("\n", "<br>")
                        content = highlight_text(content, keyword) if keyword else content
                        st.markdown(content, unsafe_allow_html=True)
                        
                        # [추가된 복사 기능]
                        st.info("💡 규정 내용 복사하기 (우측 상단 📋 클릭)")
                        st.code(f"[{display_name} {item['조문']}]\n{item['내용']}", language="text")
                        
    if keyword:
        st.caption(f"총 {count}건이 검색되었습니다.")
    else:
        st.caption(f"총 {count}건의 전체 목록입니다.")

# ==========================================
# [4 & 5. 모바일 최적화 화면 배치 및 카테고리 분리]
# ==========================================

st.markdown("<h4 style='text-align: center; color: #2c3e50; font-size: 1.3rem; margin-top: -40px; margin-bottom: 10px;'>🔍 지적재조사 통합 검색</h4>", unsafe_allow_html=True)
# 1. 검색창과 버튼 배치
col1, col2 = st.columns([3, 1])
with col1:
    keyword = st.text_input("검색어를 입력하세요", label_visibility="collapsed", placeholder="🔍 검색어 입력 (예: 경계설정)")
with col2:
    search_btn = st.button("검색", use_container_width=True)

# [핵심 추가] 제목만 검색 체크박스 (기본적으로 체크되게 설정)
only_title = st.checkbox("☑️ 제목만 검색", value=True)

# 2. 5개의 카테고리로 세분화된 아이콘 메뉴
tabs = ["📑 질의회신", "⚖️ 법령", "🏢 업무규정", "📐 측량규정", "🧑‍⚖️ 판례"]
mode = st.radio("자료 선택", tabs, horizontal=True, label_visibility="collapsed")

st.markdown("---")

def highlight_text(text, kw):
    if not kw: return text
    return text.replace(kw, f"<mark style='background-color: yellow;'>{kw}</mark>")

# 3. 카테고리별 데이터 출력 로직
if mode in ["📑 질의회신", "🧑‍⚖️ 판례"]:
    target_df = df_qna if mode == "📑 질의회신" else df_case
    if keyword:
        # 체크박스 상태에 따라 검색 범위 다르게 적용!
        if only_title:
            res = target_df[target_df['제목'].str.contains(keyword, case=False, na=False)]
        else:
            res = target_df[target_df['제목'].str.contains(keyword, case=False, na=False) | 
                            target_df['내용'].str.contains(keyword, case=False, na=False)]
    else:
        res = target_df 
        
    st.caption(f"총 {len(res)}건의 자료가 있습니다.")
    
    for idx, row in res.iterrows(): 
        icon = "🟢" if str(row.get("수정여부")).strip().upper() == "Y" else "📑"
        with st.expander(f"{icon} {row['제목']}"):
            content = row['내용'].replace("\n", "<br>")
            content = highlight_text(content, keyword) if keyword else content
            st.markdown(content, unsafe_allow_html=True)

elif mode == "⚖️ 법령":
    if keyword:
        count = 0
        for item in law_db:
            # 체크박스 상태에 따라 조문(제목)만 볼지, 법률/시행령 내용까지 볼지 결정
            match = (keyword in item['조문']) if only_title else (keyword in item['조문'] or keyword in item['법률'] or keyword in item['시행령'] or keyword in item['시행규칙'])
            if match:
                count += 1
                with st.expander(f"⚖️ [법령] {item['조문']}"):
                    st.markdown(f"**📜 [법률]**<br>{highlight_text(item['법률'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
                    st.markdown("---")
                    st.markdown(f"**⚙️ [시행령]**<br>{highlight_text(item['시행령'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
                    st.markdown("---")
                    st.markdown(f"**📝 [시행규칙]**<br>{highlight_text(item['시행규칙'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
        st.caption(f"총 {count}건의 조문이 검색되었습니다.")
        
    else:
        st.caption(f"총 {len(law_db)}개의 전체 조문입니다.")
        for item in law_db:
            with st.expander(f"⚖️ [법령] {item['조문']}"):
                st.markdown(f"**📜 [법률]**<br>{item['법률'].replace(chr(10), '<br>')}", unsafe_allow_html=True)
                st.markdown("---")
                st.markdown(f"**⚙️ [시행령]**<br>{item['시행령'].replace(chr(10), '<br>')}", unsafe_allow_html=True)
                st.markdown("---")
                st.markdown(f"**📝 [시행규칙]**<br>{item['시행규칙'].replace(chr(10), '<br>')}", unsafe_allow_html=True)

elif mode in ["🏢 업무규정", "📐 측량규정"]:
    is_survey = (mode == "📐 측량규정")
    
    count = 0
    for reg_name, reg_data in reg_db.items():
        if (is_survey and "측량" in reg_name) or (not is_survey and "측량" not in reg_name):
            display_name = reg_name.replace("규정_", "")
            
            for item in reg_data:
                match = True
                if keyword:
                    # 규정도 조문(제목)만 볼지, 규정 내용까지 볼지 결정
                    match = (keyword in item['조문']) if only_title else (keyword in item['조문'] or keyword in item['내용'])
                
                if match:
                    count += 1
                    with st.expander(f"📖 [{display_name}] {item['조문']}"):
                        content = item['내용'].replace("\n", "<br>")
                        content = highlight_text(content, keyword) if keyword else content
                        st.markdown(content, unsafe_allow_html=True)
                        
    if keyword:
        st.caption(f"총 {count}건이 검색되었습니다.")
    else:
        st.caption(f"총 {count}건의 전체 목록입니다.")
# 하단 푸터
st.markdown("---")
st.caption("v4.0 Web Version - 데이터 수정은 data.xlsx 파일을 변경 후 다시 배포해주세요.")
