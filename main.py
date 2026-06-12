import streamlit as st
import pandas as pd
import os
import json
import re
from bs4 import BeautifulSoup
from datetime import datetime, date

# ==========================================
# [1. 웹 페이지 기본 설정]
# ==========================================
st.write("")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
EXCEL_PATH = os.path.join(DATA_DIR, "data.xlsx")
EVENT_PATH = os.path.join(DATA_DIR, "calendar_events.json")
LAW_HTML_PATH = os.path.join(DATA_DIR, "지적재조사에 관한 특별법(인용조문 3단비교).html")
REG_DIR = os.path.join(DATA_DIR, "규정")

# ==========================================
# [2. 데이터 파싱 함수 (기존 로직 동일 유지)]
# ==========================================
@st.cache_data
def load_all_data():
    # 1. 법령 로드
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

    # 2. 엑셀 데이터 로드
    def safe_load_sheet(sheet_name):
        if not os.path.exists(EXCEL_PATH): return pd.DataFrame(columns=["제목", "내용", "수정여부"])
        try:
            return pd.read_excel(EXCEL_PATH, sheet_name=sheet_name).fillna("").astype(str)
        except:
            return pd.DataFrame(columns=["제목", "내용", "수정여부"])

    df_qna = safe_load_sheet("질의회신")
    df_case = safe_load_sheet("판례검색")

    # 3. 규정 파일 로드
    reg_db = {}
    if os.path.exists(REG_DIR):
        for f in os.listdir(REG_DIR):
            if f.endswith(".html"):
                try:
                    with open(os.path.join(REG_DIR, f), "r", encoding="utf-8") as file:
                        soup = BeautifulSoup(file.read(), "html.parser")
                    reg_list = []
                    for tr in soup.select("tr"):
                        tds = tr.select("td")
                        if tds and (jo_span := tr.select_one("span.bl")):
                            reg_list.append({"조문": jo_span.get_text(strip=True), "내용": "\n".join([td.get_text("\n", strip=True) for td in tds])})
                    if reg_list: reg_db[f.replace(".html", "")] = reg_list
                except: pass

    return df_qna, df_case, law_db, reg_db

df_qna, df_case, law_db, reg_db = load_all_data()

# ==========================================
# [3. 일정 및 알람 기능 로직]
# ==========================================
def load_events():
    if os.path.exists(EVENT_PATH):
        try:
            with open(EVENT_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except: pass
    return {}

events = load_events()

# 상단 알람 배너 띄우기
upcoming = []
for d_str, info in events.items():
    if isinstance(info, dict) and info.get("use_alarm"):
        try:
            delta = (datetime.strptime(d_str, "%Y-%m-%d").date() - datetime.now().date()).days
            if 0 <= delta <= info.get("alarm_days", 1):
                upcoming.append({"date": d_str, "d_day": delta, "memo": info["memo"]})
        except: continue

upcoming.sort(key=lambda x: x["d_day"])

st.title("🔍 지적재조사 업무지원 시스템 (Web v1.0) 경상북도")

if upcoming:
    first = upcoming[0]
    d_text = "[오늘]" if first["d_day"] == 0 else f"[{first['d_day']}일 후]"
    st.warning(f"🔔 **중요 예정 알림:** {d_text} {first['memo']}")

# ==========================================
# [4 & 5. 모바일 최적화 화면 배치 및 검색 결과]
# ==========================================

# 1. 메인 화면 한가운데에 직관적인 검색창과 버튼 배치
st.subheader("🔍 지적재조사 통합 검색")

col1, col2 = st.columns([3, 1])
with col1:
    keyword = st.text_input("검색어를 입력하세요", label_visibility="collapsed", placeholder="🔍 검색어 입력 (예: 경계설정)")
with col2:
    search_btn = st.button("검색", use_container_width=True)

# [수정 2] 좌우에 빈 칸(점선/공백)을 만들어 자료 선택 버튼들을 화면 한가운데(중간)로 정렬
_, center_col, _ = st.columns([1, 8, 1])
with center_col:
    mode = st.radio("자료 선택", ["질의회신", "법령검색", "판례검색"], horizontal=True, label_visibility="collapsed")

st.markdown("---")

def highlight_text(text, kw):
    if not kw: return text
    return text.replace(kw, f"<mark style='background-color: yellow;'>{kw}</mark>")

# 데이터 출력 로직
if mode in ["질의회신", "판례검색"]:
    target_df = df_qna if mode == "질의회신" else df_case
    if keyword:
        res = target_df[target_df['제목'].str.contains(keyword, case=False, na=False) | 
                        target_df['내용'].str.contains(keyword, case=False, na=False)]
    else:
        res = target_df 
        
    st.caption(f"총 {len(res)}건의 자료가 있습니다.")
    
    for idx, row in res.head(100).iterrows():
        icon = "🟢" if str(row.get("수정여부")).strip().upper() == "Y" else "📑"
        with st.expander(f"{icon} {row['제목']}"):
            content = row['내용'].replace("\n", "<br>")
            content = highlight_text(content, keyword) if keyword else content
            st.markdown(content, unsafe_allow_html=True)

elif mode == "법령검색":
    # [수정 3] 법령검색도 검색어가 없을 때 상위 30개 조문을 기본으로 쫙 띄워줍니다!
    if keyword:
        count = 0
        for item in law_db:
            if keyword in item['조문'] or keyword in item['법률'] or keyword in item['시행령']:
                count += 1
                with st.expander(f"⚖️ [법령] {item['조문']}"):
                    st.markdown(f"**📜 [법률]**<br>{highlight_text(item['법률'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
                    st.markdown("---")
                    st.markdown(f"**⚙️ [시행령]**<br>{highlight_text(item['시행령'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
        
        for reg_name, reg_data in reg_db.items():
            display_name = reg_name.replace("규정_", "")
            for item in reg_data:
                if keyword in item['조문'] or keyword in item['내용']:
                    count += 1
                    with st.expander(f"📖 [{display_name}] {item['조문']}"):
                        st.markdown(highlight_text(item['내용'].replace("\n", "<br>"), keyword), unsafe_allow_html=True)
        st.caption(f"총 {count}건의 조문/규정이 검색되었습니다.")
        
    else:
        # 검색창이 비어있을 때 터지는 걸 막기 위해 상위 30개만 슬라이싱[:30]해서 노출
        st.caption("💡 법령 정보 상위 30개 조문을 먼저 보여줍니다. 더 찾으시려면 검색어를 입력하세요.")
        for item in law_db[:30]:
            with st.expander(f"⚖️ [법령] {item['조문']}"):
                st.markdown(f"**📜 [법률]**<br>{item['법률'].replace(chr(10), '<br>')}", unsafe_allow_html=True)
                st.markdown("---")
                st.markdown(f"**⚙️ [시행령]**<br>{item['시행령'].replace(chr(10), '<br>')}", unsafe_allow_html=True)
# 하단 푸터
st.markdown("---")
st.caption("v4.0 Web Version - 데이터 수정은 data.xlsx 파일을 변경 후 다시 배포해주세요.")
