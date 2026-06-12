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
st.set_page_config(page_title="지적재조사 통합 업무지원 시스템", page_icon="🔍", layout="wide")

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

st.title("🔍 지적재조사 업무지원 시스템 (Web v4.0)")

if upcoming:
    first = upcoming[0]
    d_text = "[오늘]" if first["d_day"] == 0 else f"[{first['d_day']}일 후]"
    st.warning(f"🔔 **중요 예정 알림:** {d_text} {first['memo']}")

# ==========================================
# [4. 사이드바 (통합 검색 및 일정 등록)]
# ==========================================
st.sidebar.header("통합 검색")
mode = st.sidebar.selectbox("검색 모드", ["질의회신", "법령검색", "판례검색"])
keyword = st.sidebar.text_input("검색어를 입력하세요")
only_title = st.sidebar.checkbox("제목만 검색", value=True)

st.sidebar.markdown("---")
st.sidebar.subheader("📅 일정 및 알람 등록")
with st.sidebar.form("calendar_form"):
    e_date = st.date_input("날짜 선택")
    e_memo = st.text_area("일정 메모 (알람 내용)")
    e_alarm = st.checkbox("🔔 알람 켜기")
    e_days = st.selectbox("알람 기간 (며칠 전부터 알릴까요?)", [0, 1, 3, 5, 7, 10, 30], index=1)
    
    if st.form_submit_button("일정 저장"):
        date_key = e_date.strftime("%Y-%m-%d")
        if e_memo:
            events[date_key] = {"memo": e_memo, "use_alarm": e_alarm, "alarm_days": e_days}
        elif date_key in events:
            del events[date_key]
            
        with open(EVENT_PATH, "w", encoding="utf-8") as f:
            json.dump(events, f, ensure_ascii=False, indent=4)
        st.success("일정이 저장되었습니다! 새로고침하면 알람 배너에 반영됩니다.")

# ==========================================
# [5. 메인 화면 검색 결과 출력]
# ==========================================
def highlight_text(text, kw):
    """검색어 노란색 하이라이트 효과"""
    if not kw: return text
    return text.replace(kw, f"<mark style='background-color: yellow;'>{kw}</mark>")

if mode in ["질의회신", "판례검색"]:
    target_df = df_qna if mode == "질의회신" else df_case
    if keyword:
        if only_title:
            res = target_df[target_df['제목'].str.contains(keyword, case=False, na=False)]
        else:
            res = target_df[target_df['제목'].str.contains(keyword, case=False, na=False) | 
                            target_df['내용'].str.contains(keyword, case=False, na=False)]
        
        st.subheader(f"총 {len(res)}건의 결과가 있습니다.")
        for idx, row in res.iterrows():
            icon = "🟢" if str(row.get("수정여부")).strip().upper() == "Y" else "📑"
            with st.expander(f"{icon} {row['제목']}"):
                content = row['내용'].replace("\n", "<br>")
                content = highlight_text(content, keyword)
                st.markdown(content, unsafe_allow_html=True)
    else:
        st.info("왼쪽 검색창에 검색어를 입력하시면 결과가 나타납니다.")

elif mode == "법령검색":
    if keyword:
        count = 0
        # 1. 법령 데이터 검색
        for item in law_db:
            if (only_title and keyword in item['조문']) or (not only_title and any(keyword in item[k] for k in ['조문', '법률', '시행령', '시행규칙'])):
                count += 1
                with st.expander(f"⚖️ [법령] {item['조문']}"):
                    st.markdown(f"**📜 [법률]**<br>{highlight_text(item['법률'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
                    st.markdown("---")
                    st.markdown(f"**⚙️ [시행령]**<br>{highlight_text(item['시행령'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
                    st.markdown("---")
                    st.markdown(f"**📝 [시행규칙]**<br>{highlight_text(item['시행규칙'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
        
        # 2. 규정 데이터 검색
        for reg_name, reg_data in reg_db.items():
            display_name = reg_name.replace("규정_", "")
            for item in reg_data:
                if (only_title and keyword in item['조문']) or (not only_title and (keyword in item['조문'] or keyword in item['내용'])):
                    count += 1
                    with st.expander(f"📖 [{display_name}] {item['조문']}"):
                        st.markdown(highlight_text(item['내용'].replace("\n", "<br>"), keyword), unsafe_allow_html=True)
                        
        st.subheader(f"총 {count}건의 조문/규정이 검색되었습니다.")
    else:
        st.info("조문 번호나 키워드를 입력하세요. (예: 제5조, 경계설정)")

# 하단 푸터
st.markdown("---")
st.caption("v4.0 Web Version - 데이터 수정은 data.xlsx 파일을 변경 후 다시 배포해주세요.")
