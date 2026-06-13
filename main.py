import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import json
import re
import base64
import gspread
import traceback
from bs4 import BeautifulSoup
from datetime import datetime

# ==========================================
# [1. 웹 페이지 기본 설정 및 구글 시트 연동]
# ==========================================
st.set_page_config(page_title="지적재조사 통합 업무지원 시스템", page_icon="🔍", layout="wide")

DATA_DIR = "data"
EXCEL_PATH = f"{DATA_DIR}/data.xlsx"
LAW_HTML_PATH = f"{DATA_DIR}/지적재조사에 관한 특별법(인용조문 3단비교).html"
REG_DIR = f"{DATA_DIR}/규정"

# 🛠️ 구글 시트 실시간 연동 핵심 함수 (에러 추적기 탑재)
def get_google_sheet():
    try:
        # 보이지 않는 특수 공백문자(\xa0)를 일반 공백으로 자동 청소
        raw_json = st.secrets["google_json"].replace('\xa0', ' ').replace('\u00A0', ' ')
        secret_json = json.loads(raw_json)
        
        gc = gspread.service_account_from_dict(secret_json)
        
        # URL 앞뒤에 혹시 모를 공백 제거
        clean_url = st.secrets["spreadsheet_url"].strip()
        sh = gc.open_by_url(clean_url)
        return sh.get_worksheet(0) # 첫 번째 시트 사용
        
    except Exception as e:
        st.error("🚨 구글 시트 연결 중 에러가 발생했습니다!")
        st.error(f"원인: {e}")
        st.code(traceback.format_exc(), language="text")
        return None

def load_events_from_google():
    sheet = get_google_sheet()
    if sheet is None: return {}
    try:
        records = sheet.get_all_records()
        events_dict = {}
        for r in records:
            d_str = str(r.get("날짜", "")).strip()
            memo = str(r.get("메모", "")).strip()
            use_alarm = str(r.get("알람여부", "")).strip().upper() in ["TRUE", "Y", "YES", "1"]
            try: alarm_days = int(r.get("알람기간", 1))
            except: alarm_days = 1
            
            if d_str:
                events_dict[d_str] = {"memo": memo, "use_alarm": use_alarm, "alarm_days": alarm_days}
        return events_dict
    except Exception as e:
        st.error(f"데이터 읽기 실패: {e}")
        return {}

def save_event_to_google(date_key, memo, use_alarm, alarm_days):
    sheet = get_google_sheet()
    if sheet is None: return
    try:
        records = sheet.get_all_records()
        row_idx = -1
        for i, r in enumerate(records):
            if str(r.get("날짜", "")).strip() == date_key:
                row_idx = i + 2 # 헤더라인 제외 보정
                break
        
        alarm_val = "TRUE" if use_alarm else "FALSE"
        if not memo:
            if row_idx != -1: sheet.delete_rows(row_idx) # 메모 비우면 삭제
        else:
            if row_idx != -1:
                sheet.update_cell(row_idx, 2, memo)
                sheet.update_cell(row_idx, 3, alarm_val)
                sheet.update_cell(row_idx, 4, alarm_days)
            else:
                sheet.append_row([date_key, memo, alarm_val, alarm_days])
    except Exception as e:
        st.error(f"구글 시트 저장 오류: {e}")

# 구글 시트로부터 실시간 일정 읽어오기
events = load_events_from_google()

# ==========================================
# [보조 마법: 복사 버튼 생성기]
# ==========================================
def custom_copy_button(text_to_copy):
    b64_text = base64.b64encode(text_to_copy.encode('utf-8')).decode('utf-8')
    button_html = f"""
    <body style="margin: 0; padding: 0;">
        <div style="text-align: right;">
            <button id="copyBtn" style="border: 1px solid #bdc3c7; border-radius: 5px; padding: 5px 10px; background-color: white; color: #34495e; font-size: 13px; font-weight: bold; cursor: pointer;">
                📋 복사하기
            </button>
        </div>
        <script>
            document.getElementById("copyBtn").addEventListener("click", function() {{
                const decodedText = decodeURIComponent(escape(window.atob("{b64_text}")));
                if (navigator.clipboard && window.isSecureContext) {{
                    navigator.clipboard.writeText(decodedText).then(() => {{ changeBtnState(); }});
                }} else {{
                    let textArea = document.createElement("textarea");
                    textArea.value = decodedText;
                    textArea.style.position = "absolute";
                    textArea.style.left = "-999999px";
                    document.body.prepend(textArea);
                    textArea.select();
                    try {{ document.execCommand('copy'); changeBtnState(); }} catch (error) {{ console.error(error); }} finally {{ textArea.remove(); }}
                }}
                function changeBtnState() {{
                    var btn = document.getElementById("copyBtn");
                    btn.innerHTML = "✅ 복사 완료!";
                    btn.style.color = "#27ae60"; btn.style.borderColor = "#27ae60";
                    setTimeout(function() {{
                        btn.innerHTML = "📋 복사하기"; btn.style.color = "#34495e"; btn.style.borderColor = "#bdc3c7";
                    }}, 2000);
                }}
            }});
        </script>
    </body>
    """
    components.html(button_html, height=35)

# ==========================================
# [2. 데이터 파싱 함수 (지적재조사 자료)]
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
        try: return pd.read_excel(EXCEL_PATH, sheet_name=sheet_name).fillna("").astype(str)
        except: return pd.DataFrame(columns=["제목", "내용", "수정여부"])

    df_qna = safe_load_sheet("질의회신")
    df_case = safe_load_sheet("판례검색")

    reg_db = {}
    if os.path.exists(REG_DIR):
        for f in os.listdir(REG_DIR):
            if f.endswith(".html"):
                file_path = os.path.join(REG_DIR, f)
                html_content = ""
                try:
                    with open(file_path, "r", encoding="utf-8") as file: html_content = file.read()
                except UnicodeDecodeError:
                    try:
                        with open(file_path, "r", encoding="cp949") as file: html_content = file.read()
                    except: continue
                
                if html_content:
                    soup = BeautifulSoup(html_content, "html.parser")
                    reg_list = []
                    for tr in soup.select("tr"):
                        tds = tr.select("td")
                        if tds and (jo_span := tr.select_one("span.bl")):
                            reg_list.append({"조문": jo_span.get_text(strip=True), "내용": "\n".join([td.get_text("\n", strip=True) for td in tds])})
                    
                    if not reg_list:
                        text_lines = soup.get_text(separator="\n").split("\n")
                        current_jo, current_content = "전체 내용", []
                        for line in text_lines:
                            line = line.strip()
                            if not line: continue
                            if re.match(r'^제\s*\d+\s*조', line) or re.match(r'^\[별표', line):
                                if current_content: reg_list.append({"조문": current_jo, "내용": "\n".join(current_content)})
                                current_jo, current_content = line, [line]
                            else: current_content.append(line)
                        if current_content: reg_list.append({"조문": current_jo, "내용": "\n".join(current_content)})

                    if reg_list: reg_db[f.replace(".html", "")] = reg_list

    return df_qna, df_case, law_db, reg_db

df_qna, df_case, law_db, reg_db = load_all_data()

# ==========================================
# [3. D-Day 알람 배너 로직 (구글 데이터 기반)]
# ==========================================
upcoming = []
for d_str, info in events.items():
    if isinstance(info, dict) and info.get("use_alarm"):
        try:
            delta = (datetime.strptime(d_str, "%Y-%m-%d").date() - datetime.now().date()).days
            if 0 <= delta <= info.get("alarm_days", 1):
                upcoming.append({"date": d_str, "d_day": delta, "memo": info["memo"]})
        except: continue
upcoming.sort(key=lambda x: x["d_day"])

if upcoming:
    first = upcoming[0]
    d_text = "[오늘]" if first["d_day"] == 0 else f"[{first['d_day']}일 후]"
    st.warning(f"🔔 **중요 예정 업무 알림:** {d_text} {first['memo']}")

# ==========================================
# [4. 모바일 화면 배치 및 카테고리 분리]
# ==========================================
st.markdown("<h4 style='text-align: center; color: #2c3e50; font-size: 1.3rem; margin-top: -40px; margin-bottom: 10px;'>🔍 지적재조사 통합 검색</h4>", unsafe_allow_html=True)

col1, col2 = st.columns([3, 1])
with col1:
    keyword = st.text_input("검색어를 입력하세요", label_visibility="collapsed", placeholder="🔍 검색어 입력 (예: 경계설정)")
with col2:
    search_btn = st.button("검색", use_container_width=True)

only_title = st.checkbox("☑️ 제목만 검색", value=True)

tabs = ["📑 질의회신", "⚖️ 법령", "🏢 업무규정", "📐 측량규정", "🧑‍⚖️ 판례", "📅 공유달력"]
mode = st.radio("자료 선택", tabs, horizontal=True, label_visibility="collapsed")

st.markdown("---")

def highlight_text(text, kw):
    if not kw: return text
    return text.replace(kw, f"<mark style='background-color: yellow;'>{kw}</mark>")

# ==========================================
# [5. 카테고리별 출력 및 일정 관리 등록]
# ==========================================
if mode in ["📑 질의회신", "🧑‍⚖️ 판례"]:
    target_df = df_qna if mode == "📑 질의회신" else df_case
    if keyword:
        res = target_df[target_df['제목'].str.contains(keyword, case=False, na=False)] if only_title else target_df[target_df['제목'].str.contains(keyword, case=False, na=False) | target_df['내용'].str.contains(keyword, case=False, na=False)]
    else: res = target_df 
        
    st.caption(f"총 {len(res)}건의 자료가 있습니다.")
    for idx, row in res.iterrows(): 
        icon = "🟢" if str(row.get("수정여부")).strip().upper() == "Y" else "📑"
        with st.expander(f"{icon} {row['제목']}"):
            custom_copy_button(f"[{row['제목']}]\n{row['내용']}")
            content = row['내용'].replace("\n", "<br>")
            st.markdown(highlight_text(content, keyword) if keyword else content, unsafe_allow_html=True)

elif mode == "⚖️ 법령":
    if keyword:
        count = 0
        for item in law_db:
            match = (keyword in item['조문']) if only_title else (keyword in item['조문'] or keyword in item['법률'] or keyword in item['시행령'] or keyword in item['시행규칙'])
            if match:
                count += 1
                with st.expander(f"⚖️ [법령] {item['조문']}"):
                    custom_copy_button(f"[{item['조문']}]\n\n[법률]\n{item['법률']}\n\n[시행령]\n{item['시행령']}")
                    st.markdown(f"**📜 [법률]**<br>{highlight_text(item['법률'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
                    st.markdown("---")
                    st.markdown(f"**⚙️ [시행령]**<br>{highlight_text(item['시행령'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
        st.caption(f"총 {count}건의 조문이 검색되었습니다.")
    else:
        st.caption(f"총 {len(law_db)}개의 전체 조문입니다.")
        for item in law_db:
            with st.expander(f"⚖️ [법령] {item['조문']}"):
                custom_copy_button(f"[{item['조문']}]\n\n[법률]\n{item['법률']}\n\n[시행령]\n{item['시행령']}")
                st.markdown(f"**📜 [법률]**<br>{item['법률'].replace(chr(10), '<br>')}", unsafe_allow_html=True)
                st.markdown("---")
                st.markdown(f"**⚙️ [시행령]**<br>{item['시행령'].replace(chr(10), '<br>')}", unsafe_allow_html=True)

elif mode in ["🏢 업무규정", "📐 측량규정"]:
    is_survey = (mode == "📐 측량규정")
    count = 0
    for reg_name, reg_data in reg_db.items():
        if (is_survey and "측량" in reg_name) or (not is_survey and "측량" not in reg_name):
            display_name = reg_name.replace("규정_", "")
            for item in reg_data:
                match = (keyword in item['조문']) if keyword and only_title else (keyword in item['조문'] or keyword in item['내용']) if keyword else True
                if match:
                    count += 1
                    with st.expander(f"📖 [{display_name}] {item['조문']}"):
                        custom_copy_button(f"[{display_name} {item['조문']}]\n{item['내용']}")
                        content = item['내용'].replace("\n", "<br>")
                        st.markdown(highlight_text(content, keyword) if keyword else content, unsafe_allow_html=True)
    st.caption(f"총 {count}건의 목록이 있습니다.")

elif mode == "📅 공유달력":
    st.subheader("📅 지적재조사팀 실시간 공유 달력")
    
    # 1. 일정 등록 폼
    with st.form("google_calendar_form"):
        e_date = st.date_input("날짜 선택")
        e_memo = st.text_area("일정 메모 (비우고 저장하면 일정이 삭제됩니다)")
        e_alarm = st.checkbox("🔔 알람 켜기 배너 노출")
        e_days = st.selectbox("알람 기간 (며칠 전부터 알릴까요?)", [0, 1, 3, 5, 7, 10, 30], index=1)
        
        if st.form_submit_button("일정 저장하기"):
            date_key = e_date.strftime("%Y-%m-%d")
            with st.spinner("구글 클라우드 시트에 동기화 중..."):
                save_event_to_google(date_key, e_memo, e_alarm, e_days)
            st.success("구글 시트에 실시간 공유 완료되었습니다!")
            st.rerun()

    # 2. 등록된 전체 일정 리스트 표시
    st.markdown("---")
    st.subheader("📌 예정된 일정 목록 (실시간 연동)")
    if events:
        sorted_events = sorted(events.items())
        for d_key, info in sorted_events:
            alarm_icon = "🔔" if info.get("use_alarm") else "📌"
            with st.expander(f"{alarm_icon} [{d_key}] {info['memo'][:20]}..."):
                st.write(f"**날짜:** {d_key}")
                st.write(f"**상세 내용:** {info['memo']}")
                st.write(f"**알람 설정 여부:** {'켜짐 (' + str(info['alarm_days']) + '일 전부터)' if info['use_alarm'] else '꺼짐'}")
    else:
        st.info("등록된 팀 일정이 없습니다. 위 양식에서 일정을 새로 등록해 보세요!")

# 하단 푸터
st.markdown("---")
st.caption("v5.1 Web Cloud Version - 구글 스프레드시트 실시간 동기화 지원")
