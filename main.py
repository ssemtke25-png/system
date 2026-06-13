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

def get_google_sheet():
    try:
        raw_json = st.secrets["google_json"].replace('\xa0', ' ').replace('\u00A0', ' ')
        secret_json = json.loads(raw_json)
        gc = gspread.service_account_from_dict(secret_json)
        clean_url = st.secrets["spreadsheet_url"].strip()
        return gc.open_by_url(clean_url)
    except Exception as e:
        st.error("🚨 구글 시트 연결 에러!")
        return None

# 시트 탭 분리 (0번: 달력, 1번: 공지사항)
sh = get_google_sheet()
sheet_main = sh.get_worksheet(0) if sh else None
sheet_notice = sh.get_worksheet(1) if sh else None

# ==========================================
# [공지사항 데이터 관리 함수]
# ==========================================
def load_notice():
    try: return sheet_notice.get_all_records() if sheet_notice else []
    except: return []

def save_notice(content):
    if sheet_notice:
        sheet_notice.clear()
        sheet_notice.append_row(["날짜", "내용"])
        sheet_notice.append_row([datetime.now().strftime("%Y-%m-%d"), content])

# ==========================================
# [달력 데이터 관리 함수]
# ==========================================
def load_events_from_google():
    if sheet_main is None: return []
    try:
        records = sheet_main.get_all_records()
        events_list = []
        for i, r in enumerate(records):
            d_str = str(r.get("날짜", "")).strip()
            memo = str(r.get("메모", "")).strip()
            region = str(r.get("시군구", "공통")).strip()
            use_alarm = str(r.get("알람여부", "")).strip().upper() in ["TRUE", "Y", "YES", "1"]
            try: alarm_days = int(r.get("알람기간", 1))
            except: alarm_days = 1
            
            if d_str:
                events_list.append({"date": d_str, "memo": memo, "use_alarm": use_alarm, "alarm_days": alarm_days, "region": region, "row_idx": i + 2})
        return events_list
    except Exception as e:
        return []

def save_event_to_google(date_key, memo, use_alarm, alarm_days, region):
    if sheet_main is None: return
    try:
        alarm_val = "TRUE" if use_alarm else "FALSE"
        if memo:
            sheet_main.append_row([date_key, memo, alarm_val, alarm_days, region])
    except Exception as e:
        st.error(f"저장 오류: {e}")

def delete_event_from_google(row_idx):
    if sheet_main is not None:
        try:
            sheet_main.delete_rows(row_idx)
        except Exception as e:
            st.error(f"삭제 오류: {e}")

all_events = load_events_from_google()
notices = load_notice()

# ==========================================
# [보조 마법: 스마트폰 네이티브 공유 팝업 버튼]
# ==========================================
# (주의: 모바일 메모리 과부하 방지를 위해 복사버튼(custom_copy_button)은 제거하고, 스마트폰 기본 텍스트 선택(롱터치) 기능을 사용하도록 최적화했습니다.)
def native_share_button(region, date, memo):
    app_url = "https://system-ydyhcgqqhe6dncgekqklcv.streamlit.app"
    share_title = "지적재조사팀 일정 공유"
    share_text = f"📢 [지적재조사팀 중요 일정 안내]\n\n🏢 담당 구역: {region}\n📅 지정 날짜: {date}\n📝 세부 내용: {memo}\n"
    
    b64_title = base64.b64encode(share_title.encode('utf-8')).decode('utf-8')
    b64_text = base64.b64encode(share_text.encode('utf-8')).decode('utf-8')
    b64_url = base64.b64encode(app_url.encode('utf-8')).decode('utf-8')
    
    button_html = f"""
    <body style="margin: 0; padding: 0;">
        <button id="shareBtn" style="border: 1px solid #3498db; border-radius: 5px; padding: 5px 10px; background-color: #3498db; color: white; font-size: 13px; font-weight: bold; cursor: pointer; width: 100%;">
            📲 일정 공유하기 (카톡/문자)
        </button>
        <script>
            document.getElementById("shareBtn").addEventListener("click", async function() {{
                const title = decodeURIComponent(escape(window.atob("{b64_title}")));
                const text = decodeURIComponent(escape(window.atob("{b64_text}")));
                const url = decodeURIComponent(escape(window.atob("{b64_url}")));
                if (navigator.share) {{
                    try {{
                        await navigator.share({{ title: title, text: text, url: url }});
                    }} catch (err) {{ console.log("공유 취소 또는 오류:", err); }}
                }} else {{
                    const fullText = text + "\\n🔗 팀 공유 달력 바로가기:\\n" + url;
                    if (navigator.clipboard && window.isSecureContext) {{
                        navigator.clipboard.writeText(fullText);
                    }} else {{
                        let textArea = document.createElement("textarea");
                        textArea.value = fullText;
                        textArea.style.position = "absolute";
                        textArea.style.left = "-999999px";
                        document.body.prepend(textArea);
                        textArea.select();
                        try {{ document.execCommand('copy'); }} catch(e) {{}} finally {{ textArea.remove(); }}
                    }}
                    var btn = document.getElementById("shareBtn");
                    var originalText = btn.innerHTML;
                    btn.innerHTML = "✅ 복사 완료! (PC는 직접 붙여넣으세요)";
                    btn.style.backgroundColor = "#27ae60"; btn.style.borderColor = "#27ae60";
                    setTimeout(function() {{ 
                        btn.innerHTML = originalText; 
                        btn.style.backgroundColor = "#3498db"; btn.style.borderColor = "#3498db";
                    }}, 2500);
                }}
            }});
        </script>
    </body>
    """
    components.html(button_html, height=35)

# ==========================================
# [2. 데이터 파싱 함수]
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
# [3. 최상단 배너 (알람 및 공지사항)]
# ==========================================
# 1) D-Day 알람 배너
upcoming = []
for info in all_events:
    if info.get("use_alarm"):
        try:
            delta = (datetime.strptime(info["date"], "%Y-%m-%d").date() - datetime.now().date()).days
            if 0 <= delta <= info.get("alarm_days", 1):
                upcoming.append({"date": info["date"], "d_day": delta, "memo": info["memo"], "region": info["region"]})
        except: continue
upcoming.sort(key=lambda x: x["d_day"])

if upcoming:
    first = upcoming[0]
    d_text = "[오늘]" if first["d_day"] == 0 else f"[{first['d_day']}일 후]"
    st.warning(f"🔔 **중요 예정 업무 알림 [{first['region']}]:** {d_text} {first['memo']}")

# 2) 전체 공지사항 배너
if notices:
    st.info(f"📢 **[전체 공지사항]** {notices[-1]['내용']}")

# ==========================================
# [4. 화면 배치]
# ==========================================
st.markdown("<h4 style='text-align: center; color: #2c3e50; font-size: 1.3rem; margin-top: 15px; margin-bottom: 15px;'>🔍 지적재조사 통합 검색</h4>", unsafe_allow_html=True)

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
                    st.markdown(f"**📜 [법률]**<br>{highlight_text(item['법률'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
                    st.markdown("---")
                    st.markdown(f"**⚙️ [시행령]**<br>{highlight_text(item['시행령'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
                    if item['시행규칙'].strip():
                        st.markdown("---")
                        st.markdown(f"**📋 [시행규칙]**<br>{highlight_text(item['시행규칙'].replace(chr(10), '<br>'), keyword)}", unsafe_allow_html=True)
        st.caption(f"총 {count}건의 조문이 검색되었습니다.")
    else:
        st.caption(f"총 {len(law_db)}개의 전체 조문입니다.")
        for item in law_db:
            with st.expander(f"⚖️ [법령] {item['조문']}"):
                st.markdown(f"**📜 [법률]**<br>{item['법률'].replace(chr(10), '<br>')}", unsafe_allow_html=True)
                st.markdown("---")
                st.markdown(f"**⚙️ [시행령]**<br>{item['시행령'].replace(chr(10), '<br>')}", unsafe_allow_html=True)
                if item['시행규칙'].strip():
                    st.markdown("---")
                    st.markdown(f"**📋 [시행규칙]**<br>{item['시행규칙'].replace(chr(10), '<br>')}", unsafe_allow_html=True)

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
                        content = item['내용'].replace("\n", "<br>")
                        st.markdown(highlight_text(content, keyword) if keyword else content, unsafe_allow_html=True)
    st.caption(f"총 {count}건의 목록이 있습니다.")

elif mode == "📅 공유달력":
    st.subheader("🔐 지역별 보안 공유 달력")
    
    # 0. 보안 로그인 시스템
    regions = ["포항시", "경주시", "김천시", "안동시", "구미시", "영주시", "영천시", "상주시", "문경시", "경산시", "의성군", "청송군", "영양군", "영덕군", "청도군", "고령군", "성주군", "칠곡군", "예천군", "봉화군", "울진군", "울릉군", "경상북도(총괄)"]
    selected_region = st.selectbox("📌 담당 시/군을 선택하세요", regions)
    
    col_pw, col_btn = st.columns([3, 1])
    with col_pw:
        entered_pw = st.text_input("🔑 비밀번호 4자리를 입력하세요", type="password")
    with col_btn:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        login_btn = st.button("확인", use_container_width=True)
    
    is_unlocked = False
    if entered_pw or login_btn:
        if entered_pw:
            try:
                if entered_pw == st.secrets["passwords"][selected_region]:
                    is_unlocked = True
                else:
                    st.error("❌ 비밀번호가 일치하지 않습니다.")
            except KeyError:
                st.warning("⚠️ 이 지역의 비밀번호가 아직 설정되지 않았습니다. 관리자에게 문의하세요.")
        elif login_btn:
            st.warning("비밀번호를 입력해주세요.")
            
    # 비밀번호 통과 시 달력 오픈
    if is_unlocked:
        st.success(f"🔓 [{selected_region}] 전용 달력에 접속되었습니다!")
        st.markdown("---")
        
        # 1. 일정 등록 폼
        with st.form("google_calendar_form"):
            st.write(f"**[{selected_region}] 새로운 일정 등록**")
            e_date = st.date_input("날짜 선택")
            e_memo = st.text_area("일정 메모 (하루에 여러 개의 일정을 등록할 수 있습니다)")
            e_alarm = st.checkbox("🔔 상단 D-Day 알람 켜기")
            e_days = st.selectbox("알람 기간 (며칠 전부터 알릴까요?)", [0, 1, 3, 5, 7, 10, 30], index=1)
            
            if st.form_submit_button("일정 저장 및 동기화"):
                if e_memo.strip():
                    with st.spinner("구글 시트에 보안 저장 중..."):
                        save_event_to_google(e_date.strftime("%Y-%m-%d"), e_memo, e_alarm, e_days, selected_region)
                    st.success(f"✅ {selected_region} 일정이 안전하게 추가되었습니다!")
                    st.rerun()
                else:
                    st.warning("일정 메모를 입력해주세요.")

        st.markdown("---")

        # 2. 일정 목록 조회
        if selected_region == "경상북도(총괄)":
            st.info("👑 총괄 관리자 모드: 도청 자체 일정을 등록하고, 모든 시군의 일정을 열람/관리합니다.")
            
            # [관리자 전용 공지사항 등록 폼]
            with st.expander("📢 관리자용: 팀 전체 공지사항 등록"):
                new_notice = st.text_area("앱 최상단에 띄울 공지 내용을 입력하세요")
                if st.button("공지사항 업데이트", use_container_width=True):
                    if new_notice.strip():
                        with st.spinner("공지사항 등록 중..."):
                            save_notice(new_notice)
                        st.success("공지가 업데이트되었습니다!")
                        st.rerun()
                    else:
                        st.warning("내용을 입력해주세요.")

            st.subheader("📋 전체 예정된 일정 목록")
            if all_events:
                all_events.sort(key=lambda x: x["date"])
                for info in all_events:
                    alarm_icon = "🔔" if info.get("use_alarm") else "📌"
                    with st.expander(f"{alarm_icon} [{info['region']}] {info['date']} | {info['memo'][:15]}..."):
                        st.write(f"**🏢 담당 지역:** {info['region']}")
                        st.write(f"**📅 날짜:** {info['date']}")
                        st.write(f"**📝 상세 내용:** {info['memo']}")
                        st.write(f"**🔔 알람 여부:** {'켜짐 (' + str(info['alarm_days']) + '일 전부터)' if info['use_alarm'] else '꺼짐'}")
                        
                        col_share, col_del = st.columns(2)
                        with col_share:
                            native_share_button(info['region'], info['date'], info['memo'])
                        with col_del:
                            if st.button("🗑️ 일정 삭제", key=f"del_admin_{info['row_idx']}", use_container_width=True):
                                with st.spinner("삭제 중..."):
                                    delete_event_from_google(info['row_idx'])
                                st.rerun()
            else:
                st.write("등록된 전체 일정이 없습니다.")
                
        else:
            st.subheader(f"📋 [{selected_region}] 예정된 일정 목록")
            region_events = [e for e in all_events if e["region"] == selected_region]
            
            if region_events:
                region_events.sort(key=lambda x: x["date"])
                for info in region_events:
                    alarm_icon = "🔔" if info.get("use_alarm") else "📌"
                    with st.expander(f"{alarm_icon} [{info['date']}] {info['memo'][:20]}..."):
                        st.write(f"**날짜:** {info['date']}")
                        st.write(f"**상세 내용:** {info['memo']}")
                        
                        col_share, col_del = st.columns(2)
                        with col_share:
                            native_share_button(info['region'], info['date'], info['memo'])
                        with col_del:
                            if st.button("🗑️ 일정 삭제", key=f"del_{info['row_idx']}", use_container_width=True):
                                with st.spinner("삭제 중..."):
                                    delete_event_from_google(info['row_idx'])
                                st.rerun()
            else:
                st.info("등록된 일정이 없습니다. 위 양식에서 첫 일정을 등록해 보세요!")

# 하단 푸터
st.markdown("---")
st.caption("v5.8 UX Pro Update - 모바일 최적화 및 공지사항 기능 통합")
