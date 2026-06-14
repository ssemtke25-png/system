import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import json
import re
import base64
import gspread
import html
import qrcode
from io import BytesIO
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

sh = get_google_sheet()
sheet_main = sh.get_worksheet(0) if sh else None
sheet_notice = sh.get_worksheet(1) if sh else None

def load_notice():
    try: return sheet_notice.get_all_records() if sheet_notice else []
    except: return []

def save_notice(content):
    if sheet_notice:
        sheet_notice.clear()
        sheet_notice.append_row(["날짜", "내용"])
        sheet_notice.append_row([datetime.now().strftime("%Y-%m-%d"), content])

def delete_notice():
    if sheet_notice:
        try:
            sheet_notice.clear()
            sheet_notice.append_row(["날짜", "내용"])
        except: pass

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
    if sheet_main:
        sheet_main.append_row([date_key, memo, "TRUE" if use_alarm else "FALSE", alarm_days, region])

def delete_event_from_google(row_idx):
    if sheet_main:
        sheet_main.delete_rows(row_idx)

all_events = load_events_from_google()
notices = load_notice()

def native_share_button(region, date, memo):
    app_url = "https://system-ydyhcgqqhe6dncgekqklcv.streamlit.app"
    share_text = f"📢 [지적재조사팀 중요 일정 안내]\n\n🏢 담당 구역: {region}\n📅 지정 날짜: {date}\n📝 세부 내용: {memo}\n"
    b64_text = base64.b64encode(share_text.encode('utf-8')).decode('utf-8')
    b64_url = base64.b64encode(app_url.encode('utf-8')).decode('utf-8')
    
    button_html = f"""
    <body style="margin: 0; padding: 0;">
        <button id="shareBtn" style="border: 1px solid #3498db; border-radius: 5px; padding: 5px 10px; background-color: #3498db; color: white; font-size: 13px; font-weight: bold; cursor: pointer; width: 100%;">
            📲 일정 공유하기 (카톡/문자)
        </button>
        <script>
            document.getElementById("shareBtn").addEventListener("click", async function() {{
                const text = decodeURIComponent(escape(window.atob("{b64_text}")));
                const url = decodeURIComponent(escape(window.atob("{b64_url}")));
                if (navigator.share) {{
                    try {{ await navigator.share({{ title: "지적재조사 일정", text: text, url: url }}); }} catch (err) {{}}
                }} else {{
                    const fullText = text + "\\n🔗 시스템 바로가기:\\n" + url;
                    if (navigator.clipboard && window.isSecureContext) {{ navigator.clipboard.writeText(fullText); }} 
                    else {{ let ta = document.createElement("textarea"); ta.value = fullText; ta.style.position = "absolute"; ta.style.left = "-9999px"; document.body.prepend(ta); ta.select(); try {{ document.execCommand('copy'); }} catch(e) {{}} finally {{ ta.remove(); }} }}
                    var btn = document.getElementById("shareBtn");
                    var orig = btn.innerHTML;
                    btn.innerHTML = "✅ 복사 완료! (PC는 붙여넣기 하세요)"; btn.style.backgroundColor = "#27ae60";
                    setTimeout(() => {{ btn.innerHTML = orig; btn.style.backgroundColor = "#3498db"; }}, 2500);
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
def load_all_data_final_v8():
    def clean_reg_text(text):
        match = re.search(r'(【질의회신|\[질의회신|【참고판례|\[참고판례|\[질의요지\])', text)
        if match: return text[:match.start()].strip()
        return text.strip()

    law_db = []
    if os.path.exists(LAW_HTML_PATH):
        try:
            with open(LAW_HTML_PATH, "r", encoding="utf-8") as f:
                soup = BeautifulSoup(f, "html.parser")
            for tr in soup.select("tr"):
                tds = tr.select("td")
                if len(tds) >= 3 and (jo_span := tds[0].select_one("span.bl")):
                    law_db.append({"조문": jo_span.get_text(strip=True), "법률": tds[0].get_text("\n", strip=True), "시행령": tds[1].get_text("\n", strip=True), "시행규칙": tds[2].get_text("\n", strip=True)})
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
                except:
                    try:
                        with open(file_path, "r", encoding="cp949") as file: html_content = file.read()
                    except: continue
                
                if html_content:
                    soup = BeautifulSoup(html_content, "html.parser")
                    reg_list = []
                    for tr in soup.select("tr"):
                        tds = tr.select("td")
                        if tds and (jo_span := tr.select_one("span.bl")):
                            raw_content = "\n".join([td.get_text("\n", strip=True) for td in tds])
                            cleaned_content = clean_reg_text(raw_content)
                            if cleaned_content: reg_list.append({"조문": jo_span.get_text(strip=True), "내용": cleaned_content})
                    
                    if not reg_list:
                        text_lines = soup.get_text(separator="\n").split("\n")
                        current_jo, current_content = "전체 내용", []
                        for line in text_lines:
                            line = line.strip()
                            if not line: continue
                            if re.match(r'^제\s*\d+\s*조', line) or re.match(r'^\[별표', line):
                                if current_content:
                                    cleaned_content = clean_reg_text("\n".join(current_content))
                                    if cleaned_content: reg_list.append({"조문": current_jo, "내용": cleaned_content})
                                current_jo, current_content = line, [line]
                            else: current_content.append(line)
                        if current_content:
                            cleaned_content = clean_reg_text("\n".join(current_content))
                            if cleaned_content: reg_list.append({"조문": current_jo, "내용": cleaned_content})

                    if reg_list: reg_db[f.replace(".html", "")] = reg_list

    return df_qna, df_case, law_db, reg_db

df_qna, df_case, law_db, reg_db = load_all_data_final_v8()

# ==========================================
# [3. 화면 뷰 상태 관리 및 기억 장치]
# ==========================================
if 'view_mode' not in st.session_state:
    st.session_state.view_mode = 'main'
if 'view_law_data' not in st.session_state:
    st.session_state.view_law_data = None
if 'active_tab' not in st.session_state:
    st.session_state.active_tab = "📑 질의회신"
if 'saved_region' not in st.session_state:
    st.session_state.saved_region = "포항시"

def render_safe_html(text, kw=""):
    safe = html.escape(str(text)).replace("\n", "<br>")
    if kw:
        safe_kw = html.escape(str(kw))
        if safe_kw:
            safe = safe.replace(safe_kw, f"<mark style='background-color: yellow;'>{safe_kw}</mark>")
    return f'<div translate="no" class="notranslate" style="line-height:1.6;">{safe}</div>'

# 🚨 법령 상세 보기 화면 (쏙 들어가기 모드)
if st.session_state.view_mode == 'law_detail':
    st.markdown("<h4 style='text-align: center; color: #2c3e50;'>📖 관련 법령 상세조회</h4>", unsafe_allow_html=True)
    
    if st.button("🔙 이전 질의회신으로 돌아가기", use_container_width=True):
        st.session_state.view_mode = 'main'
        st.rerun()
    
    law = st.session_state.view_law_data
    st.markdown(f"### ⚖️ {law['조문']}")
    st.markdown(f"**📜 [법률]**<br>{render_safe_html(law['법률'])}", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown(f"**⚙️ [시행령]**<br>{render_safe_html(law['시행령'])}", unsafe_allow_html=True)
    if law['시행규칙'].strip():
        st.markdown("---")
        st.markdown(f"**📋 [시행규칙]**<br>{render_safe_html(law['시행규칙'])}", unsafe_allow_html=True)
    
    st.markdown("---")
    if st.button("🔙 목록으로 돌아가기", key="btn_bottom_back", use_container_width=True):
        st.session_state.view_mode = 'main'
        st.rerun()
        
    st.stop()

# ==========================================
# [4. 최상단 배너 및 메인 타이틀 (홈버튼 + QR)]
# ==========================================

# 🔥 마법의 CSS: 스마트폰에서는 QR코드 숨기기, PC에서는 예쁘게 띄우기
st.markdown("""
<style>
button[kind="primary"] {
    background-color: transparent !important;
    border: none !important;
    box-shadow: none !important;
    padding: 0 !important;
}
button[kind="primary"] p {
    font-size: 1.3rem !important;
    font-weight: bold !important;
    color: #2c3e50 !important;
    margin: 15px 0px !important;
    text-align: center !important; /* 다시 예전처럼 정중앙 정렬! */
}
button[kind="primary"]:hover p {
    color: #3498db !important;
}

/* QR코드 박스 설정 */
.qr-container {
    position: relative;
    width: 100%;
    height: 0px;
    top: -55px; /* 타이틀 버튼 위로 겹쳐 올리기 */
    display: flex;
    justify-content: flex-end;
    z-index: 10;
    pointer-events: none; /* 버튼 클릭을 방해하지 않도록 설정 */
}
.qr-container img {
    width: 45px; 
    height: 45px; 
    border-radius: 5px; 
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    pointer-events: auto; /* QR코드 이미지는 보이게 유지 */
}

/* 🚨 스마트폰 화면(폭 768px 이하)에서는 QR코드를 완전히 숨깁니다! */
@media (max-width: 768px) {
    .qr-container {
        display: none !important;
    }
}
</style>
""", unsafe_allow_html=True)

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
    if st.button(f"🔔 중요 예정 업무 알림 [{first['region']}]: {d_text} {first['memo']}", use_container_width=True):
        st.session_state.active_tab = "📅 공유달력"
        st.session_state.view_mode = 'main'
        st.rerun()

if notices:
    st.info(f"📢 **[전체 공지사항]** {notices[-1]['내용']}")

# 타이틀(홈버튼) 출력
if st.button("🔍 지적재조사 통합 검색", type="primary", use_container_width=True):
    st.session_state.view_mode = 'main'
    st.session_state.active_tab = "📑 질의회신" 
    st.rerun()

# ==========================================
# [파이썬 자체 내장 QR 이미지 생성 및 출력]
# ==========================================
target_url = "https://system-ydyhcgqqhe6dncgekqklcv.streamlit.app"

# 파이썬이 직접 QR코드 그림을 그립니다 (외부망 접속 X)
qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
qr.add_data(target_url)
qr.make(fit=True)
img = qr.make_image(fill_color="black", back_color="white")

# 이미지를 글자(base64)로 변환하여 메모리에 임시 저장
buf = BytesIO()
img.save(buf, format="PNG")
b64_img = base64.b64encode(buf.getvalue()).decode()

# 주무관님이 만드신 예쁜 CSS 컨테이너 안에 변환된 이미지를 쏙 넣습니다
st.markdown(
    f'''
    <div class="qr-container">
        <img src="data:image/png;base64,{b64_img}">
    </div>
    ''', unsafe_allow_html=True
)

# ==========================================
# [5. 검색 UI 및 탭 설정]
# ==========================================
col1, col2 = st.columns([3, 1])
with col1: keyword = st.text_input("검색어를 입력하세요", label_visibility="collapsed", placeholder="🔍 검색어 입력 (예: 경계설정)")
with col2: search_btn = st.button("검색", use_container_width=True)

only_title = st.checkbox("☑️ 제목만 검색", value=True)

tabs = ["📑 질의회신", "⚖️ 법령", "🏢 업무규정", "📐 측량규정", "🏢 판례", "📅 공유달력"]
mode = st.radio("자료 선택", tabs, horizontal=True, label_visibility="collapsed", key="active_tab")
st.markdown("---")

# ==========================================
# [6. 카테고리별 출력 및 일정 관리]
# ==========================================
if mode in ["📑 질의회신", "🏢 판례"]:
    target_df = df_qna if mode == "📑 질의회신" else df_case
    if keyword:
        res = target_df[target_df['제목'].str.contains(keyword, case=False, na=False)] if only_title else target_df[target_df['제목'].str.contains(keyword, case=False, na=False) | target_df['내용'].str.contains(keyword, case=False, na=False)]
    else: res = target_df 
        
    st.caption(f"총 {len(res)}건의 자료가 있습니다.")
    for idx, row in res.iterrows(): 
        icon = "🟢" if str(row.get("수정여부")).strip().upper() == "Y" else "📑"
        with st.expander(f"{icon} {row['제목']}"):
            st.markdown(render_safe_html(row['내용'], keyword), unsafe_allow_html=True)
            
            if mode == "📑 질의회신":
                raw_jos = re.findall(r'제\s*\d+\s*조(?:의\s*\d+)?', row['내용'])
                normalized_jos = set([re.sub(r'\s+', '', jo) for jo in raw_jos])
                
                matched_laws = []
                for law in law_db:
                    base_jo = law['조문'].split('(')[0].replace(" ", "")
                    if base_jo in normalized_jos:
                        matched_laws.append(law)
                
                if matched_laws:
                    st.markdown("---")
                    st.markdown("🔗 **관련 법령 바로보기**")
                    cols = st.columns(min(len(matched_laws), 3)) 
                    for i, law in enumerate(matched_laws):
                        with cols[i % 3]:
                            if st.button(f"📖 {law['조문'].split('(')[0]}", key=f"btn_law_{idx}_{i}"):
                                st.session_state.view_mode = 'law_detail'
                                st.session_state.view_law_data = law
                                st.rerun()

elif mode == "⚖️ 법령":
    if keyword:
        count = 0
        for item in law_db:
            match = (keyword in item['조문']) if only_title else (keyword in item['조문'] or keyword in item['법률'] or keyword in item['시행령'] or keyword in item['시행규칙'])
            if match:
                count += 1
                with st.expander(f"⚖️ [법령] {item['조문']}"):
                    st.markdown(f"**📜 [법률]**<br>{render_safe_html(item['법률'], keyword)}", unsafe_allow_html=True)
                    st.markdown("---")
                    st.markdown(f"**⚙️ [시행령]**<br>{render_safe_html(item['시행령'], keyword)}", unsafe_allow_html=True)
                    if item['시행규칙'].strip():
                        st.markdown("---")
                        st.markdown(f"**📋 [시행규칙]**<br>{render_safe_html(item['시행규칙'], keyword)}", unsafe_allow_html=True)
        st.caption(f"총 {count}건의 조문이 검색되었습니다.")
    else:
        st.caption(f"총 {len(law_db)}개의 전체 조문입니다.")
        for item in law_db:
            with st.expander(f"⚖️ [법령] {item['조문']}"):
                st.markdown(f"**📜 [법률]**<br>{render_safe_html(item['법률'])}", unsafe_allow_html=True)
                st.markdown("---")
                st.markdown(f"**⚙️ [시행령]**<br>{render_safe_html(item['시행령'])}", unsafe_allow_html=True)
                if item['시행규칙'].strip():
                    st.markdown("---")
                    st.markdown(f"**📋 [시행규칙]**<br>{render_safe_html(item['시행규칙'])}", unsafe_allow_html=True)

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
                        st.markdown(render_safe_html(item['내용'], keyword), unsafe_allow_html=True)
    st.caption(f"총 {count}건의 목록이 있습니다.")

elif mode == "📅 공유달력":
    st.subheader("🔐 지역별 보안 공유 달력")
    regions = ["포항시", "경주시", "김천시", "안동시", "구미시", "영주시", "영천시", "상주시", "문경시", "경산시", "의성군", "청송군", "영양군", "영덕군", "청도군", "고령군", "성주군", "칠곡군", "예천군", "봉화군", "울진군", "울릉군", "경상북도(총괄)"]
    
    default_idx = regions.index(st.session_state.saved_region) if st.session_state.saved_region in regions else 0
    selected_region = st.selectbox("📌 담당 시/군을 선택하세요", regions, index=default_idx)
    
    if selected_region != st.session_state.saved_region:
        st.session_state.saved_region = selected_region
    
    col_pw, col_btn = st.columns([3, 1])
    with col_pw: entered_pw = st.text_input("🔑 비밀번호 4자리를 입력하세요", type="password")
    with col_btn:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        login_btn = st.button("확인", use_container_width=True)
    
    is_unlocked = False
    if entered_pw or login_btn:
        if entered_pw:
            try:
                if entered_pw == st.secrets["passwords"][selected_region]: is_unlocked = True
                else: st.error("❌ 비밀번호가 일치하지 않습니다.")
            except KeyError: st.warning("⚠️ 이 지역의 비밀번호가 설정되지 않았습니다.")
        elif login_btn: st.warning("비밀번호를 입력해주세요.")
            
    if is_unlocked:
        st.success(f"🔓 [{selected_region}] 전용 달력에 접속되었습니다!")
        st.markdown("---")
        
        with st.form("google_calendar_form"):
            st.write(f"**[{selected_region}] 새로운 일정 등록**")
            e_date = st.date_input("날짜 선택")
            e_memo = st.text_area("일정 메모")
            e_alarm = st.checkbox("🔔 상단 D-Day 알람 켜기")
            e_days = st.selectbox("알람 기간 (며칠 전부터 알릴까요?)", [0, 1, 3, 5, 7, 10, 30], index=1)
            
            if st.form_submit_button("일정 저장 및 동기화"):
                if e_memo.strip():
                    with st.spinner("저장 중..."): save_event_to_google(e_date.strftime("%Y-%m-%d"), e_memo, e_alarm, e_days, selected_region)
                    st.success("✅ 일정이 추가되었습니다!")
                    st.rerun()
                else: st.warning("메모를 입력해주세요.")

        st.markdown("---")

        if selected_region == "경상북도(총괄)":
            with st.expander("📢 관리자용: 팀 전체 공지사항 관리"):
                current_notice_text = notices[-1]['내용'] if notices else ""
                new_notice = st.text_area("앱 최상단에 띄울 공지 내용을 입력/수정하세요", value=current_notice_text)
                
                col_n1, col_n2 = st.columns(2)
                with col_n1:
                    if st.button("📝 공지 저장/수정", use_container_width=True):
                        if new_notice.strip():
                            with st.spinner("저장 중..."): save_notice(new_notice)
                            st.success("공지가 성공적으로 업데이트되었습니다!")
                            st.rerun()
                        else:
                            st.warning("내용을 입력해주세요.")
                with col_n2:
                    if st.button("🗑️ 공지 삭제 (배너 숨기기)", use_container_width=True):
                        with st.spinner("삭제 중..."): delete_notice()
                        st.success("공지가 완전히 삭제되었습니다!")
                        st.rerun()
                        
            st.subheader("📋 전체 일정 목록")
            if all_events:
                all_events.sort(key=lambda x: x["date"])
                for info in all_events:
                    alarm_icon = "🔔" if info.get("use_alarm") else "📌"
                    with st.expander(f"{alarm_icon} [{info['region']}] {info['date']} | {info['memo'][:15]}..."):
                        st.write(f"**🏢 지역:** {info['region']} / **📅 날짜:** {info['date']}\n\n**📝 내용:** {info['memo']}")
                        col_share, col_del = st.columns(2)
                        with col_share: native_share_button(info['region'], info['date'], info['memo'])
                        with col_del:
                            if st.button("🗑️ 일정 삭제", key=f"del_admin_{info['row_idx']}", use_container_width=True):
                                delete_event_from_google(info['row_idx'])
                                st.rerun()
        else:
            st.subheader(f"📋 [{selected_region}] 일정 목록")
            region_events = [e for e in all_events if e["region"] == selected_region]
            if region_events:
                region_events.sort(key=lambda x: x["date"])
                for info in region_events:
                    alarm_icon = "🔔" if info.get("use_alarm") else "📌"
                    with st.expander(f"{alarm_icon} [{info['date']}] {info['memo'][:20]}..."):
                        st.write(f"**📅 날짜:** {info['date']}\n\n**📝 내용:** {info['memo']}")
                        col_share, col_del = st.columns(2)
                        with col_share: native_share_button(info['region'], info['date'], info['memo'])
                        with col_del:
                            if st.button("🗑️ 일정 삭제", key=f"del_{info['row_idx']}", use_container_width=True):
                                delete_event_from_google(info['row_idx'])
                                st.rerun()

st.markdown("---")
st.caption("v8.8 Perfect - 스마트 기기 감지형 반응형 UI (모바일 QR 자동 숨김) 탑재")
