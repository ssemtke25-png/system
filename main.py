import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import json
import re
import time
import base64
import gspread
import html
import qrcode
import threading
from io import BytesIO
from bs4 import BeautifulSoup
from datetime import datetime
# 🤖 AI 모델: Gemini·Gemma(google-genai) / Solar(openai 호환) — ask_ai 안에서 불러옴

# ==========================================
# [1. 웹 페이지 기본 설정 및 구글 시트 연동]
# ==========================================
st.set_page_config(page_title="물어보이소, AI — 지적·지적재조사 업무지원", page_icon="🔍", layout="wide")

# ==========================================================================
# 🛡️ [v21] 브라우저 번역기 DOM 충돌 차단 (removeChild 에러 방지)
# ==========================================================================
components.html("""
<script>
(function () {
  try {
    var doc = window.parent.document;
    var htmlEl = doc.documentElement;
    htmlEl.setAttribute('translate', 'no');
    if (!htmlEl.classList.contains('notranslate')) htmlEl.classList.add('notranslate');
    htmlEl.setAttribute('lang', 'ko');
    if (!doc.querySelector('meta[name="google"]')) {
      var m = doc.createElement('meta');
      m.name = 'google';
      m.content = 'notranslate';
      doc.head.appendChild(m);
    }
  } catch (e) {}
})();
</script>
""", height=0)
# ==========================================================================

# ==========================================================================
# 👥 [v23] 동시접속자 실시간 카운터 (구글 시트 미사용 / 앱 메모리 방식)
# --------------------------------------------------------------------------
#  원리:
#   - 접속 세션마다 고유 ID를 부여하고, 마지막 활동 시각을 앱 공용 메모리에 기록
#   - "최근 30초 안에 활동한 세션 수" = 현재 동시접속자 수
#   - @st.cache_resource 는 앱 인스턴스 전체가 공유하는 단일 객체 → 세션 간 공유 가능
#   - threading.Lock 으로 동시 쓰기 충돌 방지
#  주의:
#   - 앱이 재시작되면 초기화됨(메모리 방식). 발표 몇 시간은 문제 없음.
#   - 시트/DB 를 안 건드리므로 API 한도에 영향 없음.
# ==========================================================================
PRESENCE_WINDOW_SEC = 30      # 이 시간 안에 활동한 세션만 '접속중'으로 집계
WARN_LEVEL = 50               # 이 수를 넘으면 노랑(주의)
DANGER_LEVEL = 80             # 이 수를 넘으면 빨강(위험)


@st.cache_resource
def _get_presence_store():
    """앱 전체가 공유하는 접속 기록 저장소 (session_id -> last_seen_epoch)."""
    return {"lock": threading.Lock(), "seen": {}}


def _touch_presence():
    """현재 세션의 활동 시각을 갱신하고, 최근 활동 세션 수를 반환."""
    store = _get_presence_store()
    # 세션마다 1회 고유 ID 부여
    if "presence_id" not in st.session_state:
        st.session_state["presence_id"] = f"{time.time()}_{id(st.session_state)}"
    sid = st.session_state["presence_id"]
    now = time.time()
    with store["lock"]:
        seen = store["seen"]
        seen[sid] = now
        # 오래된(윈도우 밖) 세션 정리
        cutoff = now - PRESENCE_WINDOW_SEC
        stale = [k for k, v in seen.items() if v < cutoff]
        for k in stale:
            del seen[k]
        return len(seen)


def render_presence_badge(count):
    """동시접속자 수를 색상(초록/노랑/빨강)으로 크게 표시."""
    if count >= DANGER_LEVEL:
        color, bg, label = "#c0392b", "#fdecea", "⚠️ 접속 폭주 (서버 부하 위험)"
    elif count >= WARN_LEVEL:
        color, bg, label = "#e67e22", "#fef5e7", "🔶 접속 많음 (주의)"
    else:
        color, bg, label = "#27ae60", "#eafaf1", "🟢 안정적으로 접속 중"

    st.markdown(
        f"""
        <div translate="no" class="notranslate" style="
            display:flex; align-items:center; justify-content:center; gap:14px;
            background:{bg}; border:1px solid {color}33; border-radius:10px;
            padding:10px 16px; margin-bottom:10px;">
            <span style="font-size:0.95rem; color:#555; font-weight:600;">👥 현재 동시 접속자</span>
            <span style="font-size:2.0rem; font-weight:800; color:{color}; line-height:1;">{count}</span>
            <span style="font-size:0.9rem; color:{color}; font-weight:700;">{label}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
# ==========================================================================

DATA_DIR = "data"
EXCEL_PATH = f"{DATA_DIR}/data.xlsx"
LAW_HTML_PATH = f"{DATA_DIR}/지적재조사에 관한 특별법(인용조문 3단비교).html"
REG_DIR = f"{DATA_DIR}/규정"
# 📐 [지적 분야] 데이터 (지적업무사례집·토지이동업무편람·지적판례·지적규정 + 공간정보관리법 3단비교)
JJ_EXCEL_PATH = f"{DATA_DIR}/02_지적_자료.xlsx"
JJ_LAW_HTML_PATH = f"{DATA_DIR}/공간정보의 구축 및 관리 등에 관한 법률(인용조문 3단비교).html"
JJ_SHEETS = [("지적업무사례집", "📘"), ("토지이동업무편람", "📗"), ("지적판례", "⚖️"), ("지적규정", "📕")]
DOMAINS = ["🏗️ 지적재조사", "📐 지적"]
MAX_LIST = 200   # 검색어 없이 전체 목록을 펼칠 때 최대 표시 건수(화면 속도용)

@st.cache_resource
def get_google_sheet():
    try:
        raw_json = st.secrets["google_json"].replace('\xa0', ' ').replace('\u00A0', ' ')
        secret_json = json.loads(raw_json)
        gc = gspread.service_account_from_dict(secret_json)
        clean_url = st.secrets["spreadsheet_url"].strip()
        return gc.open_by_url(clean_url)
    except Exception as e:
        return None

sh = get_google_sheet()
if sh is None:
    st.error("🚨 구글 시트 연결 에러! (Secrets 설정을 확인하세요)")

sheet_main = sh.get_worksheet(0) if sh else None
sheet_notice = sh.get_worksheet(1) if sh else None


@st.cache_data(ttl=60, show_spinner=False)
def _cached_notice_records(_sheet):
    try: return _sheet.get_all_records() if _sheet else []
    except: return []

@st.cache_data(ttl=60, show_spinner=False)
def _cached_event_records(_sheet):
    try: return _sheet.get_all_records() if _sheet else []
    except: return []

def clear_data_cache():
    _cached_notice_records.clear()
    _cached_event_records.clear()


def load_notice():
    return _cached_notice_records(sheet_notice)

def save_notice(content):
    if sheet_notice:
        sheet_notice.clear()
        sheet_notice.append_row(["날짜", "내용"])
        sheet_notice.append_row([datetime.now().strftime("%Y-%m-%d"), content])
        clear_data_cache()

def delete_notice():
    if sheet_notice:
        try:
            sheet_notice.clear()
            sheet_notice.append_row(["날짜", "내용"])
            clear_data_cache()
        except: pass

def load_events_from_google():
    if sheet_main is None: return []
    try:
        records = _cached_event_records(sheet_main)
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
        clear_data_cache()

def delete_event_from_google(row_idx):
    if sheet_main:
        sheet_main.delete_rows(row_idx)
        clear_data_cache()

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
def parse_law_html(path):
    """법제처 '인용조문 3단비교' html → [{조문, 법률, 시행령, 시행규칙}]"""
    law_db = []
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                soup = BeautifulSoup(f, "html.parser")
            for tr in soup.select("tr"):
                tds = tr.select("td")
                if len(tds) >= 3 and (jo_span := tds[0].select_one("span.bl")):
                    law_db.append({"조문": jo_span.get_text(strip=True), "법률": tds[0].get_text("\n", strip=True), "시행령": tds[1].get_text("\n", strip=True), "시행규칙": tds[2].get_text("\n", strip=True)})
        except: pass
    return law_db


@st.cache_data(ttl=600)
def load_jijeok_data():
    """📐 지적 분야: 시트별 DataFrame dict + 공간정보관리법 3단비교"""
    sheets = {}
    for name, _icon in JJ_SHEETS:
        try:
            df = pd.read_excel(JJ_EXCEL_PATH, sheet_name=name).fillna("").astype(str)
        except Exception:
            df = pd.DataFrame(columns=["제목", "내용"])
        if "수정여부" not in df.columns:
            df["수정여부"] = ""
        sheets[name] = df
    return sheets, parse_law_html(JJ_LAW_HTML_PATH)


@st.cache_data(ttl=600)
def load_all_data_final_v8():
    def clean_reg_text(text):
        match = re.search(r'(【질의회신|\[질의회신|【참고판례|\[참고판례|\[질의요지\])', text)
        if match: return text[:match.start()].strip()
        return text.strip()

    law_db = parse_law_html(LAW_HTML_PATH)

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
# 🤖 [AI 검색] 조문번호 정규화 + 인덱스 캐싱 함수
# ==========================================
def _normalize_jo(text):
    def repl(m):
        jo = m.group(1)
        ui = m.group(2)
        return f"제{jo}조의{ui}" if ui else f"제{jo}조"
    return re.sub(r'제?\s*(\d+)\s*조(?:\s*의\s*(\d+))?', repl, str(text))


def _strip_spaces(text):
    return re.sub(r'\s+', '', str(text))


_JOSA = sorted(set([
    "이라는", "라는", "이라고", "라고", "으로서", "으로써", "이라면", "라면",
    "이란", "란", "에서", "에게", "으로", "로서", "로써", "까지", "부터",
    "마다", "조차", "에는", "에도", "이나", "은", "는", "이", "가", "을",
    "를", "의", "에", "로", "도", "만", "뿐", "와", "과",
]), key=len, reverse=True)


def _strip_josa(word):
    if len(word) <= 2:
        return word
    for j in _JOSA:
        if word.endswith(j) and len(word) - len(j) >= 2:
            return word[:-len(j)]
    return word


def _extract_query_jos(question):
    jos = set()
    for m in re.finditer(r'제?\s*(\d+)\s*조(?:\s*의\s*(\d+))?', question):
        jo, ui = m.group(1), m.group(2)
        jos.add(f"제{jo}조의{ui}" if ui else f"제{jo}조")
    return jos


def _extract_keywords(question):
    stopwords = {"어떻게", "무엇", "뭐", "인가요", "인가", "하나요", "되나요", "될까요",
                 "있나요", "있는", "있을", "경우", "관련", "대한", "대해", "그리고",
                 "또는", "해야", "하는", "합니까", "됩니까", "가능", "여부", "알려줘",
                 "알려주세요", "설명", "질문", "궁금", "무슨", "어떤", "이런", "저런",
                 "하는지", "되는지", "지정은", "위한", "위해", "규정", "조문", "조항"}
    q = re.sub(r'제?\s*\d+\s*조(?:\s*의\s*\d+)?', ' ', question)
    words = re.findall(r'[가-힣A-Za-z0-9]{2,}', q)
    return [w for w in words if w not in stopwords]


@st.cache_data(ttl=600)
def build_search_index(_df_qna, _df_case, _law_db, _reg_db):
    index = []

    def add(kind, title, content):
        title, content = str(title), str(content)
        norm_t = _normalize_jo(title)
        norm_c = _normalize_jo(content)
        nosp_t = _strip_spaces(norm_t)
        nosp_c = _strip_spaces(norm_c)
        index.append((kind, title, content, norm_t, norm_c, nosp_t, nosp_c))

    for _, row in _df_qna.iterrows():
        add("질의회신", row.get("제목", ""), row.get("내용", ""))
    for _, row in _df_case.iterrows():
        add("판례", row.get("제목", ""), row.get("내용", ""))
    for item in _law_db:
        c = f"[법률]\n{item.get('법률','')}\n[시행령]\n{item.get('시행령','')}\n[시행규칙]\n{item.get('시행규칙','')}"
        add("법령", item.get("조문", ""), c)
    for reg_name, reg_data in _reg_db.items():
        for item in reg_data:
            add("규정", f"{reg_name} {item.get('조문','')}", item.get("내용", ""))
    return index


@st.cache_data(ttl=600)
def build_search_index_jj(_sheets, _law_db):
    """📐 지적 분야 AI 검색용 인덱스 (시트 4종 + 공간정보관리법)"""
    index = []
    def add(kind, title, content):
        title, content = str(title), str(content)
        norm_t, norm_c = _normalize_jo(title), _normalize_jo(content)
        index.append((kind, title, content, norm_t, norm_c, _strip_spaces(norm_t), _strip_spaces(norm_c)))
    for name, df in _sheets.items():
        for _, row in df.iterrows():
            add(name, row.get("제목", ""), row.get("내용", ""))
    for item in _law_db:
        c = f"[법률]\n{item.get('법률','')}\n[시행령]\n{item.get('시행령','')}\n[시행규칙]\n{item.get('시행규칙','')}"
        add("법령", item.get("조문", ""), c)
    return index


def find_relevant_materials(question, df_qna, df_case, law_db, reg_db, max_items=8):
    index = build_search_index(df_qna, df_case, law_db, reg_db)
    return score_materials(index, question, max_items)


def find_relevant_materials_jj(question, sheets, law_db, max_items=8):
    return score_materials(build_search_index_jj(sheets, law_db), question, max_items, capped=True)


def score_materials(index, question, max_items=8, capped=False):
    """capped=True(지적 분야): 긴 판례가 단어 횟수만으로 위로 올라오지 않도록 횟수 상한 + 제목 가중"""
    keywords = _extract_keywords(question)
    query_jos = _extract_query_jos(question)

    if not keywords and not query_jos:
        keywords = [question]

    scored = []
    for kind, title, content, norm_t, norm_c, nosp_t, nosp_c in index:
        s = 0
        matched_kw = 0
        for jo in query_jos:
            if jo in norm_t:
                s += 40 if capped else 20
            if jo in norm_c:
                s += 8
        for kw in keywords:
            t_hit = title.count(kw)
            c_hit = content.count(kw)
            kw_key = _strip_spaces(_strip_josa(kw))
            if kw_key and len(kw_key) >= 2:
                t_hit = max(t_hit, nosp_t.count(kw_key))
                c_hit = max(c_hit, nosp_c.count(kw_key))
            if capped:
                s += min(t_hit, 2) * 6 + min(c_hit, 3)
            else:
                s += t_hit * 3
                s += c_hit * 1
            if t_hit or c_hit:
                matched_kw += 1
        if matched_kw >= 2:
            s += matched_kw * (10 if capped else 8)
        if s > 0 and kind in ("법령", "규정", "지적규정"):
            s += 3
        if s > 0:
            scored.append((s, kind, title, content))

    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[:max_items]


class AIBusyError(Exception):
    pass


AI_MODELS = {
    "Gemini 2.5 Flash (빠름)": "gemini",
    "Gemma 4 (공통기반 지원모델)": "gemma",
    "Solar (국산 · Upstage)": "solar",
}


def _secret(*keys):
    for k in keys:
        try:
            v = st.secrets.get(k, "")
        except Exception:
            v = ""
        if v:
            return str(v).strip()
    return ""


def _model_picker(key):
    """AI 질문 탭에서 쓰는 모델 선택 (지적재조사·지적 공통으로 선택값 유지)"""
    labels = list(AI_MODELS.keys())
    cur = st.session_state.get("ai_model_label", labels[0])
    idx = labels.index(cur) if cur in labels else 0
    lab = st.selectbox("AI 모델", labels, index=idx, key=key)
    st.session_state.ai_model_label = lab
    return AI_MODELS[lab]


def _is_transient(e):
    msg = str(e).lower()
    return any(k in msg for k in
               ["429", "503", "quota", "rate", "overload", "unavailable", "resource_exhausted",
                "resource has been exhausted"])


def ask_ai(question, materials, max_retries=2, field="지적재조사", model_key="gemini", history=None):
    context_parts = []
    for i, (sc, kind, title, content) in enumerate(materials, 1):
        context_parts.append(f"[자료{i}] ({kind}) {title}\n{content[:2200]}")
    context = "\n\n---\n\n".join(context_parts)

    system_instruction = (
        f"당신은 {field} 업무를 지원하는 공공행정 전문 AI입니다.\n"
        "\n"
        "[가장 중요한 원칙 - 반드시 지킬 것]\n"
        "1. 아래 '참고자료'에 실제로 있는 내용만 근거로 답변하세요.\n"
        "2. 자료에 없는 사실·수치·조항 번호·날짜·금액은 절대로 추가하거나 지어내지 마세요. "
        "만약 자료에서 확인되지 않으면 '제공된 자료에서는 확인되지 않습니다'라고 명확히 밝히세요.\n"
        "3. 답변에 사용한 근거는 반드시 [자료1], [자료2] 형태로 어느 자료에서 나왔는지 표시하세요.\n"
        "\n"
        "[답변 방식 - 위 원칙을 지키는 선에서 최대한 자세히]\n"
        "- 실무자가 이 답변만 보고도 업무를 처리할 수 있도록, 자료의 내용을 충분히 풀어서 자세하고 친절하게 설명하세요.\n"
        "- 조문이라면 각 항(①②③)의 내용을 하나씩 구체적으로 설명하고, 핵심 요건·절차·예외가 있으면 빠짐없이 짚어주세요.\n"
        "- 질문과 관련된 여러 자료(법령+질의회신+판례 등)가 있으면 서로 연결해서, 실무에서 어떤 의미인지까지 설명하세요.\n"
        "- 다만 이렇게 '풀어서 설명'하는 것과 '없는 내용을 지어내는 것'은 다릅니다. "
        " 설명을 길게 하더라도 그 근거는 반드시 제공된 자료 안에 있어야 하며, 자료에 없는 배경지식이나 일반론을 사실인 것처럼 덧붙이지 마세요.\n"
        "- 필요하면 항목·번호를 나눠 보기 좋게 정리하고, 마지막에 근거 조문·질의회신 출처를 함께 제시하세요.\n"
        "- 답변이 너무 길어질 것 같으면 핵심 위주로 정리하되, 반드시 문장과 설명을 끝까지 완결지어 마무리하세요. 중간에 문장이 끊기지 않도록 하세요."
    )
    system_instruction += (
        "\n- 질문이 모호해서 자료상 경우(사업 종류·지목·절차 단계 등)에 따라 답이 달라지면, 경우를 나눠 답하고 "
        "마지막에 '어느 경우에 해당하는지 알려주시면 더 정확히 안내드리겠습니다'처럼 짧게 되물어 주세요."
    )
    hist_txt = ""
    if history:
        system_instruction += (
            "\n\n[이어지는 질문 처리]\n"
            "- '이전 대화'는 '그럼', '그 경우', '그건' 같은 말이 무엇을 가리키는지 이해하는 데만 사용하세요.\n"
            "- 사실·조문·수치의 근거는 반드시 이번 '참고자료'에서만 찾고 [자료N]으로 표시하세요. "
            "이전 답변의 내용이라도 이번 참고자료에서 확인되지 않으면 단정하지 마세요."
        )
        hist_txt = "[이전 대화]\n" + "\n\n".join(
            f"질문{i}: {hq}\n답변{i}: {ha[:1500]}" for i, (hq, ha) in enumerate(history, 1)
        ) + "\n\n---\n\n"
    user_prompt = f"{hist_txt}[참고자료]\n{context}\n\n---\n\n[질문]\n{question}"

    if model_key in ("gemini", "gemma"):
        gkey = _secret("GEMINI_API_KEY", "gemini_api_key")
        if not gkey:
            raise ValueError("GEMINI_API_KEY가 Secrets에 없습니다.")
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=gkey)
        model_name = "gemini-2.5-flash" if model_key == "gemini" else "gemma-4-31b-it"
        if model_key == "gemma":
            # Gemma는 system_instruction 미지원 → 지시문을 본문 앞에 붙임
            cfg = types.GenerateContentConfig(temperature=0.2, max_output_tokens=4096)
            contents = system_instruction + "\n\n" + user_prompt
        else:
            cfg = types.GenerateContentConfig(system_instruction=system_instruction,
                                              temperature=0.2, max_output_tokens=4096)
            contents = user_prompt
        call = lambda: client.models.generate_content(model=model_name, contents=contents, config=cfg).text
    elif model_key == "solar":
        ukey = _secret("UPSTAGE_API_KEY", "upstage_api_key")
        if not ukey:
            raise ValueError("UPSTAGE_API_KEY가 Secrets에 없습니다.")
        from openai import OpenAI
        client = OpenAI(api_key=ukey, base_url="https://api.upstage.ai/v1")
        call = lambda: client.chat.completions.create(
            model="solar-pro2",
            messages=[{"role": "system", "content": system_instruction},
                      {"role": "user", "content": user_prompt}],
            temperature=0.2, max_tokens=4096).choices[0].message.content
    else:
        raise ValueError(model_key)

    last_err = None
    for attempt in range(max_retries + 1):
        try:
            return call()
        except Exception as e:
            last_err = e
            if _is_transient(e) and attempt < max_retries:
                time.sleep(1.5 * (attempt + 1))
                continue
            if _is_transient(e):
                raise AIBusyError(str(e))
            raise
    raise AIBusyError(str(last_err) if last_err else "unknown")

MAX_TURNS = 5        # 한 대화에서 이어서 묻기 최대 횟수(첫 질문 포함)
HIST_TURNS = 3       # AI에게 함께 보내는 최근 대화 수


def _merge_materials(primary, secondary, max_items=7):
    """두 검색결과를 번갈아 합침(앞쪽 우선, 제목 중복 제거)"""
    out, seen = [], set()
    for a, b in __import__("itertools").zip_longest(primary, secondary):
        for m in (a, b):
            if m and m[2] not in seen:
                seen.add(m[2]); out.append(m)
    return out[:max_items]


def _ai_status_msg(stt):
    if stt["status"] == "empty":
        st.warning("질문을 입력해주세요.")
    elif stt["status"] == "no_material":
        st.info("질문과 관련된 자료를 찾지 못했습니다. 다른 키워드로 다시 질문해보세요.")
    elif stt["status"] == "busy":
        st.warning("⏳ 지금 접속자가 많아 잠시 응답이 지연되고 있습니다.\n\n10~20초 후 다시 눌러주세요.")
    elif stt["status"] == "error":
        st.error(f"AI 답변 생성 중 오류가 발생했습니다: {stt['msg']}")
        st.info("Secrets에 GEMINI_API_KEY(Gemini·Gemma) / UPSTAGE_API_KEY(Solar)가 설정되었는지 확인하거나, 다른 모델로 바꿔 다시 질문해보세요.")


def _run_turn(dk, q, finder, field, model_key):
    """질문 1회 처리 → 성공 시 대화에 추가, 실패 시 상태 반환"""
    chat = st.session_state[f"chat_{dk}"]
    mats = finder(q)
    if chat:   # 이어서 묻기 : "관련 판례도?"처럼 짧은 질문은 앞 질문과 합친 검색을 우선
        ctx_q = " ".join(t["q"] for t in chat[-2:]) + " " + q
        mats = _merge_materials(finder(ctx_q), mats)
    if not mats:
        return {"status": "no_material"}
    hist = [(t["q"], t["answer"]) for t in chat[-HIST_TURNS:]]
    try:
        ans = ask_ai(q, mats, field=field, model_key=model_key, history=hist or None)
    except AIBusyError:
        return {"status": "busy"}
    except Exception as e:
        return {"status": "error", "msg": str(e)}
    chat.append({"q": q, "answer": ans, "model": st.session_state.get("ai_model_label", ""),
                 "materials": [(k, t, c) for (sc, k, t, c) in mats]})
    return None


def render_ai_chat(dk, title, placeholder, caption, finder, field):
    """🤖 AI 질문 화면(지적재조사·지적 공통) — 첫 질문 + 이어서 묻기"""
    ck, sk = f"chat_{dk}", f"ai_state_{dk}"
    st.session_state.setdefault(ck, [])
    st.session_state.setdefault(sk, None)

    st.markdown(f"**🤖 질문을 문장으로 입력하세요 ({title})**")
    q = st.text_area("질문을 문장으로 입력하세요", height=80, key=f"ai_question_input_{dk}",
                     label_visibility="collapsed", placeholder=placeholder)
    mk = _model_picker(f"ai_model_sel_{dk}")
    ask = st.button("🤖 AI에게 질문하기", type="secondary", use_container_width=True, key=f"ai_ask_btn_{dk}")
    st.caption(caption)
    _scroll_to_question_on_mobile()

    if ask:
        st.session_state[ck] = []          # 새 질문 = 새 대화
        st.session_state[sk] = None
        if not q.strip():
            st.session_state[sk] = {"status": "empty"}
        else:
            with st.spinner("관련 자료를 찾고 답변을 작성 중입니다..."):
                st.session_state[sk] = _run_turn(dk, q.strip(), finder, field, mk)

    chat = st.session_state[ck]
    if not chat:
        if st.session_state[sk]:
            _ai_status_msg(st.session_state[sk])
        else:
            st.info("궁금한 내용을 문장으로 입력하고 [AI에게 질문하기]를 눌러주세요. 답변 아래에서 이어서 물어볼 수 있습니다.")
        return

    for n, t in enumerate(chat):
        last = (n == len(chat) - 1)
        if n == 0:
            st.markdown("### 💡 AI 답변")
        else:
            st.markdown("---")
            st.markdown(f"#### ↪️ 이어서 질문 {n} : {html.escape(t['q'])}")
        if t.get("model"):
            st.caption(f"답변 모델 : {t['model']}")
        st.markdown(render_ai_answer_html(t["answer"]), unsafe_allow_html=True)
        if last:
            st.markdown("#### 📚 답변 근거 자료")
            st.caption("AI가 참고한 실제 원문입니다. 반드시 아래 원문으로 확인하세요.")
            for i, (k, ti, c) in enumerate(t["materials"], 1):
                with st.expander(f"[자료{i}] ({k}) {ti}"):
                    st.markdown(render_safe_html(c), unsafe_allow_html=True)
        else:
            with st.expander(f"📚 이 답변의 근거 자료 {len(t['materials'])}건"):
                for i, (k, ti, c) in enumerate(t["materials"], 1):
                    st.markdown(f"**[자료{i}] ({k}) {html.escape(ti)}**")
                    st.markdown(render_safe_html(c), unsafe_allow_html=True)

    st.markdown("---")
    if len(chat) >= MAX_TURNS:
        st.info(f"한 대화에서는 {MAX_TURNS}번까지 이어서 물을 수 있습니다. 위 질문칸에서 새 질문으로 시작해 주세요.")
        return
    st.markdown("**↪️ 이어서 묻기** — 위 답변에 이어서 궁금한 점을 물어보세요.")
    fq = st.text_area("이어서 묻기", height=68, key=f"ai_followup_{dk}_{len(chat)}", label_visibility="collapsed",
                      placeholder="예) 그럼 이 경우 토지소유자 동의도 받아야 해? / 관련 판례도 있어?")
    c1, c2 = st.columns([3, 1])
    with c1:
        fu = st.button("↪️ 이어서 묻기", use_container_width=True, key=f"ai_followup_btn_{dk}")
    with c2:
        if st.button("🧹 대화 지우기", use_container_width=True, key=f"ai_clear_btn_{dk}"):
            st.session_state[ck] = []
            st.session_state[sk] = None
            st.rerun()
    if fu:
        if not fq.strip():
            st.warning("이어서 물어볼 내용을 입력해주세요.")
        else:
            with st.spinner("앞 대화를 참고해 자료를 찾고 답변을 작성 중입니다..."):
                err = _run_turn(dk, fq.strip(), finder, field, mk)
            if err:
                _ai_status_msg(err)
            else:
                st.rerun()


# ==========================================
# [3. 화면 뷰 상태 관리 및 기억 장치]
# ==========================================
if 'view_law_data' not in st.session_state:
    st.session_state.view_law_data = None
if 'active_tab' not in st.session_state:
    st.session_state.active_tab = "📑 질의회신"
if 'saved_region' not in st.session_state:
    st.session_state.saved_region = "포항시"
if 'unlocked_region' not in st.session_state:
    st.session_state.unlocked_region = None
if 'ai_result' not in st.session_state:
    st.session_state.ai_result = None
if 'domain' not in st.session_state:
    st.session_state.domain = DOMAINS[0]
if 'active_tab_jj' not in st.session_state:
    st.session_state.active_tab_jj = "📘 지적업무사례집"
if 'ai_result_jj' not in st.session_state:
    st.session_state.ai_result_jj = None

current_view = st.query_params.get("view", "main")

def render_safe_html(text, kw=""):
    safe = html.escape(str(text)).replace("\n", "<br>")
    if kw:
        safe_kw = html.escape(str(kw))
        if safe_kw:
            safe = safe.replace(safe_kw, f"<mark style='background-color: yellow;'>{safe_kw}</mark>")
    return f'<div translate="no" class="notranslate" style="line-height:1.6;">{safe}</div>'


def render_ai_answer_html(text):
    raw = str(text)
    safe = html.escape(raw)
    safe = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', safe)
    # "## 제목" 같은 마크다운 머리줄은 # 기호를 지우고 굵은 글씨로
    safe = re.sub(r'(?m)^\s*#{1,6}\s*(.+)$', r'<strong style="font-size:1.05rem;">\1</strong>', safe)
    safe = safe.replace("\n", "<br>")
    return (
        '<div translate="no" class="notranslate" '
        'style="line-height:1.75; font-size:1rem; '
        'background:#f8fbff; border:1px solid #e1ecf7; '
        'border-radius:8px; padding:16px 18px;">'
        f'{safe}</div>'
    )

@st.dialog("📖 관련 법령 상세조회", width="large")
def show_law_detail_popup(law):
    st.markdown(f"### ⚖️ {law['조문']}")
    st.markdown(f"**📜 [법률]**<br>{render_safe_html(law['법률'])}", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown(f"**⚙️ [시행령]**<br>{render_safe_html(law['시행령'])}", unsafe_allow_html=True)
    if law['시행규칙'].strip():
        st.markdown("---")
        st.markdown(f"**📋 [시행규칙]**<br>{render_safe_html(law['시행규칙'])}", unsafe_allow_html=True)
    if st.button("닫기", use_container_width=True):
        st.rerun()

st.markdown("""
    <style>
    .notranslate { unicode-bidi: isolate; }
    button[kind="secondary"] {
        border: 2px solid #333333 !important;  
        border-radius: 5px !important;                     
    }
    button[kind="primary"] {
        background-color: transparent !important;
        border: none !important;
        box-shadow: none !important;
    }
    button[kind="primary"] p {
        font-size: 1.3rem !important;
        font-weight: bold !important;
        color: #2c3e50 !important;
        margin: 15px 0px !important;
        text-align: center !important; 
    }
    button[kind="primary"]:hover p {
        color: #3498db !important;
    }
    .qr-container {
        position: relative;
        width: 100%;
        height: 0px;
        top: -75px; 
        display: flex;
        justify-content: flex-end;
        z-index: 10;
        pointer-events: none; 
    }
    .qr-container img {
        width: 80px; 
        height: 80px; 
        border-radius: 5px; 
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        pointer-events: auto; 
    }
    @media (max-width: 768px) {
        .qr-container {
            display: none !important;
        }
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================================================
# 👥 [v23] 접속 기록 갱신 + 동시접속자 배지 표시 (최상단)
# ==========================================================================
_live_count = _touch_presence()
render_presence_badge(_live_count)
# ==========================================================================

upcoming = []
for info in all_events:
    if info.get("use_alarm") and info.get("region") == st.session_state.unlocked_region and info.get("region") != "경상북도(총괄)":
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
        st.session_state.active_tab_jj = "📅 공유달력"
        st.query_params.clear()
        st.rerun()

if notices:
    st.info(f"📢 **[전체 공지사항]** {notices[-1]['내용']}")

if st.button("🔍 물어보이소, AI — 지적·지적재조사 통합 검색", type="primary", use_container_width=True):
    st.query_params.clear()
    st.session_state.active_tab = "📑 질의회신"
    st.session_state.active_tab_jj = "📘 지적업무사례집"
    st.rerun()

# ==========================================
# [QR 이미지 생성 및 출력]
# ==========================================
target_url = "https://system-ydyhcgqqhe6dncgekqklcv.streamlit.app"
qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
qr.add_data(target_url)
qr.make(fit=True)
img = qr.make_image(fill_color="black", back_color="white")
buf = BytesIO()
img.save(buf, format="PNG")
b64_img = base64.b64encode(buf.getvalue()).decode()
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
# 📱 AI 질문 탭에서는 위쪽 키워드 검색창이 필요 없으므로 숨겨서, 질문 입력칸이 화면 위쪽에 바로 보이게 함
_cur_dom = st.session_state.get("domain", DOMAINS[0])
ai_mode = ((_cur_dom == DOMAINS[0] and st.session_state.get("active_tab") == "🤖 AI 질문") or
           (_cur_dom != DOMAINS[0] and st.session_state.get("active_tab_jj") == "🤖 AI 질문"))

if not ai_mode:
    col1, col2 = st.columns([3, 1])
    with col1: keyword = st.text_input("검색어를 입력하세요", label_visibility="collapsed", placeholder="🔍 검색어 입력 (예: 경계설정)")
    with col2: search_btn = st.button("검색", use_container_width=True)
    only_title = st.checkbox("☑️ 제목만 검색", value=True)
else:
    keyword, search_btn, only_title = "", False, True

domain = st.radio("업무 분야", DOMAINS, horizontal=True, key="domain")

if domain == DOMAINS[0]:
    tabs = ["🤖 AI 질문", "📑 질의회신", "⚖️ 법령", "🏢 업무규정", "📐 측량규정", "🏢 판례", "📅 공유달력"]
    mode = st.radio("자료 선택", tabs, horizontal=True, label_visibility="collapsed", key="active_tab")
else:
    jj_tabs = ["🤖 AI 질문"] + [f"{ic} {nm}" for nm, ic in JJ_SHEETS] + ["⚖️ 공간정보관리법", "📅 공유달력"]
    mode_jj = st.radio("자료 선택", jj_tabs, horizontal=True, label_visibility="collapsed", key="active_tab_jj")
    # 공유달력은 아래 기존 화면을 그대로 사용, 그 밖에는 지적 화면만 그리고 기존 분기는 건너뜀
    mode = "📅 공유달력" if mode_jj == "📅 공유달력" else "__지적__"
if not ai_mode:
    st.markdown("---")

# 탭이 바뀐 실행인지 기록 (AI 질문 탭에 '막 들어왔을 때'만 자동 스크롤하기 위함)
_tab_now = (domain, mode if domain == DOMAINS[0] else mode_jj)
_tab_changed = st.session_state.get("_last_tab_for_scroll") != _tab_now
st.session_state["_last_tab_for_scroll"] = _tab_now


def _scroll_to_question_on_mobile():
    """📱 휴대폰 화면에서 AI 질문 탭을 막 눌렀을 때 질문 입력칸이 화면 맨 위로 오도록 자동 스크롤"""
    if not _tab_changed:
        return
    components.html("""
    <script>
    setTimeout(function () {
      try {
        var w = window.parent;
        if (w.innerWidth > 768) return;
        var ta = w.document.querySelector('textarea');
        if (ta) { ta.scrollIntoView({behavior: 'smooth', block: 'center'}); }
      } catch (e) {}
    }, 400);
    </script>""", height=0)


def _search_df(df, kw, title_only):
    """띄어쓰기로 여러 단어 AND, 공백 무시"""
    if not kw:
        return df
    words = [_strip_spaces(w) for w in kw.split() if w.strip()]
    t = df["제목"].map(_strip_spaces)
    hay = t if title_only else t + "\n" + df["내용"].map(_strip_spaces)
    mask = pd.Series(True, index=df.index)
    for w in words:
        mask &= hay.str.contains(w, regex=False, na=False)
    return df[mask]


def _law_buttons(text, laws, key_prefix):
    """본문에 나온 '제N조'와 같은 조문이 법령에 있으면 바로보기 버튼"""
    jos = {re.sub(r"\s+", "", j) for j in re.findall(r"제\s*\d+\s*조(?:의\s*\d+)?", str(text))}
    matched = [law for law in laws if law["조문"].split("(")[0].replace(" ", "") in jos]
    if matched:
        st.markdown("---")
        st.markdown("🔗 **관련 법령 바로보기** (공간정보관리법)")
        cols = st.columns(min(len(matched), 3))
        for i, law in enumerate(matched[:9]):
            with cols[i % 3]:
                if st.button(f"📖 {law['조문'].split('(')[0]}", key=f"{key_prefix}_{i}"):
                    show_law_detail_popup(law)


if mode == "__지적__":
    jj_sheets, jj_law_db = load_jijeok_data()

    if mode_jj == "🤖 AI 질문":
        render_ai_chat(
            "jj", "지적",
            "예) 등록전환할 때 면적 오차 허용범위는? / 공간정보관리법 제84조 알려줘",
            "지적업무사례집·토지이동업무편람·지적판례·지적규정·공간정보관리법에서 근거를 찾아 답변합니다. (근거 없는 내용은 답하지 않습니다)",
            lambda qq: find_relevant_materials_jj(qq, jj_sheets, jj_law_db, max_items=7),
            "지적(토지이동·지적측량·지적공부)",
        )

    elif mode_jj == "⚖️ 공간정보관리법":
        hits = [it for it in jj_law_db if not keyword or
                all(_strip_spaces(w) in _strip_spaces(it["조문"] if only_title else it["조문"] + it["법률"] + it["시행령"] + it["시행규칙"])
                    for w in keyword.split())]
        st.caption(f"총 {len(hits)}개 조문 (공간정보의 구축 및 관리 등에 관한 법률 · 3단비교)")
        for item in hits:
            with st.expander(f"⚖️ [법령] {item['조문']}"):
                st.markdown(f"**📜 [법률]**<br>{render_safe_html(item['법률'], keyword)}", unsafe_allow_html=True)
                st.markdown("---")
                st.markdown(f"**⚙️ [시행령]**<br>{render_safe_html(item['시행령'], keyword)}", unsafe_allow_html=True)
                if item['시행규칙'].strip():
                    st.markdown("---")
                    st.markdown(f"**📋 [시행규칙]**<br>{render_safe_html(item['시행규칙'], keyword)}", unsafe_allow_html=True)

    else:
        sheet_name = mode_jj.split(" ", 1)[1]
        icon = mode_jj.split(" ", 1)[0]
        res = _search_df(jj_sheets.get(sheet_name, pd.DataFrame(columns=["제목", "내용", "수정여부"])), keyword, only_title)
        st.caption(f"총 {len(res)}건의 자료가 있습니다." + (f" (앞 {MAX_LIST}건만 표시 — 검색어로 좁혀 보세요)" if len(res) > MAX_LIST else ""))
        for idx, row in res.head(MAX_LIST).iterrows():
            ic = "🟢" if str(row.get("수정여부")).strip().upper() == "Y" else icon
            with st.expander(f"{ic} {row['제목']}"):
                st.markdown(render_safe_html(row['내용'], keyword), unsafe_allow_html=True)
                if sheet_name != "지적규정":
                    _law_buttons(row['내용'], jj_law_db, f"jjlaw_{sheet_name}_{idx}")

# ==========================================
# [6. 카테고리별 출력 및 일정 관리]
# ==========================================

# 🤖 AI 질문 탭
if mode == "__지적__":
    pass   # 📐 지적 분야 화면은 위에서 그림
elif mode == "🤖 AI 질문":
    render_ai_chat(
        "jr", "지적재조사",
        "예) 사업지구를 경미하게 변경할 때 토지소유자 동의가 필요한가요? / 제38조 알려줘",
        "질의회신·법령·규정·판례 전체에서 관련 근거를 찾아 답변합니다. (근거 없는 내용은 답하지 않습니다)",
        lambda qq: find_relevant_materials(qq, df_qna, df_case, law_db, reg_db, max_items=7),
        "지적재조사",
    )

elif mode in ["📑 질의회신", "🏢 판례"]:
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
                                show_law_detail_popup(law)
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
    regions = ["포항시", "포항시 남구", "포항시 북구", "경주시", "김천시", "안동시", "구미시", "영주시", "영천시", "상주시", "문경시", "경산시", "의성군", "청송군", "영양군", "영덕군", "청도군", "고령군", "성주군", "칠곡군", "예천군", "봉화군", "울진군", "울릉군", "경상북도(총괄)"]

    default_idx = regions.index(st.session_state.saved_region) if st.session_state.saved_region in regions else 0
    selected_region = st.selectbox("📌 담당 시/군을 선택하세요", regions, index=default_idx)

    if selected_region != st.session_state.saved_region:
        st.session_state.saved_region = selected_region
        st.session_state.unlocked_region = None

    col_pw, col_btn = st.columns([3, 1])
    with col_pw: entered_pw = st.text_input("🔑 비밀번호 4자리를 입력하세요", type="password")
    with col_btn:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        login_btn = st.button("확인", use_container_width=True)

    is_unlocked = False

    if st.session_state.unlocked_region == selected_region:
        is_unlocked = True

    if entered_pw or login_btn:
        if entered_pw:
            try:
                if entered_pw == st.secrets["passwords"][selected_region]: 
                    if not is_unlocked:
                        st.session_state.unlocked_region = selected_region
                        st.rerun() 
                    is_unlocked = True
                else: 
                    st.error("❌ 비밀번호가 일치하지 않습니다.")
                    st.session_state.unlocked_region = None
                    is_unlocked = False
            except KeyError: 
                st.warning("⚠️ 이 지역의 비밀번호가 설정되지 않았습니다.")
        elif login_btn: 
            st.warning("비밀번호를 입력해주세요.")

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
st.caption("v26.0 - 이어서 묻기(대화형 AI 질문) + 모델 선택(Gemini·Gemma·Solar)")
st.caption("v24.1 - 모바일: AI 질문 탭에서 검색창 숨김 + 질문칸 자동 스크롤")
st.caption("v24.0 - 물어보이소, AI: 지적 분야 추가(지적업무사례집·토지이동업무편람·지적판례·지적규정·공간정보관리법)")
st.caption(f"v23.0 - 동시접속자 실시간 카운터(최근 {PRESENCE_WINDOW_SEC}초 기준 / 주의 {WARN_LEVEL}명·위험 {DANGER_LEVEL}명)")
