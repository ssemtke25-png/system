"""
탭17: 서무비서 — 상황을 말하면 절차·서식·근거를 안내 (RAG)
담당자가 '상황'을 말하면 서무규정(공개자료)에서 근거를 찾아
① 처리 절차 ② 기한·주의 ③ 필요 서식 ④ 근거([자료N])를 정리한다.
모델 교체형(Gemini / Gemma / Solar) — 공공 AI 대전환 챌린지 '모델 비종속' 대응.

데이터 : 서무규정.xlsx (제목 / 내용 / 수정여부) — 전부 공개자료
키     : st.secrets 의 GEMINI_API_KEY(Gemini·Gemma), UPSTAGE_API_KEY(Solar)
         (secrets에 없으면 탭 안의 '🔑 API 키' 칸에 붙여넣어 사용)
필요   : pip install google-genai openai
app.py 에서 with 탭: render() 형태로 호출.
"""
import re, time
from pathlib import Path
import pandas as pd
import streamlit as st

# ── 데이터 위치: 아래 후보 중 먼저 찾는 파일 사용 ─────────────────────
_HERE = Path(__file__).resolve().parent
DATA_CANDIDATES = [
    _HERE / "서무규정.xlsx",
    _HERE.parent / "서무규정.xlsx",
    _HERE.parent / "data" / "서무규정.xlsx",
    Path("서무규정.xlsx"),
    Path("data") / "서무규정.xlsx",
]

def _find_data():
    for p in DATA_CANDIDATES:
        if p.exists():
            return p
    return None

# ==========================================
# [1] 데이터 로드 + 검색 인덱스
# ==========================================
@st.cache_data(ttl=3600)
def load_rules(path_str, mtime=0.0):
    """mtime(파일 수정시각)을 캐시 키에 포함 → 서무규정.xlsx를 교체하면 즉시 새 데이터로 다시 읽는다."""
    df = pd.read_excel(path_str).fillna("").astype(str)
    return [(r["제목"], r["내용"]) for _, r in df.iterrows()]

def _nosp(s):
    return re.sub(r"\s+", "", str(s))

def _norm_jo(s):
    return re.sub(r"제?\s*(\d+)\s*조(?:\s*의\s*(\d+))?",
                  lambda m: f"제{m.group(1)}조의{m.group(2)}" if m.group(2) else f"제{m.group(1)}조", str(s))

_JOSA = sorted({"으로서","으로써","이라고","라고","에서","에게","으로","로서","까지","부터","에는","에도",
                "이나","은","는","이","가","을","를","의","에","로","도","만","와","과"}, key=len, reverse=True)

def _strip_josa(w):
    if len(w) <= 2:
        return w
    for j in _JOSA:
        if w.endswith(j) and len(w) - len(j) >= 2:
            return w[:-len(j)]
    return w

_STOP = {"어떻게","무엇","인가요","하나요","되나요","있나요","경우","관련","대한","대해","해야","하는","가능",
         "여부","알려줘","알려주세요","설명","질문","해주세요","하려면","하고","싶어요","해도","되는지","하는지","할때","때",
         # 거의 모든 규정에 나와 변별력이 없는 말 / 질문 말투
         "공무원","며칠","몇일","얼마나","있어","있어요","받아","받나요","받을","쓸","쓸수","있을까","되나","돼","해줘",
         "궁금","궁금해","궁금해요","좋을까","하나","되요","돼요","인가","인지"}

# 질문 말투의 어미·존칭을 떼어 원문 단어와 맞춘다 (예: 결혼하면→결혼, 돌아가셨을→돌아가, 부모님→부모)
_ENDINGS = sorted({"하려면","했는데","했을때","하는데","하면은","하면","해서","하고","했을","했어",
                   "되면","되는데","됐는데","됐을","셨을","셨는데","셨으면","셨어","시면","으면","는데",
                   "님이","님의","님은","님께서","님"}, key=len, reverse=True)

def _stem(w):
    w = _strip_josa(w)
    for e in _ENDINGS:
        if w.endswith(e) and len(w) - len(e) >= 2:
            return w[:-len(e)]
    return w

# 질문 속 '상황'을 보고 규정 용어를 덧붙인다 (담당자는 연가·휴가·경조사를 섞어 쓰므로)
#   예) "부모님이 돌아가셨을 경우 연가" → 경조사·특별휴가·사망도 함께 검색
_SITUATION = [
    (("돌아가", "사망", "별세", "부친상", "모친상", "상을", "상당", "장례", "조문", "빈소", "작고"),
     ["경조사", "특별휴가", "사망"]),
    (("결혼", "혼인", "결혼식", "장가", "시집"), ["경조사", "특별휴가", "결혼"]),
    (("출산", "아기", "낳"), ["경조사", "특별휴가", "출산"]),
    (("입양",), ["경조사", "특별휴가", "입양"]),
]
_ALIAS = {"경북": "경상북도", "도청": "경상북도"}

def _keywords(q):
    q = re.sub(r"제?\s*\d+\s*조(?:\s*의\s*\d+)?", " ", q)
    for a, b in _ALIAS.items():
        q = q.replace(a, b)
    out = []
    for w in re.findall(r"[가-힣A-Za-z0-9]{2,}", q):
        if w in _STOP or _strip_josa(w) in _STOP:
            continue
        out.append(w)
    for triggers, extra in _SITUATION:
        if any(t in q for t in triggers):
            out += [e for e in extra if e not in out]
    return out

@st.cache_data(ttl=3600)
def build_index(rules):
    idx = []
    for title, content in rules:
        nt, nc = _norm_jo(title), _norm_jo(content)
        idx.append((title, content, nt, nc, _nosp(nt), _nosp(nc)))
    return idx

def retrieve(question, idx, topk=7):
    jos = set(re.findall(r"제\d+조(?:의\d+)?", _norm_jo(question)))
    kws = _keywords(question) or [question]
    scored = []
    for title, content, nt, nc, snt, snc in idx:
        s, hit = 0, 0
        for jo in jos:
            if jo in nt: s += 20
            if jo in nc: s += 8
        for kw in kws:
            t = title.count(kw); c = content.count(kw)
            for key in {_nosp(_strip_josa(kw)), _nosp(_stem(kw))}:
                if len(key) >= 2:
                    t = max(t, snt.count(key)); c = max(c, snc.count(key))
            # 긴 문서가 같은 단어를 수십 번 반복해 1위를 차지하지 않도록 본문 횟수는 3회까지만 반영
            s += min(t, 2) * 6 + min(c, 3)
            if t or c:
                hit += 1
        # 질문의 여러 단어가 '함께' 들어있는 문서를 우대 (핵심 변별 요소)
        if hit >= 2:
            s += hit * 10
        if s > 0:
            scored.append((s, title, content))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[:topk]

# ==========================================
# [2] 모델 어댑터 (비종속) — Gemini / Gemma / Solar
# ==========================================
SYSTEM = (
    "너는 공공기관의 '서무 업무비서' AI다. 담당자가 상황을 말하면, 아래 [참고규정]에 실제로 있는 내용만 근거로 "
    "다음을 정리해 안내한다.\n"
    "1) 처리 절차 — 순서대로 번호로\n"
    "2) 기한·주의사항 — 있으면\n"
    "3) 필요 서식·서류 — 규정/별표에 언급된 것\n"
    "4) 근거 — 각 항목 끝에 [자료N] 형태로 표시\n"
    "[용어] 담당자가 용어를 잘못 쓸 수 있다(예: 부모 사망인데 '연가'라고 질문). 질문 용어보다 실제 상황에 맞는 "
    "규정(예: 경조사휴가=특별휴가)을 우선 안내하고, 용어 차이를 한 줄로 바로잡아 준다.\n"
    "[별표] 참고규정에 [별표](일수표·기준표 등)가 있으면 그 표의 해당 행 값(일수·금액 등)을 그대로 제시하고 "
    "어느 별표의 어느 항목인지 밝힌다.\n"
    "[규칙] 참고규정에 없는 사실·금액·기한·조항은 절대 지어내지 말고, 확인되지 않으면 "
    "'제공된 규정에서는 확인되지 않습니다'라고 명확히 밝힌다. 추측성 일반론을 사실처럼 쓰지 않는다."
)

MODELS = {
    "Gemini (빠름)": "gemini",
    "Gemma (공통기반 지원모델)": "gemma",
    "Solar (국산)": "solar",
}

def _context(materials):
    return "\n\n---\n\n".join(f"[자료{i}] {t}\n{c[:4500]}" for i, (s, t, c) in enumerate(materials, 1))

def _secret(key):
    try:
        return st.secrets.get(key, "")
    except Exception:
        return ""

def ask_llm(question, materials, model_key, gkey="", ukey=""):
    prompt = f"[참고규정]\n{_context(materials)}\n\n---\n\n[상황/질문]\n{question}"
    gkey, ukey = (gkey or "").strip(), (ukey or "").strip()
    if model_key in ("gemini", "gemma"):
        from google import genai
        from google.genai import types
        model = "gemini-2.5-flash" if model_key == "gemini" else "gemma-4-31b-it"
        client = genai.Client(api_key=gkey)
        last = None
        for attempt in range(3):   # 429/503 재시도
            try:
                r = client.models.generate_content(
                    model=model, contents=prompt,
                    config=types.GenerateContentConfig(system_instruction=SYSTEM,
                                                       temperature=0.2, max_output_tokens=4096))
                return r.text
            except Exception as e:
                last = e
                if any(x in str(e) for x in ("429", "503", "UNAVAILABLE", "RESOURCE_EXHAUSTED")):
                    time.sleep(2 * (attempt + 1)); continue
                raise
        raise last
    if model_key == "solar":
        from openai import OpenAI
        client = OpenAI(api_key=ukey, base_url="https://api.upstage.ai/v1")
        r = client.chat.completions.create(
            model="solar-pro2",
            messages=[{"role": "system", "content": SYSTEM}, {"role": "user", "content": prompt}],
            temperature=0.2, max_tokens=4096)
        return r.choices[0].message.content
    raise ValueError(model_key)

# ==========================================
# [3] 화면
# ==========================================
EXAMPLES = [
    "출장 다녀와서 여비 정산하려면 어떤 절차와 서식이 필요한가요?",
    "출장 가서 자가용 썼는데 연료비 받을 수 있나요?",
    "행사 집행인데 사무관리비로 써도 되나요?",
    "연가는 며칠까지 쓸 수 있나요?",
]

def render():
    """서무비서 탭 화면. app.py에서 with 탭: render() 형태로 호출."""
    st.caption("상황을 말로 입력하면 서무 규정에서 근거를 찾아 절차·기한·서식·근거를 안내합니다.")

    path = _find_data()
    if path is None:
        st.error("서무규정.xlsx 를 찾을 수 없습니다. 저장소의 modules/ 또는 data/ 폴더에 넣어주세요.")
        return
    rules = load_rules(str(path), path.stat().st_mtime)
    idx = build_index(rules)
    st.info(f"📚 담긴 규정·실무 데이터 {len(rules)}건 — 복무·여비·물품·문서·당직 규정·별표, "
            "예산 통계목 운용기준, 공무원 여비 100문100답 (전부 공개자료)")

    # 키: secrets 우선, 없으면 입력칸
    gkey_s, ukey_s = _secret("GEMINI_API_KEY"), _secret("UPSTAGE_API_KEY")
    gkey, ukey = gkey_s, ukey_s
    if not (gkey_s and ukey_s):
        with st.expander("🔑 API 키 (secrets에 없을 때만 입력)", expanded=not gkey_s):
            if not gkey_s:
                gkey = st.text_input("GEMINI_API_KEY (Gemini·Gemma)", type="password", key="seomu_gkey")
            if not ukey_s:
                ukey = st.text_input("UPSTAGE_API_KEY (Solar)", type="password", key="seomu_ukey")

    c1, c2 = st.columns([3, 1])
    with c2:
        mlabel = st.selectbox("AI 모델", list(MODELS.keys()), key="seomu_model")
        ex = st.selectbox("예시 질문", ["(직접 입력)"] + EXAMPLES, key="seomu_ex")
    with c1:
        default_q = "" if ex == "(직접 입력)" else ex
        q = st.text_area("상황을 말로 입력하세요", value=default_q, height=110, key=f"seomu_q_{ex}",
                         placeholder="예) 출장 다녀와서 여비 정산하려면 어떤 절차와 서식이 필요한가요?")
    go = st.button("📋 안내받기", type="primary", key="seomu_go")

    if go:
        if not q.strip():
            st.warning("상황을 입력해주세요."); return
        mk = MODELS[mlabel]
        if mk in ("gemini", "gemma") and not gkey:
            st.warning("GEMINI_API_KEY가 필요합니다 (secrets 또는 위 키 입력칸)."); return
        if mk == "solar" and not ukey:
            st.warning("UPSTAGE_API_KEY가 필요합니다 (secrets 또는 위 키 입력칸)."); return
        with st.spinner("관련 규정을 찾고 절차를 정리하는 중..."):
            mats = retrieve(q, idx, topk=7)
            if not mats:
                st.info("관련 규정을 찾지 못했어요. 다른 표현으로 다시 물어보세요."); return
            t0 = time.time()
            try:
                ans = ask_llm(q, mats, mk, gkey, ukey)
            except Exception as e:
                st.error(f"모델 호출 오류: {e}"); return
            dt = time.time() - t0
        st.session_state["seomu_last"] = (mlabel, dt, ans, mats)

    last = st.session_state.get("seomu_last")
    if last:
        mlabel_l, dt, ans, mats = last
        st.markdown("### 💡 안내")
        st.success(f"🤖 **{mlabel_l}** 모델이 생성 · {dt:.1f}초  (같은 질문을 모델만 바꿔 실행 → 모델 비종속 확인)")
        st.markdown(ans)
        st.markdown("---")
        st.markdown("#### 📚 근거 규정 (원문 확인)")
        for i, (s, t, c) in enumerate(mats, 1):
            with st.expander(f"[자료{i}] {t}"):
                st.text(c)


if __name__ == "__main__":
    st.set_page_config(page_title="서무비서", page_icon="🧑‍💼", layout="wide")
    render()
