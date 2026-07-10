# -*- coding: utf-8 -*-
"""
tab_spellcheck.py — 공문서 맞춤법 검사 탭

설계 원칙
  1) 과교정 억제 : 가운뎃점(·), 공문서 고유표현(붙임/끝.), 법령용어, 숫자+단위는 건드리지 않음
  2) 환각 방지  : 교정 근거 필수 · 확신 없으면 "확인 필요" · 없는 단어 지어내기 금지 · 오류 없으면 "없음"
  3) 신뢰도 3단계: ✅ 확실한 오류 / 🔶 참고 제안 / ❓ 확인 필요
  4) 입력       : HWPX 업로드만 (화면 확인 전용, 교정본 생성 없음)

호출: main.py 에서  import  후  render_spellcheck_tab()  실행
"""

import io
import re
import json
import zipfile
import xml.etree.ElementTree as ET

import streamlit as st

try:
    import google.generativeai as genai
except ImportError:
    genai = None


# ──────────────────────────────────────────────────────────────
# 1. HWPX 텍스트 추출  (탭7 로직 재활용: content.hpf spine 순서 + 섹션 XML 파싱)
# ──────────────────────────────────────────────────────────────
_NS_TEXT_TAGS = {"t", "char"}  # 네임스페이스 접미사 기준으로 텍스트를 담는 태그


def _localname(tag: str) -> str:
    """'{ns}t' -> 't' 처럼 네임스페이스를 벗겨 로컬 태그명만 반환."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _section_order(zf: zipfile.ZipFile) -> list[str]:
    """
    content.hpf(spine)에서 섹션 파일 순서를 읽는다.
    실패하면 파일명 정렬(section0.xml, section1.xml ...)로 폴백.
    """
    ordered: list[str] = []
    try:
        hpf = None
        for name in zf.namelist():
            if name.endswith("content.hpf"):
                hpf = name
                break
        if hpf:
            root = ET.fromstring(zf.read(hpf))
            # spine 안의 itemref idref 순서 → manifest의 href 로 매핑
            manifest = {}
            for el in root.iter():
                if _localname(el.tag) == "item":
                    mid = el.get("id")
                    href = el.get("href")
                    if mid and href:
                        manifest[mid] = href
            for el in root.iter():
                if _localname(el.tag) == "itemref":
                    ref = el.get("idref")
                    if ref and ref in manifest:
                        href = manifest[ref]
                        # href 는 보통 Contents/section0.xml 형태
                        for n in zf.namelist():
                            if n.endswith(href.split("/")[-1]) and "section" in n.lower():
                                ordered.append(n)
    except Exception:
        pass

    if not ordered:
        secs = [n for n in zf.namelist()
                if re.search(r"section\d+\.xml$", n, re.IGNORECASE)]
        secs.sort(key=lambda n: int(re.search(r"section(\d+)", n, re.I).group(1)))
        ordered = secs
    return ordered


def extract_hwpx_paragraphs(file_bytes: bytes) -> list[str]:
    """HWPX 바이트 → 문단 리스트. 문단은 <p> 단위, 안의 텍스트 조각을 이어붙임."""
    paragraphs: list[str] = []
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
        for sec in _section_order(zf):
            try:
                root = ET.fromstring(zf.read(sec))
            except Exception:
                continue
            for p in root.iter():
                if _localname(p.tag) != "p":
                    continue
                buf: list[str] = []
                for node in p.iter():
                    if _localname(node.tag) in _NS_TEXT_TAGS and node.text:
                        buf.append(node.text)
                line = "".join(buf).strip()
                if line:
                    paragraphs.append(line)
    return paragraphs


# ──────────────────────────────────────────────────────────────
# 2. 프롬프트 (과교정 억제 + 환각 방지 + 신뢰도 3단계 + JSON 강제)
# ──────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """당신은 대한민국 공문서 전문 교정자입니다.
아래 [절대 원칙]과 [교정 제외 목록]을 반드시 지키고, 지정된 JSON 형식으로만 답하세요.

[절대 원칙]
- 교정 근거를 반드시 제시할 것 (어떤 맞춤법/띄어쓰기 규칙인지 한 줄로).
- 확신이 없으면 억지로 고치지 말고 신뢰도를 "확인필요"로 표시할 것.
- 원문에 없는 단어를 있다고 하지 말 것. 실제 원문 표기를 그대로 인용할 것.
- 오류가 하나도 없으면 빈 배열([])을 반환할 것. 억지로 찾아내지 말 것.
- "지적재조사"가 올바른 표기다. "지적제조사"로 바꾸지 말 것.

[교정 제외 목록 — 절대 오류로 지적하지 말 것]
- 가운뎃점(·) 사용 표현: 시·군, 검토·보고, 세계측지계 등
- 공문서 고유 표현: 붙임, 위와 같이, 끝., 아래와 같이
- 법령·조례·행정 용어 원문 (지적재조사, 등록사항정정, 이행강제금 등)
- 숫자+단위 표현: 33,676,691천원, 450,000천원, 30%, 387.8억 원
- 이미 올바른 띄어쓰기를 틀렸다고 하지 말 것

[교정 대상]
- 명백한 오탈자 (맞춤뻡→맞춤법)
- 조사 오류 (을/를, 이/가, 로/으로)
- 명백한 띄어쓰기 오류 (2026년 까지→2026년까지)

[신뢰도 3단계]
- "확실"   : 명백한 오탈자·조사오류·표준 띄어쓰기 위반
- "참고"   : 문체상 권장이나 틀렸다고 단정 못함
- "확인필요": 문맥상 판단이 갈리거나 확신 없음

[출력 형식 — 아래 JSON 배열만 출력. 설명·마크다운·코드펜스 금지]
[
  {
    "para": <문단번호(정수)>,
    "before": "<원문의 틀린 부분 그대로>",
    "after": "<교정 제안>",
    "reason": "<근거 한 줄>",
    "level": "확실|참고|확인필요"
  }
]
오류가 없으면 정확히 [] 만 출력하세요."""


def build_user_prompt(paragraphs: list[str]) -> str:
    numbered = "\n".join(f"[{i}] {p}" for i, p in enumerate(paragraphs))
    return f"다음 공문서 문단들을 검사하세요. 각 문단은 [번호]로 시작합니다.\n\n{numbered}"


# ──────────────────────────────────────────────────────────────
# 3. Gemini 호출 + 안전 파싱
# ──────────────────────────────────────────────────────────────
def _strip_fence(text: str) -> str:
    """```json ... ``` 펜스 제거."""
    text = text.strip()
    text = re.sub(r"^```(?:json)?", "", text).strip()
    text = re.sub(r"```$", "", text).strip()
    return text


def run_spellcheck(paragraphs: list[str], api_key: str,
                   model_name: str = "gemini-2.5-flash") -> list[dict]:
    if genai is None:
        raise RuntimeError("google-generativeai 패키지가 설치되어 있지 않습니다.")

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(
        model_name,
        system_instruction=SYSTEM_PROMPT,
        generation_config={
            "temperature": 0.0,        # 환각·과교정 억제
            "max_output_tokens": 8192,
            "response_mime_type": "application/json",  # JSON 강제
        },
    )
    resp = model.generate_content(build_user_prompt(paragraphs))
    raw = _strip_fence(resp.text or "[]")

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        # 배열 부분만 추출 재시도
        m = re.search(r"\[.*\]", raw, re.DOTALL)
        data = json.loads(m.group(0)) if m else []

    if not isinstance(data, list):
        data = []

    # 필드 정규화 + para 범위 검증(환각으로 없는 문단 지목하는 것 방지)
    cleaned = []
    for item in data:
        if not isinstance(item, dict):
            continue
        para = item.get("para")
        if not isinstance(para, int) or not (0 <= para < len(paragraphs)):
            continue
        before = str(item.get("before", "")).strip()
        # before 가 실제 원문에 존재하는지 검증 → 환각 차단
        if before and before not in paragraphs[para]:
            item["level"] = "확인필요"
            item["reason"] = "(원문에서 해당 표기를 찾지 못함) " + str(item.get("reason", ""))
        cleaned.append({
            "para": para,
            "before": before,
            "after": str(item.get("after", "")).strip(),
            "reason": str(item.get("reason", "")).strip(),
            "level": item.get("level", "확인필요"),
        })
    return cleaned


# ──────────────────────────────────────────────────────────────
# 4. 화면 렌더링
# ──────────────────────────────────────────────────────────────
_LEVEL_META = {
    "확실":   ("✅", "확실한 오류", "#d32f2f"),
    "참고":   ("🔶", "참고 제안",   "#f9a825"),
    "확인필요": ("❓", "확인 필요",   "#757575"),
}


def render_spellcheck_tab():
    st.header("📝 공문서 맞춤법 검사")
    st.caption("HWPX 업로드 → 문단 추출 → AI 검사 · 결과는 화면 확인 전용입니다.")

    api_key = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        api_key = st.text_input("Gemini API Key", type="password")

    uploaded = st.file_uploader("HWPX 파일 업로드", type=["hwpx"])
    if uploaded is None:
        st.info("검사할 .hwpx 파일을 올려주세요.")
        return

    try:
        paragraphs = extract_hwpx_paragraphs(uploaded.read())
    except Exception as e:
        st.error(f"HWPX 파싱 실패: {e}")
        return

    if not paragraphs:
        st.warning("추출된 텍스트가 없습니다. 파일 형식을 확인해 주세요.")
        return

    st.success(f"문단 {len(paragraphs)}개 추출 완료")
    with st.expander("추출된 문단 미리보기", expanded=False):
        for i, p in enumerate(paragraphs):
            st.text(f"[{i}] {p}")

    if not st.button("맞춤법 검사 실행", type="primary"):
        return
    if not api_key:
        st.error("API Key를 입력해 주세요.")
        return

    with st.spinner("AI가 검사 중입니다..."):
        try:
            results = run_spellcheck(paragraphs, api_key)
        except Exception as e:
            st.error(f"검사 실패: {e}")
            return

    if not results:
        st.success("발견된 오류가 없습니다. ✅")
        return

    # 신뢰도별 집계
    order = ["확실", "참고", "확인필요"]
    buckets = {k: [r for r in results if r["level"] == k] for k in order}
    unknown = [r for r in results if r["level"] not in order]
    if unknown:
        buckets["확인필요"].extend(unknown)

    c1, c2, c3 = st.columns(3)
    c1.metric("✅ 확실한 오류", len(buckets["확실"]))
    c2.metric("🔶 참고 제안", len(buckets["참고"]))
    c3.metric("❓ 확인 필요", len(buckets["확인필요"]))

    st.divider()

    for level in order:
        rows = buckets[level]
        if not rows:
            continue
        icon, label, color = _LEVEL_META[level]
        st.subheader(f"{icon} {label} ({len(rows)}건)")
        for r in rows:
            st.markdown(
                f"<div style='border-left:4px solid {color};padding:6px 12px;margin:6px 0;'>"
                f"<b>[문단 {r['para']}]</b>&nbsp;&nbsp;"
                f"<span style='color:{color};'>{r['before'] or '—'}</span> "
                f"→ <b>{r['after'] or '—'}</b><br>"
                f"<small style='color:#666;'>근거: {r['reason'] or '—'}</small>"
                f"</div>",
                unsafe_allow_html=True,
            )


# 단독 실행 테스트용
if __name__ == "__main__":
    render_spellcheck_tab()
