"""
탭4: 한글(HWPX) 파일 병합

여러 hwpx 파일을 받아서, 1번 파일 본문 끝에 2번 파일 본문을 그대로 이어붙이고,
그 끝에 3번 파일을 또 이어붙이는 식으로 순서대로 모두 합친다.

순서 판단:
- 파일명에 시군명이나 숫자(01, 02... 또는 1_, 2_)가 있으면 그 순서대로.
- 그런 단서가 전혀 없으면 사용자가 업로드한 순서 그대로 한 파일씩 이어붙인다.

기존에는 Windows에서 한글 프로그램을 COM으로 직접 띄워(pythoncom, win32com)
InsertFile 기능으로 병합했으나, 이 방식은 Windows 환경에서만 동작하고
Streamlit Cloud 같은 Linux 서버에서는 'pythoncom' 모듈 자체가 없어 실행이
불가능했다. hwpx는 zip으로 압축된 XML 문서이므로, 한글 프로그램 없이도
zip+XML을 직접 조작해서 동일한 결과(본문 이어붙이기)를 만들 수 있다.

처리 항목:
- header.xml의 스타일 정의(charProperties, paraProperties, styles, borderFills,
  tabProperties, numberings, bullets, fontfaces)를 모두 합치고, 새로 부여된 id로
  본문(section0.xml)과 header.xml 안의 모든 IDRef 참조를 일괄 치환한다.
- BinData(이미지) 파일명이 서로 겹치지 않도록 번호를 이어서 재명명하고,
  content.hpf의 이미지 등록 항목과 section0.xml의 binaryItemIDRef도 같이 갱신한다.
- section0.xml의 모든 단락(p)을 그대로 이어붙인다 (서식은 원본 그대로 보존).
"""
import io
import re
import copy
import zipfile

import streamlit as st
from lxml import etree

NS_HH = "http://www.hancom.co.kr/hwpml/2011/head"
NS_HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
NS_OPF = "http://www.idpf.org/2007/opf/"

# 언어 그룹 없는 일반 컨테이너들 (header.xml의 refList 하위)
SIMPLE_CONTAINERS = ["borderFills", "charProperties", "tabProperties",
                      "numberings", "paraProperties", "styles", "bullets"]

# 각 컨테이너의 id를 참조하는 IDRef 속성 이름 (header.xml + section0.xml 전체에서 치환 대상)
IDREF_FOR_CONTAINER = {
    "borderFills": ["borderFillIDRef"],
    "charProperties": ["charPrIDRef"],
    "tabProperties": ["tabPrIDRef"],
    "numberings": ["numberingIDRef"],
    "paraProperties": ["paraPrIDRef"],
    "styles": ["styleIDRef", "nextStyleIDRef"],
    "bullets": ["bulletIDRef"],
}

LANG_KEY_MAP = {
    "HANGUL": "hangul", "LATIN": "latin", "HANJA": "hanja", "JAPANESE": "japanese",
    "OTHER": "other", "SYMBOL": "symbol", "USER": "user",
}


def _qn_hh(tag):
    return f"{{{NS_HH}}}{tag}"


def _get_reflist(header_root):
    return header_root.find(_qn_hh("refList"))


def _merge_simple_container(reflist1, reflist2, container_tag):
    """borderFills, charProperties 등 단순 컨테이너를 병합.
    return: {old_id(int): new_id(int)} f2 기준 매핑"""
    cont1 = reflist1.find(_qn_hh(container_tag))
    cont2 = reflist2.find(_qn_hh(container_tag))
    if cont1 is None or cont2 is None:
        return {}

    offset = len(cont1)
    id_map = {}
    for child in list(cont2):
        old_id = int(child.get("id"))
        new_id = old_id + offset
        id_map[old_id] = new_id
        new_child = copy.deepcopy(child)
        new_child.set("id", str(new_id))
        cont1.append(new_child)

    return id_map


def _merge_fontfaces(reflist1, reflist2):
    """fontfaces는 언어별로 그룹화되어 있어 언어마다 독립적인 offset을 가짐.
    return: {lang_key: {old_id: new_id}}"""
    ff1 = reflist1.find(_qn_hh("fontfaces"))
    ff2 = reflist2.find(_qn_hh("fontfaces"))
    if ff1 is None or ff2 is None:
        return {}

    lang_groups1 = {fg.get("lang"): fg for fg in ff1}
    lang_id_maps = {}

    for fg2 in list(ff2):
        lang = fg2.get("lang")
        fg1 = lang_groups1.get(lang)
        if fg1 is None:
            ff1.append(copy.deepcopy(fg2))
            lang_key = LANG_KEY_MAP.get(lang)
            if lang_key:
                lang_id_maps[lang_key] = {int(f.get("id")): int(f.get("id")) for f in fg2}
            continue

        offset = len(fg1)
        id_map = {}
        for font in list(fg2):
            old_id = int(font.get("id"))
            new_id = old_id + offset
            id_map[old_id] = new_id
            new_font = copy.deepcopy(font)
            new_font.set("id", str(new_id))
            fg1.append(new_font)
        fg1.set("fontCnt", str(len(fg1)))

        lang_key = LANG_KEY_MAP.get(lang)
        if lang_key:
            lang_id_maps[lang_key] = id_map

    return lang_id_maps


def _apply_fontref_remap(root, lang_id_maps):
    for fontref in root.iter(_qn_hh("fontRef")):
        for lang_key, id_map in lang_id_maps.items():
            val = fontref.get(lang_key)
            if val is not None and val.isdigit():
                old_id = int(val)
                if old_id in id_map:
                    fontref.set(lang_key, str(id_map[old_id]))


def _apply_idref_remap(root, attr_name, id_map):
    if not id_map:
        return
    for el in root.iter():
        val = el.get(attr_name)
        if val is not None and val.isdigit():
            old_id = int(val)
            if old_id in id_map:
                el.set(attr_name, str(id_map[old_id]))


def _remap_heading_idref(root, style_id_map):
    """paraPr 안의 <heading idRef="..."> 는 style을 참조하는 특수 케이스."""
    for heading in root.iter(_qn_hh("heading")):
        val = heading.get("idRef")
        if val is not None and val.isdigit():
            old_id = int(val)
            if old_id in style_id_map:
                heading.set("idRef", str(style_id_map[old_id]))


def _merge_header_xml(header1_bytes, header2_bytes, section2_root):
    """header1에 header2의 스타일 정의를 합치고, section2_root(이미 파싱된 트리)의
    IDRef들을 새 id로 치환한다. return: 병합된 header1 트리(etree)"""
    parser = etree.XMLParser(remove_blank_text=False)
    tree1 = etree.parse(io.BytesIO(header1_bytes), parser)
    tree2 = etree.parse(io.BytesIO(header2_bytes), parser)
    root1 = tree1.getroot()
    root2 = tree2.getroot()

    reflist1 = _get_reflist(root1)
    reflist2 = _get_reflist(root2)

    all_id_maps = {}
    for container in SIMPLE_CONTAINERS:
        id_map = _merge_simple_container(reflist1, reflist2, container)
        all_id_maps[container] = id_map

    lang_id_maps = _merge_fontfaces(reflist1, reflist2)

    # header2 원본(복사 전 root2)에 새 id 기준으로 내부 참조를 먼저 갱신해둔다.
    for container, attrs in IDREF_FOR_CONTAINER.items():
        id_map = all_id_maps.get(container, {})
        for attr in attrs:
            _apply_idref_remap(root2, attr, id_map)
    _apply_fontref_remap(root2, lang_id_maps)
    _remap_heading_idref(root2, all_id_maps.get("styles", {}))

    # 위 갱신은 root2 자체에 적용한 것이라, 이미 reflist1에 복사된(새 id가 붙은) 항목들의
    # 내부 참조도 동일하게 새 id로 맞춰야 한다.
    for container, attrs in IDREF_FOR_CONTAINER.items():
        id_map = all_id_maps.get(container, {})
        cont1 = reflist1.find(_qn_hh(container))
        if cont1 is None or not id_map:
            continue
        new_ids = set(id_map.values())
        for child in cont1:
            if int(child.get("id")) in new_ids:
                for attr in attrs:
                    _apply_idref_remap(child, attr, id_map)
                _apply_fontref_remap(child, lang_id_maps)
                _remap_heading_idref(child, all_id_maps.get("styles", {}))

    # 이어붙일 본문(section2_root)에도 동일 리맵 적용
    for container, attrs in IDREF_FOR_CONTAINER.items():
        id_map = all_id_maps.get(container, {})
        for attr in attrs:
            _apply_idref_remap(section2_root, attr, id_map)
    _apply_fontref_remap(section2_root, lang_id_maps)
    _remap_heading_idref(section2_root, all_id_maps.get("styles", {}))

    return tree1


def _collect_image_items(content_hpf_root):
    """content.hpf의 manifest에서 BinData 이미지 항목들을 리스트로 추출."""
    items = []
    manifest = content_hpf_root.find(f"{{{NS_OPF}}}manifest")
    if manifest is None:
        return items
    for item in manifest:
        href = item.get("href", "")
        if href.startswith("BinData/"):
            items.append(item)
    return items


def _renumber_images(zf2, content_hpf2_root, section2_root, image_offset):
    """두번째 파일의 BinData 이미지들을 새 파일명(image{offset+1}.ext 부터)으로 매핑하고,
    content.hpf의 id/href와 section0.xml의 binaryItemIDRef를 갱신한다.
    return: {new_filename: bytes} 딕셔너리"""
    items = _collect_image_items(content_hpf2_root)
    id_str_map = {}
    new_files = {}

    for idx, item in enumerate(items, start=1):
        old_id = item.get("id")
        href = item.get("href")
        ext = href.rsplit(".", 1)[-1] if "." in href else "bin"
        new_id = f"image{image_offset + idx}"
        new_href = f"BinData/{new_id}.{ext}"

        try:
            data = zf2.read(href)
        except KeyError:
            continue
        new_files[new_href] = data
        id_str_map[old_id] = new_id

        item.set("id", new_id)
        item.set("href", new_href)

    for el in section2_root.iter():
        val = el.get("binaryItemIDRef")
        if val is not None and val in id_str_map:
            el.set("binaryItemIDRef", id_str_map[val])

    return new_files


def merge_two_hwpx(bytes1, bytes2):
    """hwpx 파일 두 개(bytes)를 받아, bytes1의 끝에 bytes2의 본문을 이어붙인
    새 hwpx 파일을 bytes로 반환한다."""
    zf1 = zipfile.ZipFile(io.BytesIO(bytes1))
    zf2 = zipfile.ZipFile(io.BytesIO(bytes2))

    parser = etree.XMLParser(remove_blank_text=False)

    section1_root = etree.parse(io.BytesIO(zf1.read("Contents/section0.xml")), parser).getroot()
    section2_root = etree.parse(io.BytesIO(zf2.read("Contents/section0.xml")), parser).getroot()

    content_hpf1_root = etree.parse(io.BytesIO(zf1.read("Contents/content.hpf")), parser).getroot()
    content_hpf2_root = etree.parse(io.BytesIO(zf2.read("Contents/content.hpf")), parser).getroot()

    # 1) 이미지 파일 재번호 부여 (파일1의 기존 이미지 개수만큼 offset)
    existing_image_count = len(_collect_image_items(content_hpf1_root))
    new_image_files = _renumber_images(zf2, content_hpf2_root, section2_root, existing_image_count)

    # 2) header.xml 병합 (스타일 정의 + section2의 IDRef 갱신)
    header1_bytes = zf1.read("Contents/header.xml")
    header2_bytes = zf2.read("Contents/header.xml")
    merged_header_tree = _merge_header_xml(header1_bytes, header2_bytes, section2_root)

    # 3) section0.xml 본문 이어붙이기 (file2의 모든 단락을 file1 끝에 추가)
    for p in list(section2_root):
        section1_root.append(copy.deepcopy(p))

    # 4) content.hpf의 manifest에 새 이미지 항목 추가
    manifest1 = content_hpf1_root.find(f"{{{NS_OPF}}}manifest")
    for item in _collect_image_items(content_hpf2_root):
        manifest1.append(copy.deepcopy(item))

    # 5) 새 zip 작성 (file1의 다른 모든 항목은 그대로 유지)
    out_buffer = io.BytesIO()
    with zipfile.ZipFile(out_buffer, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zf1.namelist():
            if name == "Contents/section0.xml":
                zout.writestr(name, etree.tostring(section1_root.getroottree(),
                                                    xml_declaration=True, encoding="UTF-8", standalone=True))
            elif name == "Contents/header.xml":
                zout.writestr(name, etree.tostring(merged_header_tree,
                                                    xml_declaration=True, encoding="UTF-8", standalone=True))
            elif name == "Contents/content.hpf":
                zout.writestr(name, etree.tostring(content_hpf1_root.getroottree(),
                                                    xml_declaration=True, encoding="UTF-8", standalone=True))
            else:
                zout.writestr(name, zf1.read(name))

        for new_href, data in new_image_files.items():
            zout.writestr(new_href, data)

    zf1.close()
    zf2.close()
    out_buffer.seek(0)
    return out_buffer.read()


def merge_hwpx_files(file_bytes_list):
    """hwpx 파일 bytes 리스트를 순서대로 모두 병합한 결과 bytes를 반환."""
    if not file_bytes_list:
        raise ValueError("병합할 파일이 없습니다")
    result = file_bytes_list[0]
    for next_bytes in file_bytes_list[1:]:
        result = merge_two_hwpx(result, next_bytes)
    return result


def sort_key_for_hwpx_filename(filename, fallback_index):
    """파일명에서 순서를 결정할 단서(앞쪽 숫자, 또는 'NN_지역명' 패턴)를 찾는다.
    단서가 전혀 없으면 업로드된 순서(fallback_index)를 그대로 사용해
    '무작위로 한 파일씩 붙여넣기'가 아니라 업로드 순서를 보존한다."""
    m = re.match(r'^\D*(\d{1,3})', filename)
    if m:
        return (0, int(m.group(1)), fallback_index)
    return (1, 0, fallback_index)


def render():
    """탭4 화면을 그린다. app.py에서 with tab4: render() 형태로 호출."""
    st.caption(
        "여러 한글(hwpx) 파일을 순서대로 하나로 이어붙입니다. "
        "한글 프로그램을 직접 실행하지 않고, 파일 구조를 직접 다뤄 병합하므로 "
        "클라우드(Linux) 환경에서도 정상 동작합니다."
    )
    st.info(
        "📌 파일명에 순서를 나타내는 숫자(예: 01_..., 1_...)가 있으면 그 순서대로 병합합니다.\n\n"
        "📌 숫자가 없으면 업로드한 순서 그대로 한 파일씩 이어붙입니다.\n\n"
        "📌 .hwp(구버전)는 지원하지 않습니다 — .hwpx로 저장된 파일만 가능합니다."
    )

    hwpx_files = st.file_uploader(
        "한글 파일 업로드 (여러 개 선택 가능)", type=["hwpx"], accept_multiple_files=True, key="hwpx_up4"
    )

    if hwpx_files and st.button("🚀 한글 파일 병합 시작", key="btn4"):
        try:
            indexed_files = list(enumerate(hwpx_files))
            indexed_files.sort(key=lambda x: sort_key_for_hwpx_filename(x[1].name, x[0]))
            sorted_files = [f for _, f in indexed_files]

            st.caption("병합 순서: " + " → ".join(f.name for f in sorted_files))

            file_bytes_list = [f.read() for f in sorted_files]
            merged_bytes = merge_hwpx_files(file_bytes_list)

            st.success("한글 파일 병합이 완료되었습니다.")
            st.download_button(
                "📥 다운로드", merged_bytes, "병합_결과.hwpx",
                mime="application/octet-stream", key="dl4"
            )

        except Exception as e:
            st.error(f"오류: {e}")
            st.exception(e)
