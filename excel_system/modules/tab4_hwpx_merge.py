"""
탭4: 한글(HWPX) 파일 병합 (v2 - 다중 섹션 지원)

핵심 발견 및 수정:
- hwpx 문서는 본문이 항상 section0.xml 하나에만 있는 게 아니라, 한글의
  '구역 나누기' 기능을 쓰면 section0.xml, section1.xml, section2.xml...
  처럼 여러 섹션 파일로 나뉠 수 있다. 어떤 섹션이 몇 개 있고 어떤 순서로
  읽어야 하는지는 content.hpf의 manifest/spine에 정의되어 있다.
- v1은 section0.xml 하나만 읽고 썼기 때문에, 구역이 여러 개로 나뉜
  파일(특히 다른 도구로 병합되었거나 페이지마다 레이아웃이 다른 문서)을
  만나면 section1.xml 이후의 내용이 통째로 사라지는 문제가 있었다.
- v2는 content.hpf의 spine을 따라 모든 섹션을 찾아서, 파일1의 섹션들
  뒤에 파일2의 섹션들을 그대로 이어붙인다(섹션을 억지로 하나로 합치지
  않고, 구역 구조를 그대로 보존하는 것이 가장 안전하다).

순서 판단:
- 파일명에 시군명이나 숫자(01, 02... 또는 1_, 2_)가 있으면 그 순서대로.
- 그런 단서가 전혀 없으면 사용자가 업로드한 순서 그대로 한 파일씩 이어붙인다.

처리 항목:
- header.xml의 스타일 정의(charProperties, paraProperties, styles, borderFills,
  tabProperties, numberings, bullets, fontfaces)를 모두 합치고, 새로 부여된 id로
  모든 섹션과 header.xml 안의 IDRef 참조를 일괄 치환한다.
- BinData(이미지) 파일명이 서로 겹치지 않도록 번호를 이어서 재명명하고,
  content.hpf의 이미지 등록 항목과 각 섹션의 binaryItemIDRef도 같이 갱신한다.
- 각 파일의 모든 섹션을 번호를 다시 매겨(section0, section1, section2...)
  순서대로 이어붙이고, content.hpf의 manifest/spine을 새로 구성한다.
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
NS_HPF = "http://www.hancom.co.kr/schema/2011/hpf"

SIMPLE_CONTAINERS = ["borderFills", "charProperties", "tabProperties",
                      "numberings", "paraProperties", "styles", "bullets"]

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


# =========================================================================
# 공통 헬퍼
# =========================================================================
def _qn_hh(tag):
    return f"{{{NS_HH}}}{tag}"


def _real_elements(container):
    """lxml은 XML 주석도 컨테이너의 자식으로 순회하는데, 주석은 id 속성이
    없어 int(child.get('id'))가 실패한다. 실제 운영 파일에는 안내용 주석이
    자주 섞여 있으므로, 진짜 정의 요소(주석이 아닌 것)만 걸러서 반환한다."""
    if container is None:
        return []
    return [el for el in container if isinstance(el.tag, str)]


def _find_opf_root_ns(content_hpf_root):
    """content.hpf는 한글 버전에 따라 루트가 opf:package 또는 hpf:package로
    저장되는 두 가지 변형이 실제로 존재한다. 실제 쓰인 네임스페이스를 찾는다."""
    tag = content_hpf_root.tag
    if tag == f"{{{NS_OPF}}}package":
        return NS_OPF
    if tag == f"{{{NS_HPF}}}package":
        return NS_HPF
    return NS_OPF


def _find_manifest(content_hpf_root):
    for ns in (NS_OPF, NS_HPF):
        manifest = content_hpf_root.find(f"{{{ns}}}manifest")
        if manifest is not None:
            return manifest
    return None


def _find_spine(content_hpf_root):
    for ns in (NS_OPF, NS_HPF):
        spine = content_hpf_root.find(f"{{{ns}}}spine")
        if spine is not None:
            return spine
    return None


# =========================================================================
# 섹션(section0.xml, section1.xml, ...) 목록 추출
# =========================================================================
def _get_section_hrefs_in_order(content_hpf_root):
    """content.hpf의 manifest+spine을 따라가서, 본문 섹션 파일들의 zip 내부
    경로를 spine에 정의된 순서대로 리스트로 반환. 예: ['Contents/section0.xml', ...]
    spine이 없거나 비정상이면 manifest에서 'section'으로 시작하는 id를 정렬해 사용."""
    manifest = _find_manifest(content_hpf_root)
    spine = _find_spine(content_hpf_root)

    id_to_href = {}
    if manifest is not None:
        for item in _real_elements(manifest):
            iid = item.get("id")
            href = item.get("href")
            if iid and href:
                id_to_href[iid] = href

    ordered_hrefs = []
    if spine is not None:
        for ref in _real_elements(spine):
            idref = ref.get("idref")
            href = id_to_href.get(idref)
            if href and "section" in idref:
                ordered_hrefs.append(href)

    if not ordered_hrefs:
        # spine이 비정상일 때를 위한 안전망: id가 'section'으로 시작하는 것들을
        # 숫자 순서로 정렬해서 사용한다.
        section_items = [(iid, href) for iid, href in id_to_href.items() if iid.startswith("section")]
        section_items.sort(key=lambda x: int(re.sub(r'\D', '', x[0]) or 0))
        ordered_hrefs = [href for _, href in section_items]

    # href가 'Contents/section0.xml' 또는 'section0.xml' 형태로 저장될 수 있어 보정
    return [href if href.startswith("Contents/") else f"Contents/{href}" for href in ordered_hrefs]


def _normalize_href(href):
    return href if href.startswith("Contents/") else f"Contents/{href}"


# =========================================================================
# header.xml 스타일 정의 병합 (charProperties, fontfaces 등)
# =========================================================================
def _get_reflist(header_root):
    return header_root.find(_qn_hh("refList"))


def _merge_simple_container(reflist1, reflist2, container_tag):
    cont1 = reflist1.find(_qn_hh(container_tag))
    cont2 = reflist2.find(_qn_hh(container_tag))
    if cont1 is None or cont2 is None:
        return {}

    offset = len(_real_elements(cont1))
    id_map = {}
    for child in _real_elements(cont2):
        cid = child.get("id")
        if cid is None:
            continue
        old_id = int(cid)
        new_id = old_id + offset
        id_map[old_id] = new_id
        new_child = copy.deepcopy(child)
        new_child.set("id", str(new_id))
        cont1.append(new_child)

    return id_map


def _merge_fontfaces(reflist1, reflist2):
    ff1 = reflist1.find(_qn_hh("fontfaces"))
    ff2 = reflist2.find(_qn_hh("fontfaces"))
    if ff1 is None or ff2 is None:
        return {}

    lang_groups1 = {fg.get("lang"): fg for fg in _real_elements(ff1)}
    lang_id_maps = {}

    for fg2 in _real_elements(ff2):
        lang = fg2.get("lang")
        fg1 = lang_groups1.get(lang)
        if fg1 is None:
            ff1.append(copy.deepcopy(fg2))
            lang_key = LANG_KEY_MAP.get(lang)
            if lang_key:
                lang_id_maps[lang_key] = {
                    int(f.get("id")): int(f.get("id"))
                    for f in _real_elements(fg2) if f.get("id") is not None
                }
            continue

        offset = len(_real_elements(fg1))
        id_map = {}
        for font in _real_elements(fg2):
            fid = font.get("id")
            if fid is None:
                continue
            old_id = int(fid)
            new_id = old_id + offset
            id_map[old_id] = new_id
            new_font = copy.deepcopy(font)
            new_font.set("id", str(new_id))
            fg1.append(new_font)
        fg1.set("fontCnt", str(len(_real_elements(fg1))))

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
    for heading in root.iter(_qn_hh("heading")):
        val = heading.get("idRef")
        if val is not None and val.isdigit():
            old_id = int(val)
            if old_id in style_id_map:
                heading.set("idRef", str(style_id_map[old_id]))


def _merge_header_xml(header1_bytes, header2_bytes, section2_roots):
    """header1에 header2의 스타일 정의를 합치고, section2_roots(파일2에 속한
    모든 섹션 트리 리스트)의 IDRef들을 새 id로 치환한다.
    return: 병합된 header1 트리(etree)"""
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

    for container, attrs in IDREF_FOR_CONTAINER.items():
        id_map = all_id_maps.get(container, {})
        for attr in attrs:
            _apply_idref_remap(root2, attr, id_map)
    _apply_fontref_remap(root2, lang_id_maps)
    _remap_heading_idref(root2, all_id_maps.get("styles", {}))

    for container, attrs in IDREF_FOR_CONTAINER.items():
        id_map = all_id_maps.get(container, {})
        cont1 = reflist1.find(_qn_hh(container))
        if cont1 is None or not id_map:
            continue
        new_ids = set(id_map.values())
        for child in _real_elements(cont1):
            cid = child.get("id")
            if cid is None:
                continue
            if int(cid) in new_ids:
                for attr in attrs:
                    _apply_idref_remap(child, attr, id_map)
                _apply_fontref_remap(child, lang_id_maps)
                _remap_heading_idref(child, all_id_maps.get("styles", {}))

    # 파일2에 속한 모든 섹션에 동일 리맵 적용 (섹션이 여러 개여도 전부 처리)
    for section_root in section2_roots:
        for container, attrs in IDREF_FOR_CONTAINER.items():
            id_map = all_id_maps.get(container, {})
            for attr in attrs:
                _apply_idref_remap(section_root, attr, id_map)
        _apply_fontref_remap(section_root, lang_id_maps)
        _remap_heading_idref(section_root, all_id_maps.get("styles", {}))

    return tree1


# =========================================================================
# 이미지(BinData) 재번호
# =========================================================================
def _collect_image_items(content_hpf_root):
    items = []
    manifest = _find_manifest(content_hpf_root)
    if manifest is None:
        return items
    for item in _real_elements(manifest):
        href = item.get("href", "")
        if href.startswith("BinData/"):
            items.append(item)
    return items


def _renumber_images(zf2, content_hpf2_root, section2_roots, image_offset):
    """두번째 파일의 BinData 이미지들을 새 파일명으로 매핑하고,
    content.hpf의 id/href와 모든 관련 섹션의 binaryItemIDRef를 갱신한다."""
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

    for section_root in section2_roots:
        for el in section_root.iter():
            val = el.get("binaryItemIDRef")
            if val is not None and val in id_str_map:
                el.set("binaryItemIDRef", id_str_map[val])

    return new_files


# =========================================================================
# zOrder / 개체 id 충돌 방지
# =========================================================================
def _renumber_zorder_and_objid(section1_roots, section2_roots):
    """표(tbl), 그림(pic) 등 'zOrder'와 'id' 속성을 가진 개체들은 문서 전체에서
    고유해야 한다. file1에 속한 모든 섹션을 기준으로 이미 쓰인 zOrder/id를
    모으고, file2에 속한 모든 섹션의 개체들에 충돌 없는 새 값을 부여한다."""
    object_tags = ["tbl", "pic", "container", "ole", "equation"]

    def collect_objects(roots):
        objs = []
        for root in roots:
            for tag in object_tags:
                objs.extend(root.iter(f"{{{NS_HP}}}{tag}"))
        return objs

    objs1 = collect_objects(section1_roots)
    objs2 = collect_objects(section2_roots)

    existing_zorders = [
        int(o.get("zOrder")) for o in objs1
        if o.get("zOrder") is not None and o.get("zOrder").lstrip('-').isdigit()
    ]
    existing_ids = {o.get("id") for o in objs1 if o.get("id")}

    next_zorder = (max(existing_zorders) + 1) if existing_zorders else 0
    next_id = 1

    for obj in objs2:
        if obj.get("zOrder") is not None:
            obj.set("zOrder", str(next_zorder))
            next_zorder += 1
        if obj.get("id"):
            new_id = obj.get("id")
            while new_id in existing_ids:
                new_id = str(int(new_id) + 1) if new_id.isdigit() else f"{new_id}_{next_id}"
                next_id += 1
            obj.set("id", new_id)
            existing_ids.add(new_id)


# =========================================================================
# 두 hwpx 병합 (다중 섹션 지원)
# =========================================================================
def merge_two_hwpx(bytes1, bytes2):
    """hwpx 파일 두 개(bytes)를 받아, bytes1의 모든 섹션 뒤에 bytes2의 모든
    섹션을 순서대로 이어붙인 새 hwpx 파일을 bytes로 반환한다."""
    zf1 = zipfile.ZipFile(io.BytesIO(bytes1))
    zf2 = zipfile.ZipFile(io.BytesIO(bytes2))

    parser = etree.XMLParser(remove_blank_text=False)

    content_hpf1_root = etree.parse(io.BytesIO(zf1.read("Contents/content.hpf")), parser).getroot()
    content_hpf2_root = etree.parse(io.BytesIO(zf2.read("Contents/content.hpf")), parser).getroot()

    section1_hrefs = _get_section_hrefs_in_order(content_hpf1_root)
    section2_hrefs = _get_section_hrefs_in_order(content_hpf2_root)

    if not section1_hrefs:
        section1_hrefs = ["Contents/section0.xml"]
    if not section2_hrefs:
        section2_hrefs = ["Contents/section0.xml"]

    section1_roots = [etree.parse(io.BytesIO(zf1.read(h)), parser).getroot() for h in section1_hrefs]
    section2_roots = [etree.parse(io.BytesIO(zf2.read(h)), parser).getroot() for h in section2_hrefs]

    # 1) 이미지 파일 재번호 부여
    existing_image_count = len(_collect_image_items(content_hpf1_root))
    new_image_files = _renumber_images(zf2, content_hpf2_root, section2_roots, existing_image_count)

    # 2) header.xml 병합 (스타일 정의 + 파일2의 모든 섹션 IDRef 갱신)
    header1_bytes = zf1.read("Contents/header.xml")
    header2_bytes = zf2.read("Contents/header.xml")
    merged_header_tree = _merge_header_xml(header1_bytes, header2_bytes, section2_roots)

    # 3) zOrder/개체 id 충돌 방지 (파일1의 모든 섹션 기준으로 파일2의 모든 섹션을 재배치)
    _renumber_zorder_and_objid(section1_roots, section2_roots)

    # 4) 전체 섹션 순서 = 파일1의 섹션들 + 파일2의 섹션들, 번호를 0부터 다시 매김
    all_section_roots = section1_roots + section2_roots
    new_section_names = [f"Contents/section{i}.xml" for i in range(len(all_section_roots))]

    # 5) content.hpf의 manifest/spine을 새 섹션 구성에 맞게 재작성
    opf_ns = _find_opf_root_ns(content_hpf1_root)
    manifest1 = _find_manifest(content_hpf1_root)
    spine1 = _find_spine(content_hpf1_root)

    if manifest1 is not None:
        # 기존 section 항목들 제거 후 새로 등록
        for item in list(_real_elements(manifest1)):
            iid = item.get("id")
            if iid and iid.startswith("section"):
                manifest1.remove(item)
        for i, name in enumerate(new_section_names):
            new_item = etree.SubElement(manifest1, f"{{{opf_ns}}}item")
            new_item.set("id", f"section{i}")
            new_item.set("href", name)
            new_item.set("media-type", "application/xml")

        # 새 이미지 항목 추가
        for item in _collect_image_items(content_hpf2_root):
            manifest1.append(copy.deepcopy(item))

    if spine1 is not None:
        for ref in list(_real_elements(spine1)):
            idref = ref.get("idref")
            if idref and idref.startswith("section"):
                spine1.remove(ref)
        # header 다음에 섹션들을 순서대로 등록 (header가 먼저 나오는 기존 관례 유지)
        header_ref_exists = any(
            r.get("idref") == "header" for r in _real_elements(spine1)
        )
        insert_pos = 1 if header_ref_exists else 0
        for i in range(len(new_section_names)):
            new_ref = etree.Element(f"{{{opf_ns}}}itemref")
            new_ref.set("idref", f"section{i}")
            new_ref.set("linear", "yes")
            spine1.insert(insert_pos + i, new_ref)

    # 6) 새 zip 작성
    out_buffer = io.BytesIO()
    with zipfile.ZipFile(out_buffer, "w", zipfile.ZIP_DEFLATED) as zout:
        written = set()

        # file1의 항목들 중 섹션이 아닌 것은 그대로, 섹션/헤더/content.hpf는 갱신본으로
        for name in zf1.namelist():
            if re.match(r"Contents/section\d+\.xml$", name):
                continue  # 아래에서 새 섹션들로 일괄 작성
            if name == "Contents/header.xml":
                zout.writestr(name, etree.tostring(merged_header_tree,
                                                    xml_declaration=True, encoding="UTF-8", standalone=True))
            elif name == "Contents/content.hpf":
                zout.writestr(name, etree.tostring(content_hpf1_root.getroottree(),
                                                    xml_declaration=True, encoding="UTF-8", standalone=True))
            else:
                zout.writestr(name, zf1.read(name))
            written.add(name)

        # 새 섹션들 작성
        for name, root in zip(new_section_names, all_section_roots):
            zout.writestr(name, etree.tostring(root.getroottree(),
                                                xml_declaration=True, encoding="UTF-8", standalone=True))

        # 새 이미지 파일들 추가
        for new_href, data in new_image_files.items():
            if new_href not in written:
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
