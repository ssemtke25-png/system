# ══════════════════════════════════════════════════════════════════════
#  탭7 행사장 약도 - 지도 함수 교체본 (2026-08)
#  변경: ① 이미지 하단 출처(attribution) 자동 표기
#        ② ESRI 위성 유지 + 출처 "Esri, Earthstar Geographics"
#        ③ VWorld 운영키 확보 시 URL만 교체하면 되도록 구조화
#
#  ▶ 교체 대상: 기존 _draw_marker / _osm_static_via_tiles / _render_map_image
#     이 세 함수를 아래 내용으로 통째로 바꾸면 됩니다.
# ══════════════════════════════════════════════════════════════════════


def _draw_marker(canvas, cx, cy):
    """PIL 이미지 중심에 빨간 핀 마커를 그린다."""
    from PIL import ImageDraw
    draw = ImageDraw.Draw(canvas)
    r = 10
    draw.ellipse([cx - r, cy - r, cx + r, cy + r],
                 fill=(214, 40, 40), outline=(255, 255, 255), width=3)
    draw.ellipse([cx - 3, cy - 3, cx + 3, cy + 3], fill=(255, 255, 255))


def _draw_attribution(canvas, text):
    """
    이미지 우측 하단에 출처 문구를 반투명 배경 위에 얹는다.
    지도 서비스 약관의 attribution(출처 표기) 의무 충족용.
    폰트 로딩 실패 시 기본 폰트로 폴백하므로 어느 환경에서도 안전.
    """
    from PIL import ImageDraw, ImageFont
    if not text:
        return

    draw = ImageDraw.Draw(canvas, "RGBA")
    W, H = canvas.size

    # 폰트 탐색: 한글(CJK) 폰트를 최우선 → 없으면 영문 폰트 → 최후엔 기본 비트맵.
    # 한글 출처(예: VWorld 전환 시 "출처: 국토교통부 브이월드")도 깨지지 않도록.
    # 한글 폰트를 못 찾고 문구에 한글이 있으면 영문 라벨로 자동 대체한다.
    kr_fonts = (
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "C:/Windows/Fonts/malgun.ttf",  # 로컬 exe(Windows) 환경
    )
    en_fonts = (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/arial.ttf",
    )
    font, is_kr_font = None, False
    for fp in kr_fonts:
        try:
            font = ImageFont.truetype(fp, 12); is_kr_font = True; break
        except Exception:
            continue
    if font is None:
        for fp in en_fonts:
            try:
                font = ImageFont.truetype(fp, 12); break
            except Exception:
                continue
    if font is None:
        font = ImageFont.load_default()

    # 한글 폰트가 없는데 문구에 한글이 섞여 있으면 → 영문 표기로 안전 대체
    has_hangul = any("\uac00" <= ch <= "\ud7a3" for ch in text)
    if has_hangul and not is_kr_font:
        _FALLBACK_EN = {
            "출처: Esri, Earthstar Geographics": "Source: Esri, Earthstar Geographics",
            "출처: OpenStreetMap contributors": "Source: OpenStreetMap contributors",
            "출처: 국토교통부 브이월드(VWorld)": "Source: VWorld (MOLIT, Korea)",
        }
        text = _FALLBACK_EN.get(text, "Source: map data")

    # 텍스트 박스 크기 측정 (Pillow 버전별 호환)
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    except AttributeError:
        tw, th = draw.textsize(text, font=font)

    pad = 4
    x1, y1 = W - tw - pad * 2, H - th - pad * 2
    # 반투명 검은 배경띠
    draw.rectangle([x1, y1, W, H], fill=(0, 0, 0, 130))
    draw.text((x1 + pad, y1 + pad), text, fill=(255, 255, 255, 235), font=font)


# ── 지도 소스 정의 ────────────────────────────────────────────────────
# VWorld 운영키를 받으면 아래 블록만 교체하면 됩니다.
# (st.secrets["VWORLD_KEY"] 추가 후 satellite/street url을 vworld로 변경)
#
#   VWORLD 예시:
#     key = st.secrets["VWORLD_KEY"]
#     satellite = f"https://api.vworld.kr/req/wmts/1.0.0/{key}/Satellite/{{z}}/{{y}}/{{x}}.jpeg"
#     overlay   = f"https://api.vworld.kr/req/wmts/1.0.0/{key}/Hybrid/{{z}}/{{y}}/{{x}}.png"
#     street    = f"https://api.vworld.kr/req/wmts/1.0.0/{key}/Base/{{z}}/{{y}}/{{x}}.png"
#     attribution = "출처: 국토교통부 브이월드(VWorld)"
#
def _tile_sources(maptype):
    """
    (base_url, overlay_url, attribution) 반환.
    overlay_url 은 None 가능. attribution 은 이미지 하단에 찍힘.
    """
    if maptype == "satellite":
        return (
            "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
            "https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}",
            "출처: Esri, Earthstar Geographics",
        )
    else:
        return (
            "https://tile.openstreetmap.org/{z}/{x}/{y}.png",
            None,
            "출처: OpenStreetMap contributors",
        )


def _osm_static_via_tiles(lat, lng, slider_zoom, width, height, maptype="satellite"):
    """
    타일 서버에서 타일을 직접 내려받아 합성해 PNG 생성.
    이미지 하단에 출처(attribution) 자동 표기.
    slider_zoom(1~14, 클수록 확대) → 표준 줌으로 매핑.
    maptype: "satellite"(ESRI 위성) 또는 "street"(OSM 일반).
    실패 시 None 반환.
    """
    from PIL import Image
    import math

    osm_zoom = max(3, min(18, int(slider_zoom) + 6))
    TILE = 256

    tile_url, overlay_url, attribution = _tile_sources(maptype)

    def deg2num(lat_deg, lon_deg, z):
        lat_rad = math.radians(lat_deg)
        n = 2.0 ** z
        xf = (lon_deg + 180.0) / 360.0 * n
        yf = (1.0 - math.asinh(math.tan(lat_rad)) / math.pi) / 2.0 * n
        return xf, yf

    xf, yf = deg2num(lat, lng, osm_zoom)
    center_px = xf * TILE
    center_py = yf * TILE

    left = center_px - width / 2
    top  = center_py - height / 2
    x0 = int(math.floor(left / TILE))
    y0 = int(math.floor(top / TILE))
    x1 = int(math.floor((left + width) / TILE))
    y1 = int(math.floor((top + height) / TILE))

    canvas = Image.new("RGB", (width, height), (233, 229, 220))
    n_tiles = 2 ** osm_zoom
    headers = {"User-Agent": "GyeongbukEventMap/1.0 (public admin tool; contact: land-info)"}

    def _fetch(url):
        try:
            tr = requests.get(url, headers=headers, timeout=8)
            if tr.status_code != 200 or "image" not in tr.headers.get("Content-Type", ""):
                return None
            return Image.open(io.BytesIO(tr.content)).convert("RGBA")
        except Exception:
            return None

    got_any = False
    for tx in range(x0, x1 + 1):
        for ty in range(y0, y1 + 1):
            if not (0 <= ty < n_tiles):
                continue
            txx = tx % n_tiles  # 경도 wrap
            base = _fetch(tile_url.format(z=osm_zoom, x=txx, y=ty))
            if base is None:
                continue
            if overlay_url:
                ov = _fetch(overlay_url.format(z=osm_zoom, x=txx, y=ty))
                if ov is not None:
                    base.alpha_composite(ov)
            paste_x = int(tx * TILE - left)
            paste_y = int(ty * TILE - top)
            canvas.paste(base.convert("RGB"), (paste_x, paste_y))
            got_any = True

    if not got_any:
        return None

    _draw_marker(canvas, width // 2, height // 2)
    _draw_attribution(canvas, attribution)   # ← 출처 표기
    out = io.BytesIO()
    canvas.save(out, format="PNG")
    return out.getvalue()


def _render_map_image(lat, lng, zoom, mw, mh, maptype):
    """
    좌표+옵션으로 지도 이미지 바이트 생성.
    위성: ESRI 타일 합성(출처표기) → 실패 시 OSM 일반
    일반: 카카오 정적지도 → 실패 시 OSM 타일 합성(출처표기)
    반환: (img_bytes, source_note) 또는 (None, "")
    """
    kkey = st.secrets["KAKAO_API_KEY"]
    hdrs = {"Authorization": f"KakaoAK {kkey}"}
    img_bytes = None
    source_note = ""

    if maptype == "satellite":
        try:
            img_bytes = _osm_static_via_tiles(lat, lng, zoom, int(mw), int(mh),
                                              maptype="satellite")
            if img_bytes:
                source_note = "※ 위성영상(Esri) + 지명 오버레이 · 출처 이미지 내 표기"
        except Exception:
            img_bytes = None
        if img_bytes is None:
            try:
                img_bytes = _osm_static_via_tiles(lat, lng, zoom, int(mw), int(mh),
                                                  maptype="street")
                if img_bytes:
                    source_note = "※ 위성 실패 → OpenStreetMap 일반지도로 대체"
            except Exception:
                img_bytes = None
    else:
        # 카카오 정적지도는 일반지도(ROADMAP)만 지원 → 위성은 위 분기에서 처리
        try:
            mr = requests.get("https://dapi.kakao.com/v2/maps/staticmap", headers=hdrs,
                              params={"center": f"{lng},{lat}", "level": zoom,
                                      "w": int(mw), "h": int(mh),
                                      "markers": f"color:red|{lng},{lat}"},
                              timeout=10)
            if mr.status_code == 200 and "image" in mr.headers.get("Content-Type", ""):
                img_bytes = mr.content
                source_note = "※ 카카오맵 기반"
        except Exception:
            img_bytes = None
        if img_bytes is None:
            try:
                img_bytes = _osm_static_via_tiles(lat, lng, zoom, int(mw), int(mh),
                                                  maptype="street")
                if img_bytes:
                    source_note = "※ OpenStreetMap 타일 기반"
            except Exception:
                img_bytes = None

    return img_bytes, source_note
