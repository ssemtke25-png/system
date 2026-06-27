"""
공통 유틸 모듈.
탭1(단순 합산), 탭2(총괄표 채우기), 탭3(실거래 월보)에서 공유하는
지역명 정규화, 숫자/수식 판별 함수들을 모아둔다.
탭4(한글 파일 병합)는 지역/숫자 개념이 없어 이 모듈을 쓰지 않는다.
"""
import re

PREFIX_SPECIAL = {
    '포항시남구': '포항남', '포항남구': '포항남', '포항남': '포항남',
    '포항시북구': '포항북', '포항북구': '포항북', '포항북': '포항북',
}

VALID_REGION_KEYS = {
    '포항남', '포항북', '경주', '김천', '안동', '구미', '영주', '영천', '상주', '문경',
    '경산', '의성', '청송', '영양', '영덕', '청도', '고령', '성주', '칠곡', '예천',
    '봉화', '울진', '울릉'
}


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def get_safe_value(v):
    """값을 안전하게 스마트 변환 ('-' 표기는 0으로, 에러는 무시, 나머진 그대로)"""
    if v is None:
        return None
    if is_number(v):
        return v
    if isinstance(v, str):
        s = v.strip()
        if s in ['-', '']:
            return 0
        if s.startswith('#'):
            return None
    return v


def region_key(name):
    if not name or not isinstance(name, str):
        return None
    n = re.sub(r'\s+', '', name.strip())
    n = n.replace('광역시', '').replace('특별시', '')
    if n in PREFIX_SPECIAL:
        return PREFIX_SPECIAL[n]
    n2 = re.sub(r'(시|군|구)$', '', n)
    return n2 if n2 else None


def is_valid_region(label):
    return region_key(label) in VALID_REGION_KEYS


def extract_own_region_from_filename(filename):
    m = re.match(r'^\d{1,3}[_.\s]+([가-힣]+)', filename)
    if m:
        key = region_key(m.group(1))
        if key:
            return key
    for raw_name, mapped_key in PREFIX_SPECIAL.items():
        if raw_name in filename.replace(' ', ''):
            return mapped_key
    for k in VALID_REGION_KEYS:
        if k in filename:
            return k
    return None


def get_sheet_by_index(wb, idx):
    if idx < len(wb.sheetnames):
        return wb[wb.sheetnames[idx]]
    return None


def target_keys_for_region(own_key):
    """포항은 남구/북구 2개 구간 모두 대상으로 처리해야 함.
    set이 아닌 list를 반환해 순서를 고정한다 (set은 실행마다 순회 순서가
    달라질 수 있어 간헐적인 오류를 유발할 수 있음)."""
    if own_key == '포항':
        return ['포항남', '포항북']
    return [own_key] if own_key else []
