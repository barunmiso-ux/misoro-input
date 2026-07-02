"""미소로한의원 일일 보고 입력 앱 — 설정값.

년도·월은 절대 하드코딩하지 않는다 (GAS 의 YEAR:2026 하드코딩 사고 방지).
시트 이름은 제출 날짜 기준으로 동적으로 생성한다.
"""

# 구글 시트 ID (misoro-dashboard / GAS 와 동일 시트 공유)
SHEET_ID = "1uTkikVDCUfVry6l-GX2L8yvlhRSVFdUtEwlubrhpZLk"

# 서비스 계정 (이 이메일을 일일 보고 시트에 '편집자'로 추가해야 쓰기 가능)
SERVICE_ACCOUNT_EMAIL = "rank-check-483711@rank-check-483711.iam.gserviceaccount.com"

# 9개 지점 (부산 2026-07 탈퇴)
ALL_BRANCHES = [
    "대전", "분당", "수원", "안산",
    "영등포", "인천", "전주", "천안", "평택",
]

# PoC 단계: 분당만 실제 운영. (검증 후 ALL_BRANCHES 로 확장)
ACTIVE_BRANCHES = ["분당"]

# ── 초진·문의 섹션 정의 ─────────────────────────────────────────
# key      : 폼 내부 식별자 / 세션 상태 키
# label    : 화면에 보이는 섹션 제목
# allow_route: 유입경로 드롭다운(초진만) / sheet_col: 일일시트 기록 컬럼
# "기타" 채널은 일일시트에 대응 컬럼이 없어 합계는 기록 안 되고 초진상세 로그에만 남는다.
INQUIRY_SECTIONS = [
    {"key": "초진",   "label": "초진",   "sheet_col": "내원",   "allow_route": True},
    {"key": "네이버", "label": "네이버", "sheet_col": "네이버", "allow_route": False},
    {"key": "카카오", "label": "카카오", "sheet_col": "카카오", "allow_route": False},
    {"key": "전화",   "label": "전화",   "sheet_col": "전화",   "allow_route": False},
    {"key": "홈페이지", "label": "홈페이지", "sheet_col": "홈페이지", "allow_route": False},
    {"key": "워크인", "label": "워크인", "sheet_col": "워크인", "allow_route": False},
    {"key": "기타",   "label": "기타",   "sheet_col": "기타",   "allow_route": False},
]

# ── 병명 드롭다운 (초진통계 _TEMPLATE 기준) ─────────────────────
DISEASE_GROUPS = {
    "호흡기질환": ["비염", "축농증", "중이염", "기타호흡기"],
    "피부질환": ["아토피", "사마귀", "건선", "습진", "한포진",
               "지루성피부염", "지루성두피염", "모낭염", "백반증",
               "두드러기", "다한증", "기타피부"],
    "기타": ["통증", "교통사고", "성장질환", "산후질환", "공진단", "보약", "기타"],
    "첩약보험": ["생리통", "안면신경마비", "뇌혈관질환후유증",
               "요추추간판탈출증", "기능성소화불량", "알레르기비염"],
}

# ── 유입경로 드롭다운 (초진 한정, 초진통계 _TEMPLATE 기준) ───────
ROUTE_GROUPS = {
    "온라인": ["홈페이지", "블로그", "카페", "뉴스기사", "지식인", "지도", "배너",
             "유튜브", "SNS", "페이스북", "인스타", "밴드", "카톡", "모름"],
    "오프라인": ["방송(TV)", "라디오", "소개", "강의", "간판", "협약기관", "버스광고",
              "지하철광고", "택시광고", "신문광고", "현수막", "기타"],
}

PLACEHOLDER = "(선택)"


def _flatten(groups):
    """그룹 dict → 셀렉트박스용 평탄 리스트. 그룹 구분선(── 그룹 ──) 포함."""
    opts = [PLACEHOLDER]
    for g, names in groups.items():
        opts.append(f"── {g} ──")
        opts.extend(names)
    return opts


DISEASE_OPTIONS = _flatten(DISEASE_GROUPS)   # 모든 채널 공통: 세부 병명
ROUTE_OPTIONS = _flatten(ROUTE_GROUPS)       # 초진 한정: 유입경로

# 그룹 구분선(선택 불가 항목) 판별용
GROUP_DIVIDERS = {o for o in DISEASE_OPTIONS + ROUTE_OPTIONS if o.startswith("── ")}

# 병명 → 대분류 매핑 (저장 시 분류 기록용)
DISEASE_CATEGORY = {n: g for g, names in DISEASE_GROUPS.items() for n in names}
ROUTE_CATEGORY = {n: g for g, names in ROUTE_GROUPS.items() for n in names}

# ── 카페숙제체크 탭 ──────────────────────────────────────────────
HOMEWORK_TAB = "카페숙제체크"
HOMEWORK_HEADERS = [
    "날짜", "지점", "건강_카페", "건강_홈페", "일상게시글",
    "댓글수", "댓글URL", "입력자", "기록시각",
]

# ── 초진상세 탭 (병명·경로 상세 로그, append 방식) ──────────────
DETAIL_TAB = "초진상세"
DETAIL_HEADERS = [
    "날짜", "지점", "채널", "대분류", "병명",
    "경로대분류", "경로", "인원", "입력자", "기록시각",
]

# ── 댓글 이미지 업로드 (공유 드라이브) ──────────────────────────
# 공유 드라이브 폴더 ID (비우면 이미지 업로드 비활성). 클라우드는 st.secrets["drive_folder_id"] 가능.
DRIVE_FOLDER_ID = "1sW5gy82lcV9R9Lh44E3ld2jyFtqhv4F5"
IMAGE_TYPES = ["png", "jpg", "jpeg", "gif", "webp"]

# 과거 입력 허용 범위 (오늘 포함 최근 N일)
PAST_DAYS_ALLOWED = 7

# 타임존 (KST)
TIMEZONE = "Asia/Seoul"


def daily_sheet_name(d):
    """제출 날짜 d(datetime.date) 기준 월별 일일 시트 이름.

    예: 2026-06-16 → '[26.06] 지점별 주간 문의내역'
    """
    return f"[{d:%y.%m}] 지점별 주간 문의내역"
