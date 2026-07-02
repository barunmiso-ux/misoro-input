"""OKTAS 환자 export(.xls) 파서 — 케이스 추적기 P1-1.

OKTAS '환자검색결과' export(구형 OLE2 .xls)를 읽어 구조화된 초진 환자 레코드로 변환.
- 열은 '헤더명'으로 접근 → 열 순서가 바뀌어도 안전.
- PII(전화·휴대폰·보험증·보험기호·주민·생일·주소)는 버린다(개인정보 최소화).
- 질환명 → 질환군 자동 분류(config.DISEASE_CATEGORY). 미분류 질환명은 '기타'로 떨구고 로그.
- 차트의 재활용 칸 해석: EMail=예약여부, 직업=상담자, 진행치료=한약/치료, URL=비예약원인.

사용:
  python export_parser.py "<export.xls 경로>"
  (코드)  from export_parser import parse_export;  r = parse_export(path)
"""
from __future__ import annotations

import sys
from dataclasses import dataclass, field

import pandas as pd


def _read_excel(path):
    """OKTAS export 읽기 — .xls(xlrd) / .xlsx(openpyxl) 모두 처리(확장자·내용 무관)."""
    try:
        return pd.read_excel(path, engine="xlrd", header=0, dtype=str)
    except Exception:
        return pd.read_excel(path, engine="openpyxl", header=0, dtype=str)

try:
    import config  # 같은 폴더: DISEASE_CATEGORY / DISEASE_GROUPS 재활용
    _DISEASE_CATEGORY = dict(config.DISEASE_CATEGORY)
except Exception:  # config 없을 때도 단독 동작
    _DISEASE_CATEGORY = {}

# ── 헤더명(영문키 ← OKTAS export 헤더) ───────────────────────────
# 보존: 분석/식별에 필요한 것만. (PII 헤더는 의도적으로 제외)
KEEP = {
    "chart_no": "차트번호", "name": "이름", "age": "나이", "sex": "성별",
    "last_visit": "최근내원", "registered": "등록일",
    "disease": "VIP", "doctor": "주치의",
    "inflow": "유입경로", "counselor": "직업",
    "treatment": "진행치료", "booking_raw": "EMail", "no_resv_reason": "URL",
}
# 버리는 PII 헤더(참고용 — 코드는 KEEP 만 읽으므로 자동 제외)
PII = ["전화번호", "휴대폰", "보험증번호", "보험기호", "주민등록(앞자리)", "생일", "주소", "우편번호", "피보험자", "예정치료"]

# 시트 환자테이블 C~Y 의 정규(canonical) 열순서 = OKTAS export 23열 순서.
# 시트 기록은 '위치 기반'(C+i)이므로 이 순서로 재배열해 써야 한다(검증됨 2026-06-30).
EXPORT_HEADERS = [
    "이름", "차트번호", "나이", "전화번호", "휴대폰", "보험증번호", "보험기호",
    "주민등록(앞자리)", "성별", "생일", "최근내원", "등록일", "우편번호", "주소",
    "EMail", "VIP", "주치의", "피보험자", "유입경로", "직업", "진행치료", "예정치료", "URL",
]
# 시트 기록 시 비울 민감 PII. 휴대폰은 유지(문의↔초진 노쇼 매칭 키 — 결정 2026-07-01).
# 주소·이름도 유지(지역구통계·식별 최소).
SHEET_PII_BLANK = ["전화번호", "보험증번호", "보험기호",
                   "주민등록(앞자리)", "생일", "우편번호"]

# config 4분류 → 시트 3그룹(피부/호흡기/통증기타) 매핑
GROUP_TO_SHEET = {"피부질환": "피부", "호흡기질환": "호흡기",
                  "기타": "통증기타", "첩약보험": "통증기타"}


@dataclass
class Patient:
    chart_no: str = ""
    name: str = ""          # 식별 최소(마스킹 표시용). 분석엔 차트번호 사용.
    age: str = ""
    sex: str = ""
    disease: str = ""       # 원본 질환명
    disease_group: str = "" # 시트 그룹(피부/호흡기/통증기타) — 미분류는 '기타'
    doctor: str = ""        # 주치의(원장)
    inflow: str = ""        # 유입경로
    counselor: str = ""     # 상담자
    booking: str = ""       # 예약/예약안함/미상
    herbal_paid: object = None  # True(한약결제)/False(치료만)/None(미정) — 시트호환용 보조
    outcome: str = ""           # 진료결과: 한약결제/약침결제/일반치료/상담만/미정/기타
    treatment_raw: str = ""     # 진행치료 원본
    no_resv_reason: str = ""    # 비예약원인
    registered: str = ""
    last_visit: str = ""
    missing: list = field(default_factory=list)  # 비어있는 핵심 주석 필드
    nonstd: str = ""  # 진행치료 비표준값(특화질환, 교정필요). 표준이면 ""


def _s(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


def _classify_disease(name: str) -> tuple[str, bool]:
    """질환명 → 시트그룹. 반환 (그룹, 분류성공여부)."""
    n = _s(name)
    if not n:
        return "", True
    grp = _DISEASE_CATEGORY.get(n)
    if grp:
        return GROUP_TO_SHEET.get(grp, "통증기타"), True
    return "기타", False  # 미분류


def _parse_booking(email_cell: str) -> str:
    v = _s(email_cell).replace("@", "")
    if "예약안함" in v:
        return "예약안함"
    if "예약" in v:
        return "예약"
    return "미상"


# 진료결과 분류 — 표준 어휘(2026-07-01 확정) + 지점 변형 흡수.
# 표준값: 한약N달 / 약침패키지 / 특화치료 / 일반치료 / 상담외용제 / 그냥감
# '결제'(전환)=한약결제+약침결제. 특화치료(1회·패키지X)·일반치료·상담만·그냥감=미전환.
# 규칙: 한약 있으면(한약만 or 한약+약침) 한약결제(시트 한약N달 수식 보존), 한약없이 약침만=약침결제.
# 결제(전환) = 한약N달 + 약침패키지 + 첩약보험(급여 첩약). 첩약보험은 한약N달과 구분 집계.
PAID_OUTCOMES = ("한약결제", "약침결제", "첩약보험")

# 표준 어휘(정확일치) — 이 값 아니면 '비표준=교정필요'로 표시(흡수는 하되 교정 유도)
STANDARD_TREATMENTS = {"한약1달", "한약3달", "한약6달", "한약12달", "약침패키지",
                       "첩약보험", "특화치료", "일반치료", "상담만", "그냥감"}
STANDARD_RESULTS = {"예약완료", "예약안함", "재통화필요"}
# 진료결과 → 지점이 쳐야 할 표준 입력값(교정 제안용)
OUTCOME_TO_STD = {"한약결제": "한약N달(기간명시)", "약침결제": "약침패키지",
                  "첩약보험": "첩약보험", "특화치료": "특화치료", "일반치료": "일반치료",
                  "상담만": "상담만", "그냥감": "그냥감", "기타": "표준값 중 하나"}


def classify_treatment(treatment_cell: str) -> str:
    """진행치료 원본 → 진료결과 라벨. 표준어휘 + 변형(치료만/그냥감 등) 흡수."""
    v = _s(treatment_cell)
    if not v:
        return "미정"
    if "첩약" in v:                       # 첩약보험(급여 첩약) → 결제(한약N달과 구분)
        return "첩약보험"
    if "한약" in v or "힌약" in v or "공진단" in v:  # 한약N달·공진단(비급여) → 결제 (힌약=오타)
        return "한약결제"
    if "약침" in v or "패키지" in v:     # 약침패키지(한약없이) → 결제
        return "약침결제"
    if "특화" in v:                       # 특화치료 1회(패키지X) — 치료받음, 미전환
        return "특화치료"
    if "상담" in v or "외용" in v:        # 상담외용제·상담만 = 치료 안 받고 감
        return "상담만"
    if "그냥" in v:                       # 그냥감 = 진료 안 봄(완전 이탈)
        return "그냥감"
    if "치료" in v:                       # 일반치료·치료만 = 급여치료(미결제)
        return "일반치료"
    return "기타"                         # 미지값(통증 등 오기입) → 로그


def to_standard_treatment(raw: str):
    """진행치료 원본 → (표준값, 자동변환됨?).

    - 이미 표준값 → (그값, False)
    - 확신되는 변형(일빈치료→일반치료·상담+외용제→상담만·힌약3달→한약3달) → (표준값, True)
    - 애매/엉뚱(통증·기간없는 한약·알수없는값) → **(None, False)** ← 기록 강제차단 대상
    - 빈값 → ("", False)
    """
    import re as _re
    v = _s(raw)
    if not v:
        return "", False
    if v in STANDARD_TREATMENTS:
        return v, False
    cat = classify_treatment(v)
    simple = {"일반치료": "일반치료", "상담만": "상담만", "그냥감": "그냥감",
              "약침결제": "약침패키지", "첩약보험": "첩약보험", "특화치료": "특화치료"}
    if cat in simple:
        return simple[cat], True
    if cat == "한약결제":
        m = _re.search(r"(\d+)\s*달", v)   # 기간 숫자 있으면 살림(힌약3달→한약3달)
        if m:
            return f"한약{m.group(1)}달", True
        return None, False                 # 기간 미상(한약·공진단 등) → 강제 교정
    return None, False                     # 기타/미정(엉뚱한 값) → 강제 교정


def normalize_result(raw: str) -> str:
    """상담결과 원본 → 표준값(예약완료/예약안함/재통화필요). 못 맞추면 원본 유지.

    지점 표기 대혼란 흡수: '예약'(천안195·인천128)·'예약@'→예약완료 /
    '예약취소'·'예약후취소'·'예약안함@'·'예약암함'·'예역언험'→예약안함 /
    '부재중'·'재통화'→재통화필요. ('예약전'은 모호 → 원본 유지=미전환 취급.)
    순서 주의: 취소/안함을 '예약' 일반매칭보다 먼저.
    """
    v = _s(raw)
    if not v:
        return ""
    low = v.replace(" ", "")
    if any(k in low for k in ("취소", "안함", "암함", "언험")):
        return "예약안함"
    if any(k in low for k in ("재통화", "부재중", "부재", "재콜")):
        return "재통화필요"
    if low == "예약전":                   # 예약 전(前)/모호 → 표준화 안 함
        return v
    if "예약" in low:                     # 예약완료·예약·예약@
        return "예약완료"
    return v


def _parse_herbal(treatment_cell: str):
    """진행치료 → (herbal_paid, 원본). 시트호환 보조: 한약=True, 그 외=False, 공란=None."""
    v = _s(treatment_cell)
    if not v:
        return None, ""
    return ("한약" in v), v


def parse_export(path: str) -> dict:
    """export 파일 → {patients, summary, completeness, unmapped_diseases}."""
    df = _read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    cols = set(df.columns)
    missing_headers = [h for h in KEEP.values() if h not in cols]

    patients: list[Patient] = []
    unmapped: dict[str, int] = {}
    for _, row in df.iterrows():
        get = lambda key: _s(row.get(KEEP[key])) if KEEP[key] in cols else ""
        chart = get("chart_no")
        if not chart and not get("name"):
            continue  # 빈 행
        grp, ok = _classify_disease(get("disease"))
        if not ok:
            d = get("disease")
            unmapped[d] = unmapped.get(d, 0) + 1
        herbal_paid, traw = _parse_herbal(get("treatment"))
        p = Patient(
            chart_no=chart, name=get("name"), age=get("age"), sex=get("sex"),
            disease=get("disease"), disease_group=grp, doctor=get("doctor"),
            inflow=get("inflow"), counselor=get("counselor"),
            booking=_parse_booking(get("booking_raw")),
            herbal_paid=herbal_paid, outcome=classify_treatment(get("treatment")),
            treatment_raw=traw, no_resv_reason=get("no_resv_reason"),
            registered=get("registered"), last_visit=get("last_visit"),
        )
        # 완전성(강제입력): 특화질환(피부·호흡기)은 진행치료·예약·상담자·유입 필수.
        # 통증/기타는 전환분석 대상 아니라 선택(강제 안 함).
        if p.disease_group in ("피부", "호흡기"):
            for label, val in (("유입경로", p.inflow), ("진행치료", p.treatment_raw),
                               ("예약여부", p.booking if p.booking != "미상" else ""),
                               ("상담자", p.counselor)):
                if not val:
                    p.missing.append(label)
            # 비표준 진행치료값(흡수는 되지만 교정 유도)
            if p.treatment_raw and p.treatment_raw not in STANDARD_TREATMENTS:
                p.nonstd = p.treatment_raw
        patients.append(p)

    # 집계
    from collections import Counter
    grp_counts: dict[str, int] = {}
    for p in patients:
        if p.disease_group:
            grp_counts[p.disease_group] = grp_counts.get(p.disease_group, 0) + 1
    booked = sum(1 for p in patients if p.booking == "예약")
    outcome_counts = Counter(p.outcome for p in patients)
    paid = sum(outcome_counts.get(k, 0) for k in PAID_OUTCOMES)  # 결제함=한약+약침
    # 특화(피부·호흡기) = 전환 측정 대상. 통증/기타는 제외.
    # 그냥감(진료 안 봄)은 결제 판단 이전 이탈 → 전환율 분모에서 제외(진료 본 사람 기준).
    teukhwa = [p for p in patients if p.disease_group in ("피부", "호흡기")]
    teukhwa_seen = [p for p in teukhwa if p.outcome != "그냥감"]   # 진료 본 특화초진
    teukhwa_geunjang = len(teukhwa) - len(teukhwa_seen)           # 진료 전 이탈(그냥감)
    teukhwa_paid = sum(1 for p in teukhwa_seen if p.outcome in PAID_OUTCOMES)
    incomplete = [p for p in patients if p.missing]

    return {
        "patients": patients,
        "summary": {
            "초진수": len(patients),
            "질환군별": grp_counts,
            "예약": booked, "예약안함": sum(1 for p in patients if p.booking == "예약안함"),
            # 진료결과: 한약결제/약침결제/특화치료/일반치료/상담만/그냥감/미정/기타
            "진료결과별": dict(outcome_counts),
            # '결제'(전환) = 한약결제 + 약침결제. 특화치료(1회)·일반치료·상담만·그냥감=미결제.
            "결제함": paid,
            "결제안함": len(patients) - paid,
            "한약결제": outcome_counts.get("한약결제", 0),
            "약침결제": outcome_counts.get("약침결제", 0),
            "첩약보험": outcome_counts.get("첩약보험", 0),
            "특화치료": outcome_counts.get("특화치료", 0),
            "일반치료": outcome_counts.get("일반치료", 0),
            "상담만": outcome_counts.get("상담만", 0),
            "그냥감": outcome_counts.get("그냥감", 0),
            "미정": outcome_counts.get("미정", 0),
            # 특화 결제전환율 = (한약+약침) / (특화초진 − 그냥감) — 진료 본 사람 기준(정직)
            "특화초진": len(teukhwa),
            "특화진료전이탈": teukhwa_geunjang,   # 그냥감(진료 안 봄)
            "특화결제": teukhwa_paid,
            "특화전환율": (teukhwa_paid / len(teukhwa_seen)) if teukhwa_seen else None,
        },
        "completeness": {
            "미완성_환자수": len(incomplete),
            "미완성목록": [(p.chart_no, p.name[:1] + "*", p.disease, p.missing) for p in incomplete],
            "비표준목록": [(p.chart_no, p.name[:1] + "*", p.nonstd,
                          OUTCOME_TO_STD.get(classify_treatment(p.nonstd), "표준값"))
                         for p in patients if p.nonstd],
        },
        "unmapped_diseases": unmapped,
        "missing_headers": missing_headers,
    }


def _date_ymd(s: str):
    """'YYYY-MM-DD...' → (y,m,d) 또는 None."""
    t = _s(s)[:10]
    try:
        y, m, d = (int(x) for x in t.split("-"))
        return y, m, d
    except Exception:
        return None


def week_of_month(day: int) -> int:
    """시트 주차 규칙(검증됨): ((일-1)//7)+1. 1~7일=1주 ... 22~28일=4주."""
    return ((day - 1) // 7) + 1


def week_range(tab: str):
    """'YY-MM-N주' → (시작date, 종료date). 주차규칙 역산(N주 = (N-1)*7+1일 ~ N*7일).
    월 마지막일로 클램프. 주차탭 아니면(월탭 등) None."""
    import re as _re
    import calendar
    import datetime
    m = _re.match(r"(\d{2})-(\d{2})-(\d)주", (tab or "").strip())
    if not m:
        return None
    yy, mm, n = int(m.group(1)), int(m.group(2)), int(m.group(3))
    year = 2000 + yy
    last = calendar.monthrange(year, mm)[1]
    start = min((n - 1) * 7 + 1, last)
    end = min(n * 7, last)
    return datetime.date(year, mm, start), datetime.date(year, mm, end)


def week_label(tab: str) -> str:
    """'26-06-4주' → '26-06-4주 (6/22~6/28)'. 범위 못구하면 원본 반환."""
    r = week_range(tab)
    if not r:
        return tab
    s, e = r
    return f"{tab} ({s.month}/{s.day}~{e.month}/{e.day})"


def _tab_from_weeks(weeks: list) -> tuple:
    from collections import Counter
    if not weeks:
        return None, {"reason": "날짜 없음", "multi": False}
    cnt = Counter(weeks)
    (y, m, w), n = cnt.most_common(1)[0]
    tab = f"{y % 100:02d}-{m:02d}-{w}주"
    return tab, {
        "counts": {f"{yy % 100:02d}-{mm:02d}-{ww}주": c for (yy, mm, ww), c in cnt.items()},
        "dominant": tab, "dominant_n": n, "total": len(weeks), "multi": len(cnt) > 1,
    }


def infer_week_from_dates(date_strs: list) -> tuple:
    """날짜문자열 목록 → 주차탭명 추정 (tab|None, info). 형식 'YY-MM-N주'."""
    weeks = []
    for s in date_strs:
        ymd = _date_ymd(s)
        if ymd:
            y, m, d = ymd
            weeks.append((y, m, week_of_month(d)))
    return _tab_from_weeks(weeks)


def infer_week_tab(patients: list) -> tuple:
    """환자 초진일(등록일, 없으면 최근내원)로 주차탭명 추정."""
    return infer_week_from_dates(
        [getattr(p, "registered", "") or getattr(p, "last_visit", "") for p in patients])


def rows_for_sheet(path: str, mask_pii: bool = True) -> list:
    """시트 환자테이블(C~Y)에 위치기록할 23열 행 목록.

    - 헤더명으로 정규 순서(EXPORT_HEADERS)로 재배열 → 열순서가 바뀌어도 안전.
    - mask_pii=True면 민감 PII(전화·휴대폰·보험·주민·생일·우편) 공란(주소·이름 유지).
    - 차트·이름 둘 다 빈 행은 제외.
    반환: [[c1..c23], ...]  (export 순서 = 시트 C..Y)
    """
    df = _read_excel(path).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    cols = set(df.columns)
    blank = set(SHEET_PII_BLANK) if mask_pii else set()

    out = []
    for _, row in df.iterrows():
        if not _s(row.get("차트번호")) and not _s(row.get("이름")):
            continue
        rec = []
        for h in EXPORT_HEADERS:
            rec.append("" if (h in blank or h not in cols) else _s(row.get(h)))
        out.append(rec)
    return out


def _week_tab_of(date_str: str):
    """등록일/상담시각 문자열 → 'YY-MM-N주' 또는 None(날짜 못읽음)."""
    ymd = _date_ymd(date_str)
    if not ymd:
        return None
    y, m, d = ymd
    return f"{y % 100:02d}-{m:02d}-{week_of_month(d)}주"


def rows_for_sheet_by_week(path: str, mask_pii: bool = True) -> dict:
    """초진 행을 **등록일 기준 주차탭별로 그룹핑**. 반환 {week_tab: [rows23], ...}.

    한 파일이 여러 주(또는 여러 달)에 걸쳐도 환자별 등록일로 각 주차에 자동 배분한다.
    날짜를 못 읽는 행은 '_미분류' 키에 모음(호출측에서 경고).
    """
    from collections import defaultdict
    df = _read_excel(path).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    cols = set(df.columns)
    blank = set(SHEET_PII_BLANK) if mask_pii else set()
    groups = defaultdict(list)
    for _, row in df.iterrows():
        if not _s(row.get("차트번호")) and not _s(row.get("이름")):
            continue
        wk = _week_tab_of(_s(row.get("등록일"))) or "_미분류"
        rec = []
        for h in EXPORT_HEADERS:
            val = "" if (h in blank or h not in cols) else _s(row.get(h))
            if h == "진행치료" and val:
                std, _conv = to_standard_treatment(val)   # 확신되는 변형은 표준값으로 기록
                if std is not None:                        # None(엉뚱)이면 원본 유지(앱이 기록 차단)
                    val = std
            rec.append(val)
        groups[wk].append(rec)
    return dict(groups)


# ══════════════════════════════════════════════════════════════════
# 문의(상담내역) — 시트 ②상담테이블(Z~AK)
# ══════════════════════════════════════════════════════════════════
# OKTAS 상담내역 export 13열 → 시트 AA~AK 정규순서 11열(번호 Z는 별도). 주의·상담메모 제외.
INQUIRY_SHEET_ORDER = [
    "상담시각", "차트번호", "성명", "전화번호", "상담구분", "유입경로",
    "진료구분", "상담결과", "상담자", "상담완료", "다음콜시각",
]


@dataclass
class Inquiry:
    consult_time: str = ""
    chart_no: str = ""
    name: str = ""
    channel: str = ""        # 상담구분(전화문의/홈피문의/카톡문의/네이버예약/재진문의…)
    inflow: str = ""         # 유입경로
    disease: str = ""        # 진료구분(질환)
    disease_group: str = ""  # 피부/호흡기/통증기타
    result: str = ""         # 상담결과(예약완료/예약안함/재통화필요…)
    counselor: str = ""
    booked: object = None    # True(예약완료)/False(예약안함)/None(진행중)


def _read_inquiry_df(path: str):
    df = _read_excel(path).fillna("")
    df.columns = [str(c).replace("\n", "").strip() for c in df.columns]  # '상담\n완료'→'상담완료'
    return df


def _inquiry_is_empty(row) -> bool:
    return not (_s(row.get("차트번호")) or _s(row.get("성명")) or _s(row.get("전화번호")))


def parse_inquiries(path: str) -> dict:
    """상담내역 export → {inquiries, summary, unmapped_diseases}."""
    from collections import Counter
    df = _read_inquiry_df(path)
    cols = set(df.columns)
    g = lambda row, h: _s(row.get(h)) if h in cols else ""

    inquiries, unmapped = [], {}
    for _, row in df.iterrows():
        if _inquiry_is_empty(row):
            continue
        dis = g(row, "진료구분")
        grp, ok = _classify_disease(dis)
        if not ok and dis:
            unmapped[dis] = unmapped.get(dis, 0) + 1
        res = g(row, "상담결과")
        # 흡수: 안함/취소=False, 그 외 '예약'포함(예약완료·예약@·예약)=True, 나머지(재통화 등)=None
        if "안함" in res or "취소" in res:
            booked = False
        elif "예약" in res:
            booked = True
        else:
            booked = None
        inquiries.append(Inquiry(
            consult_time=g(row, "상담시각"), chart_no=g(row, "차트번호"), name=g(row, "성명"),
            channel=g(row, "상담구분"), inflow=g(row, "유입경로"),
            disease=dis, disease_group=grp, result=res,
            counselor=g(row, "상담자"), booked=booked,
        ))

    grp_counts = {}
    for i in inquiries:
        if i.disease_group:
            grp_counts[i.disease_group] = grp_counts.get(i.disease_group, 0) + 1
    booked = sum(1 for i in inquiries if i.booked is True)
    n = len(inquiries)
    return {
        "inquiries": inquiries,
        "summary": {
            "문의수": n,
            "예약완료": booked,
            "예약안함": sum(1 for i in inquiries if i.booked is False),
            "진행중": sum(1 for i in inquiries if i.booked is None),
            "예약전환율": (booked / n) if n else 0.0,
            "질환군별": grp_counts,
            "상담구분별": dict(Counter(i.channel for i in inquiries if i.channel)),
            "상담결과별": dict(Counter(i.result for i in inquiries if i.result)),
            # 비표준 상담결과값(교정필요): 예약@·예약·예약취소 등
            "비표준결과": [((i.name[:1] + "*") if i.name else "무명", i.result)
                          for i in inquiries if i.result and i.result not in STANDARD_RESULTS],
        },
        "unmapped_diseases": unmapped,
    }


def inquiry_rows_for_sheet(path: str) -> list:
    """시트 ②상담테이블(AA~AK)에 위치기록할 11열 행. (번호 Z는 writer가 붙임)
    헤더명으로 정규순서 재배열 → 열순서 안전. 주의·상담메모는 시트에 없어 제외.
    """
    df = _read_inquiry_df(path)
    cols = set(df.columns)
    out = []
    for _, row in df.iterrows():
        if _inquiry_is_empty(row):
            continue
        out.append([_s(row.get(h)) if h in cols else "" for h in INQUIRY_SHEET_ORDER])
    return out


def inquiry_week_tab(path: str) -> tuple:
    """상담내역 export의 상담시각으로 주차탭 추정."""
    df = _read_inquiry_df(path)
    times = [_s(v) for v in df.get("상담시각", [])] if "상담시각" in df.columns else []
    return infer_week_from_dates(times)


def inquiry_rows_by_week(path: str) -> dict:
    """문의 행을 **상담시각 기준 주차탭별로 그룹핑**. 반환 {week_tab: [rows11], ...}."""
    from collections import defaultdict
    df = _read_inquiry_df(path)
    cols = set(df.columns)
    groups = defaultdict(list)
    for _, row in df.iterrows():
        if _inquiry_is_empty(row):
            continue
        wk = _week_tab_of(_s(row.get("상담시각"))) or "_미분류"
        rec = []
        for h in INQUIRY_SHEET_ORDER:
            val = _s(row.get(h)) if h in cols else ""
            if h == "상담결과" and val:
                norm = normalize_result(val)          # 예약→예약완료 등 표준 정규화
                if norm in STANDARD_RESULTS:
                    val = norm                         # 아니면 원본(앱이 차단)
            rec.append(val)
        groups[wk].append(rec)
    return dict(groups)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    path = sys.argv[1] if len(sys.argv) > 1 else None
    if not path:
        print("사용: python export_parser.py \"<export.xls 경로>\"")
        sys.exit()
    r = parse_export(path)
    s = r["summary"]
    print(f"=== 파싱 결과 ===")
    if r["missing_headers"]:
        print("⚠ 못 찾은 헤더:", r["missing_headers"])
    print(f"초진수 {s['초진수']} / 예약 {s['예약']} 예약안함 {s['예약안함']}")
    print(f"한약결제 {s['한약결제']} 치료만 {s['치료만']} 한약미정 {s['한약미정']}")
    print(f"질환군별: {s['질환군별']}")
    print(f"\n환자 레코드(PII 마스킹):")
    for p in r["patients"]:
        print(f"  [{p.chart_no}] {p.name[:1]}* / {p.disease}({p.disease_group}) / "
              f"유입:{p.inflow} / 예약:{p.booking} / 한약:{p.herbal_paid} / "
              f"상담:{p.counselor} / 원장:{p.doctor}"
              + (f"  ⚠미입력:{p.missing}" if p.missing else ""))
    if r["unmapped_diseases"]:
        print(f"\n⚠ 미분류 질환명(분류표에 추가 필요): {r['unmapped_diseases']}")
    c = r["completeness"]
    print(f"\n완전성: 미완성 {c['미완성_환자수']}명")
    for chart, nm, miss in c["미완성목록"]:
        print(f"  [{chart}] {nm}: {miss} 비어있음")
