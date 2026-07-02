"""문의→초진 노쇼 매칭 — 케이스 추적기 P2.

원리(2026-06-30 설계): 내원은 문의 다음에 발생. 그래서 문의 예약완료를 기준점(상담시각)으로
'문의일 + 윈도우(기본 2주)' 안에 초진 명단에 나타났는지 대조 → 3상태 판정.
- 전환: 윈도우 안에 초진 매칭됨
- 내원대기: 아직 매칭 안 됐고 오늘 <= 마감일 (= 데스크 리마인드 대상)
- 노쇼: 매칭 안 됐고 오늘 > 마감일 AND 마감 걸친 초진 주가 시트에 존재(데이터 있음)
- 데이터대기: 마감 지났지만 걸친 초진 주 탭이 아직 없음(업로드 지연) → 노쇼로 단정 안 함

매칭 키: 차트번호 → 전화(휴대폰) → 성명 순. (무명+차트공란 문의는 전화로만 가능.)
시트는 읽기만 함(쓰기·구조변경 없음).
"""
from __future__ import annotations

from datetime import datetime, timedelta

from case_sheet_writer import _svc, DEFAULT_KEY, _tabs
from export_parser import (classify_treatment, _classify_disease, PAID_OUTCOMES,
                           normalize_result)

DEFAULT_WINDOW_DAYS = 14
OVERRIDE_TAB = "_노쇼보정"       # 별도 탭(주간시트 구조 미변경). 수기보정 저장.
OVERRIDE_STATES = ("전환", "노쇼", "내원대기")


def iq_key(iq: dict) -> str:
    """문의 고유키: 주차|상담시각|식별(성명>전화>차트). 보정 저장/조회용."""
    who = iq.get("name") or iq.get("phone") or iq.get("chart") or "?"
    return f"{iq.get('week')}|{iq.get('time')}|{who}"


def load_overrides(sid: str, key_path: str = DEFAULT_KEY) -> dict:
    """_노쇼보정 탭 → {키: 보정상태}. 탭 없으면 빈 dict."""
    sh = _svc(key_path)
    try:
        vals = sh.values().get(spreadsheetId=sid, range=f"'{OVERRIDE_TAB}'!A2:B").execute().get("values", [])
    except Exception:
        return {}
    return {r[0]: r[1] for r in vals if len(r) >= 2 and r[0].strip() and r[1].strip()}


def set_override(sid: str, key: str, status: str, writer: str = "", *,
                 key_path: str = DEFAULT_KEY, now: str = "") -> dict:
    """_노쇼보정 탭에 (키→상태) upsert. status 가 '' 또는 '(자동)' 이면 해당 키 제거."""
    sh = _svc(key_path)
    if OVERRIDE_TAB not in _tabs(sh, sid):
        sh.batchUpdate(spreadsheetId=sid,
                       body={"requests": [{"addSheet": {"properties": {"title": OVERRIDE_TAB}}}]}).execute()
        sh.values().update(spreadsheetId=sid, range=f"'{OVERRIDE_TAB}'!A1",
                           valueInputOption="RAW",
                           body={"values": [["키", "보정상태", "보정자", "시각"]]}).execute()
    vals = sh.values().get(spreadsheetId=sid, range=f"'{OVERRIDE_TAB}'!A2:D").execute().get("values", [])
    rows = [r for r in vals if r and r[0].strip()]
    clear = status in ("", "(자동)")
    kept = [r for r in rows if r[0] != key]
    if not clear:
        kept.append([key, status, writer, now])
    sh.values().clear(spreadsheetId=sid, range=f"'{OVERRIDE_TAB}'!A2:D").execute()
    if kept:
        sh.values().update(spreadsheetId=sid, range=f"'{OVERRIDE_TAB}'!A2",
                           valueInputOption="RAW", body={"values": kept}).execute()
    return {"key": key, "status": status or "(자동)", "removed": clear}

# 시트 열(0-based 그리드 기준 — 읽기 범위에 맞춰 인덱싱)
# 초진 C5:Y154 → 0:이름(C) 1:차트(D) 3:전화(F) 4:휴대폰(G) 11:등록일(N) 15:질환(R) 16:주치의(S) 20:진행치료(W)
CH_NAME, CH_CHART, CH_PHONE, CH_MOBILE, CH_REG = 0, 1, 3, 4, 11
CH_DISEASE, CH_DOCTOR, CH_TREAT = 15, 16, 20
# 문의 AA5:AI139 → 0:상담시각(AA) 1:차트(AB) 2:성명(AC) 3:전화(AD) 4:상담구분(AE) 6:진료구분(AG) 7:상담결과(AH) 8:상담자(AI)
IQ_TIME, IQ_CHART, IQ_NAME, IQ_PHONE, IQ_DISEASE, IQ_RESULT = 0, 1, 2, 3, 6, 7
IQ_CHANNEL, IQ_COUNSELOR = 4, 8


def _pdate(s):
    s = str(s or "").strip()
    if not s:
        return None
    s = s.split()[0]
    for fmt in ("%Y-%m-%d", "%m/%d/%y", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _digits(s):
    return "".join(ch for ch in str(s or "") if ch.isdigit())


def _name(s):
    s = str(s or "").strip()
    return "" if s in ("", "무명") else s


def _cell(row, i):
    return str(row[i]).strip() if i < len(row) else ""


def _load(sid: str, tabs: list, key_path: str = DEFAULT_KEY, default_counselor: str = ""):
    """모든 주 탭에서 초진/문의 인덱스 로드. 반환 (chojin[list], inquiries[list], tabset).

    좁은/비표준 탭(그리드 26열 등)이 섞여도 안 죽게 그리드폭 확인 후 문의범위(AA:AI, 35열)를
    조건부로만 요청 — 안 그러면 좁은 탭 하나가 배치 전체를 실패시킴.
    """
    sh = _svc(key_path)
    meta = sh.get(spreadsheetId=sid,
                  fields="sheets(properties(title,gridProperties(columnCount)))").execute()
    width = {s["properties"]["title"]: s["properties"].get("gridProperties", {}).get("columnCount", 0)
             for s in meta.get("sheets", [])}
    ranges, has_iq = [], []
    for t in tabs:
        ranges.append(f"'{t}'!C5:Y154")
        iq_ok = width.get(t, 0) >= 35
        has_iq.append(iq_ok)
        if iq_ok:
            ranges.append(f"'{t}'!AA5:AI139")
    vrs = sh.values().batchGet(spreadsheetId=sid, ranges=ranges).execute().get("valueRanges", [])

    chojin, inquiries = [], []
    vi = 0
    for t, iq_ok in zip(tabs, has_iq):
        for row in vrs[vi].get("values", []):
            nm = _cell(row, CH_NAME)
            ch = _cell(row, CH_CHART)
            reg = _pdate(_cell(row, CH_REG))
            # 실제 초진은 숫자 차트번호 + 등록일(날짜) 둘 다 보유. 통계영역 행(비용·거리·
            # 비예약원인 라벨/숫자, 짧은 구버전 탭에서 C5:Y154가 통계구간까지 긁음)은 배제.
            if not _digits(ch) or reg is None:
                continue
            chojin.append({
                "week": t, "name": _name(nm), "chart": ch,
                "phones": {_digits(_cell(row, CH_PHONE)), _digits(_cell(row, CH_MOBILE))} - {""},
                "reg": reg,
                "doctor": _cell(row, CH_DOCTOR),
                "group": _classify_disease(_cell(row, CH_DISEASE))[0],
                "outcome": classify_treatment(_cell(row, CH_TREAT)),
            })
        vi += 1
        if not iq_ok:
            continue
        for row in vrs[vi].get("values", []):
            res = _cell(row, IQ_RESULT)
            nm = _cell(row, IQ_NAME)
            ph = _digits(_cell(row, IQ_PHONE))
            if not res and not nm and not ph:
                continue
            inquiries.append({
                "week": t, "time": _pdate(_cell(row, IQ_TIME)),
                "name": _name(nm), "chart": _cell(row, IQ_CHART),
                "phone": ph, "disease": _cell(row, IQ_DISEASE),
                "result": normalize_result(res),
                "channel": _cell(row, IQ_CHANNEL),
                "counselor": _cell(row, IQ_COUNSELOR) or default_counselor,
            })
        vi += 1
    return chojin, inquiries, set(tabs)


def aggregate_by_period(result: dict) -> dict:
    """매칭결과 rows → 기간별(주차+월) 집계. period → {문의수,전환,노쇼,내원대기,전환율…}.

    각 문의를 자기 주차('26-06-3주')와 그 달('26-06월') 양쪽에 집계.
    전환율 = 전환/(전환+노쇼) (내원대기·데이터대기 제외 = 확정분 기준).
    """
    from collections import defaultdict

    def _blank():
        return {"문의수": 0, "전환": 0, "노쇼": 0, "내원대기": 0, "데이터대기": 0, "판정불가": 0}

    agg = defaultdict(_blank)
    for r in result["rows"]:
        wk = r["week"]                          # '26-06-3주'
        mo = wk.rsplit("-", 1)[0] + "월"        # '26-06월'
        for period in (wk, mo):
            a = agg[period]
            a["문의수"] += 1
            a[r["status"]] = a.get(r["status"], 0) + 1
    out = {}
    for p, a in agg.items():
        denom = a["전환"] + a["노쇼"]
        a["전환율"] = round(a["전환"] / denom, 4) if denom else None
        out[p] = a
    return out


def _periods_of(week_tab: str) -> tuple:
    """'26-06-3주' → ('26-06-3주', '26-06월'). 월탭이면 그 자신만. (주/월 양쪽 집계용.)"""
    if week_tab.endswith("월"):
        return (week_tab,)
    return (week_tab, week_tab.rsplit("-", 1)[0] + "월")


def aggregate_doctors(result: dict) -> dict:
    """기간별 × 주치의(원장) 특화 결제전환율. 반환 {period: {doctor: {...}}}.

    결제전환율 = 결제(한약+약침+첩약보험) / (특화초진 − 그냥감).
    특화(피부·호흡기)만 대상 — 통증·기타는 결제 대상 아니라 제외(시트 한약결제율 정의와 일치).
    각 초진을 자기 주차와 그 달 양쪽에 집계.
    """
    from collections import defaultdict
    agg = defaultdict(lambda: defaultdict(lambda: {"특화초진": 0, "그냥감": 0, "결제": 0}))
    for c in result.get("chojin", []):
        if c.get("group") not in ("피부", "호흡기"):
            continue
        doc = c.get("doctor") or "(미기입)"
        for p in _periods_of(c["week"]):
            a = agg[p][doc]
            a["특화초진"] += 1
            if c["outcome"] == "그냥감":
                a["그냥감"] += 1
            elif c["outcome"] in PAID_OUTCOMES:
                a["결제"] += 1
    out = {}
    for p, docs in agg.items():
        out[p] = {}
        for doc, a in docs.items():
            denom = a["특화초진"] - a["그냥감"]
            a["진료"] = denom
            a["결제전환율"] = round(a["결제"] / denom, 4) if denom else None
            out[p][doc] = a
    return out


def aggregate_counselors(result: dict) -> dict:
    """기간별 × 상담자 내원율·결제율. 반환 {period: {counselor: {...}}}.

    내원율 = 내원 / (내원+노쇼)  — 예약 잡은 사람이 진짜 왔나(노쇼 책임).
    결제율 = 결제 / (내원+노쇼)  — 그중 결제까지 이어졌나. (결제 ⊆ 내원이라 결제율 ≤ 내원율.)
    상담구분(진료상담/전화상담) 분리. 내원대기·데이터대기·판정불가는 분모서 제외.
    각 문의를 자기 주차와 그 달 양쪽에 집계. 수기보정(_노쇼보정) 반영됨.
    """
    from collections import defaultdict
    agg = defaultdict(lambda: defaultdict(lambda: {"예약완료": 0, "내원": 0, "노쇼": 0,
                                                   "결제": 0, "진료상담": 0, "전화상담": 0}))
    for r in result["rows"]:
        who = r.get("counselor") or "(미기입)"
        for p in _periods_of(r["week"]):
            a = agg[p][who]
            a["예약완료"] += 1
            ch = r.get("channel", "")
            if "전화" in ch:
                a["전화상담"] += 1
            elif ch:
                a["진료상담"] += 1
            if r["status"] == "전환":
                a["내원"] += 1
                if (r.get("matched") or {}).get("outcome") in PAID_OUTCOMES:
                    a["결제"] += 1
            elif r["status"] == "노쇼":
                a["노쇼"] += 1
    out = {}
    for p, cous in agg.items():
        out[p] = {}
        for who, a in cous.items():
            denom = a["내원"] + a["노쇼"]
            a["확정"] = denom
            a["내원율"] = round(a["내원"] / denom, 4) if denom else None
            a["결제율"] = round(a["결제"] / denom, 4) if denom else None
            out[p][who] = a
    return out


def _match(iq, chojin, window_days):
    """문의 iq 에 대응하는 초진 찾기. 반환 (초진|None, 방법)."""
    if not iq["time"]:
        return None, ""
    lo = iq["time"] - timedelta(days=1)
    hi = iq["time"] + timedelta(days=window_days)
    for c in chojin:
        if c["reg"] and not (lo <= c["reg"] <= hi):
            continue
        if iq["chart"] and c["chart"] and iq["chart"] == c["chart"]:
            return c, "차트"
        if iq["phone"] and iq["phone"] in c["phones"]:
            return c, "전화"
        if iq["name"] and c["name"] and iq["name"] == c["name"]:
            return c, "성명"
    return None, ""


def _week_tab_of(d, sample_tab):
    """날짜 d 가 속한 주 탭명 추정 'YY-MM-N주'."""
    wk = ((d.day - 1) // 7) + 1
    return f"{d.year % 100:02d}-{d.month:02d}-{wk}주"


def match_inquiries(sid: str, asof, tabs: list, *, window_days: int = DEFAULT_WINDOW_DAYS,
                    key_path: str = DEFAULT_KEY, default_counselor: str = "") -> dict:
    """예약완료 문의 각각을 전환/내원대기/노쇼/데이터대기로 판정.

    default_counselor: 문의 상담자칸이 비면 채울 지점 기본 상담자. 안산처럼 OKTAS에서
    상담자를 아예 안 적는(구조적 공란) 지점의 임시 조치. 빈 칸만 채우므로 실제 이름은
    보존됨(안산이 상담자를 적기 시작하면 그 이름이 우선).
    """
    if isinstance(asof, str):
        asof = _pdate(asof)
    chojin, inquiries, tabset = _load(sid, tabs, key_path, default_counselor)
    overrides = load_overrides(sid, key_path)

    rows, counts = [], {"전환": 0, "내원대기": 0, "노쇼": 0, "데이터대기": 0, "판정불가": 0}
    for iq in inquiries:
        if "예약완료" not in iq["result"]:
            continue
        method, matched = "", None
        if not iq["time"]:
            status, note = "판정불가", "상담시각 없음"
        else:
            c, method = _match(iq, chojin, window_days)
            deadline = iq["time"] + timedelta(days=window_days)
            if c:
                status, note = "전환", f"{c['week']} 초진 ({c['reg']}) · {method}매칭"
                matched = {"week": c["week"], "reg": str(c["reg"]), "outcome": c.get("outcome", "")}
            elif asof <= deadline:
                status, note = "내원대기", f"마감 {deadline} (D-{(deadline - asof).days})"
            else:
                need = {_week_tab_of(iq["time"], None), _week_tab_of(deadline, None)}
                if need <= tabset:
                    status, note = "노쇼", f"마감 {deadline} 지남 · 초진 미발견"
                else:
                    status, note = "데이터대기", f"마감 지났으나 초진주 미업로드: {sorted(need - tabset)}"

        # 수기보정: 자동판정 위에 덮어씀
        auto_status = status
        k = iq_key(iq)
        overridden = k in overrides
        if overridden:
            note = f"[수기보정←{status}] {note}"
            status = overrides[k]

        counts[status] = counts.get(status, 0) + 1
        rows.append({**iq, "status": status, "auto_status": auto_status, "note": note,
                     "method": method, "matched": matched, "key": k, "overridden": overridden})

    n = sum(counts.values())
    converted = counts["전환"]
    return {
        "asof": str(asof), "window_days": window_days,
        "예약완료수": n, "counts": counts,
        "확정전환율": (converted / (converted + counts["노쇼"])) if (converted + counts["노쇼"]) else None,
        "rows": rows,
        "chojin": chojin,
    }


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    sid = sys.argv[1] if len(sys.argv) > 1 else "1GScJEpb2frMwFpRlbw-2OtXLWe9RXIioGfCRw9mUnfI"
    asof = sys.argv[2] if len(sys.argv) > 2 else "2026-07-01"
    sh = _svc()
    meta = sh.get(spreadsheetId=sid, fields="sheets(properties(title))").execute()
    tabs = [s["properties"]["title"] for s in meta["sheets"] if "주" in s["properties"]["title"]]
    r = match_inquiries(sid, asof, tabs)
    print(f"asof={r['asof']} 윈도우={r['window_days']}일 | 예약완료 {r['예약완료수']}건")
    print("분포:", r["counts"], "| 확정전환율:",
          f"{r['확정전환율']*100:.0f}%" if r["확정전환율"] is not None else "-")
    print("\n사례:")
    for x in sorted(r["rows"], key=lambda v: v["status"]):
        nm = (x["name"][:1] + "*") if x["name"] else "무명"
        print(f"  [{x['status']}] {x['week']} {nm} ({x['disease']}) — {x['note']}")
