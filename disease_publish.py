# -*- coding: utf-8 -*-
"""질환별 문의·신환·결제 집계 → 캐시시트 `_질환` 탭 (원장/본사 대시보드 질환 뷰용).

각 지점 주간시트에서 읽는다:
  · 초진명단 C5:Y154 — 질환(VIP)·진행치료·비예약원인 → 신환수·결제수
  · 문의내역 AA5:AI139 — 진료구분(질환) → 문의수

⚠️문의 질환과 초진 질환은 자주 다르다(문의 '기타피부' → 초진 '사마귀').
  그래서 **질환별 문의→신환 전환율은 만들지 않는다.** 문의는 문의대로, 신환은 신환대로 센다.
  질환별 전환은 전화번호로 이어진 건만 초진 질환 기준으로 봐야 하며 그건 별도(noshow_matcher) 몫.

⚠️질환별로 쪼개면 표본이 급격히 작아진다(분당 7월 다한증 신환 2명).
  월간 행도 쌓되 화면은 3개월 누적을 기본으로 볼 것. MIN_N 미만은 순위·비교에서 빼야 한다.

결제 정의는 원장 지표와 동일:
  분모 = 결제대상 신환 − 그냥감 − 급성   (상담 기회가 있었던 건만)
  분자 = 한약·약침·첩약보험 결제

  python disease_publish.py             # dry-run
  python disease_publish.py --commit
"""
from __future__ import annotations

import argparse
import datetime
import re
from collections import defaultdict

from case_sheet_writer import _svc
from export_parser import (classify_treatment, _classify_disease, is_paid_target,
                           PAID_OUTCOMES, ACUTE_MARKERS, _ALIAS)
from noshow_publish import BRANCH_SHEETS, CACHE_SHEET

WEEK_TAB = re.compile(r"^\d{2}-\d{2}-\d주$")   # 26-07-3주 형식만

TAB = "_질환"
HEADERS = ["지점", "기간", "질환", "질환군", "문의", "신환", "결제대상", "결제", "결제율",
           "그냥감", "급성", "갱신"]
MIN_N = 5          # 이 미만이면 화면에서 순위·비교 제외 권장(행은 그대로 둔다)

# 초진 C5:Y154 → 0:이름 1:차트 11:등록일 15:질환 20:진행치료 22:비예약원인
CH_NAME, CH_CHART, CH_REG, CH_DISEASE, CH_TREAT, CH_REASON = 0, 1, 11, 15, 20, 22
# 문의 AA5:AI139 → 0:상담시각 6:진료구분(질환)
IQ_TIME, IQ_DISEASE = 0, 6


def _cell(row, i):
    return str(row[i]).strip() if i < len(row) else ""


def _norm(dis: str) -> str:
    """질환명 정규화 — 집계 키를 하나로 모은다.

    ⚠️_classify_disease 는 별칭을 흡수하지만 그건 '질환군' 판정용이고, 집계 키로 원본을
      쓰면 '알레르기비염'과 '알레르기비염(첩약보험)'이 따로 잡혀 결제율이 쪼개진다.
    """
    d = str(dis or '').strip()
    return _ALIAS.get(d, d)


def _isdate(v) -> bool:
    t = str(v or "").strip()
    return bool(re.match(r"^\d{4}[-./]\s?\d{1,2}[-./]\s?\d{1,2}", t)
                or re.match(r"^\d{1,2}/\d{1,2}/\d{2}", t))


def _periods(week_tab: str):
    """'26-07-3주' → ['26-07-3주', '26-07월'] — 주차·월 양쪽에 집계."""
    mo = week_tab.rsplit("-", 1)[0] + "월" if "-" in week_tab else week_tab
    return [week_tab, mo] if mo != week_tab else [week_tab]


def _colmap(header: list, want: dict) -> dict:
    """헤더행 → {키: 열index}. 헤더명 앞부분 일치로 찾는다.

    ⚠️고정 오프셋을 쓰면 안 된다. 구버전 주차탭은 컬럼 배치가 달라서 같은 index가
      엉뚱한 칸을 가리킨다(질환 자리에 유입경로 '블로그'·'지도'가 잡혔다, 2026-08 실측).
    """
    out = {}
    for i, h in enumerate(header):
        t = str(h).strip()
        for key, names in want.items():
            if key in out:
                continue
            if any(t.startswith(n) for n in names):
                out[key] = i
    return out


CH_WANT = {"차트": ["차트"], "등록일": ["등록일"], "질환": ["VIP"],
           "치료": ["진행치료"], "사유": ["URL"]}
IQ_WANT = {"시각": ["상담시간", "상담시각"], "질환": ["진료구분"]}


def collect(sid: str, tabs: list) -> dict:
    """{(기간, 질환): {문의·신환·결제대상·결제·그냥감·급성}}"""
    sh = _svc()
    meta = sh.get(spreadsheetId=sid,
                  fields="sheets(properties(title,gridProperties(columnCount)))").execute()
    width = {s["properties"]["title"]: s["properties"].get("gridProperties", {}).get("columnCount", 0)
             for s in meta.get("sheets", [])}
    ranges, has_iq = [], []
    for t in tabs:
        ranges.append(f"'{t}'!C4:Y154")          # 4행=헤더 포함해서 읽는다
        ok = width.get(t, 0) >= 35
        has_iq.append(ok)
        if ok:
            ranges.append(f"'{t}'!AA4:AI139")
    vrs = sh.values().batchGet(spreadsheetId=sid, ranges=ranges).execute().get("valueRanges", [])

    agg = defaultdict(lambda: {"문의": 0, "신환": 0, "결제대상": 0, "결제": 0,
                               "그냥감": 0, "급성": 0})
    vi = 0
    for t, iq_ok in zip(tabs, has_iq):
        vals = vrs[vi].get("values", [])
        cm = _colmap(vals[0], CH_WANT) if vals else {}
        if not {"차트", "질환", "치료"} <= set(cm):
            vi += 1 + (1 if iq_ok else 0)
            continue                              # 헤더를 못 읽는 비표준 탭 — 건너뛴다
        g = lambda row, k: _cell(row, cm[k]) if k in cm else ""
        for row in vals[1:]:
            ch, dis = g(row, "차트"), _norm(g(row, "질환"))
            # 실제 초진행은 숫자 차트 + **날짜 등록일** 둘 다 있다. 초진 범위(154행) 끝에
            # 붙어 있는 통계 요약행('총내원수'·'신환내원수')은 둘째 칸 숫자가 차트로 오인돼
            # 질환이 '0'인 유령 레코드를 만든다 — 등록일 검사로 걸러낸다.
            if not ch.isdigit() or not dis or not _isdate(g(row, "등록일")):
                continue
            outcome = classify_treatment(g(row, "치료"))
            acute = any(m in g(row, "사유") for m in ACUTE_MARKERS)
            for p in _periods(t):
                a = agg[(p, dis)]
                a["신환"] += 1
                if not is_paid_target(dis):
                    continue
                if outcome == "그냥감":
                    a["그냥감"] += 1
                elif acute:
                    a["급성"] += 1
                else:
                    a["결제대상"] += 1
                    if outcome in PAID_OUTCOMES:
                        a["결제"] += 1
        vi += 1
        if not iq_ok:
            continue
        ivals = vrs[vi].get("values", [])
        icm = _colmap(ivals[0], IQ_WANT) if ivals else {}
        if {"시각", "질환"} <= set(icm):
            for row in ivals[1:]:
                dis = _norm(_cell(row, icm["질환"]))
                if not dis or not _cell(row, icm["시각"]):
                    continue
                for p in _periods(t):
                    agg[(p, dis)]["문의"] += 1
        vi += 1
    return agg


def build_rows(asof: str):
    sh = _svc()
    rows = []
    for name, sid in BRANCH_SHEETS.items():
        try:
            meta = sh.get(spreadsheetId=sid, fields="sheets(properties(title))").execute()
            # ⚠️'주'만 걸면 템플릿('_표준양식_주간'·'<신규> 주간통계')과 천안의 2025년 구형탭
            #   ('25년12월4째주' 등)까지 섞여 엉뚱한 집계가 나온다. 정규 주차탭만 본다.
            tabs = [s["properties"]["title"] for s in meta["sheets"]
                    if WEEK_TAB.match(s["properties"]["title"])]
            if not tabs:
                continue
            agg = collect(sid, tabs)
            for (period, dis), a in sorted(agg.items()):
                grp = _classify_disease(dis)[0]
                rate = (a["결제"] / a["결제대상"]) if a["결제대상"] else None
                rows.append([name, period, dis, grp, a["문의"], a["신환"], a["결제대상"],
                             a["결제"], (f"{rate*100:.0f}%" if rate is not None else ""),
                             a["그냥감"], a["급성"], asof])
        except Exception as e:
            print(f"[{name}] 실패: {type(e).__name__}: {e}")
    return rows


def publish(rows):
    sh = _svc()
    titles = [s["properties"]["title"]
              for s in sh.get(spreadsheetId=CACHE_SHEET).execute()["sheets"]]
    if TAB not in titles:
        sh.batchUpdate(spreadsheetId=CACHE_SHEET,
                       body={"requests": [{"addSheet": {"properties": {"title": TAB}}}]}).execute()
    sh.values().clear(spreadsheetId=CACHE_SHEET, range=f"'{TAB}'!A1:L20000").execute()
    sh.values().update(spreadsheetId=CACHE_SHEET, range=f"'{TAB}'!A1",
                       valueInputOption="RAW", body={"values": [HEADERS] + rows}).execute()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true")
    args = ap.parse_args()
    asof = datetime.date.today().isoformat()
    rows = build_rows(asof)
    print(f"기준일 {asof} · {len(rows)}행 ({'실기록' if args.commit else 'DRY-RUN'})")
    cur = [r for r in rows if r[1].endswith("월")]
    cur.sort(key=lambda r: -r[5])
    print(f"\n{'지점':<6}{'기간':<10}{'질환':<12}{'문의':>4}{'신환':>4}{'대상':>4}{'결제':>4}{'결제율':>7}")
    for r in cur[:20]:
        print(f"{r[0]:<6}{r[1]:<10}{r[2][:10]:<12}{r[4]:>4}{r[5]:>4}{r[6]:>4}{r[7]:>4}{r[8]:>7}")
    if args.commit and rows:
        publish(rows)
        print(f"\n✅ {TAB} 탭 기록 {len(rows)}행")


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    main()
