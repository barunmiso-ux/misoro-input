"""월간 탭 자동 생성 + 주간 데이터 집계 — 주간 자동화(create_week_tabs.py)의 월간판.

각 지점 시트에서 '{YY-MM}월' 탭이 없으면 _표준양식_주간(57열 캐논)을 복제해 만든 뒤,
그 달 주간탭(YY-MM-N주)의 초진·문의 환자행을 aggregate_month로 합쳐 채운다.
매출·총내원(157행↓)은 OKTAS 월말결산 몫이라 건드리지 않음(빈칸으로 남아 지점이 입력).

이미 월간탭이 있으면 기본 skip(수기 유지분 보호). --refresh 주면 기존 탭도 재집계(덮어씀).
집계는 clear-then-write라 재실행해도 중복 없이 그 달 합계로 갱신됨.

  python create_month_tabs.py                      # 지난달 dry-run
  python create_month_tabs.py --commit             # 지난달 실생성
  python create_month_tabs.py --month 26-06 --commit
  python create_month_tabs.py --month 26-06 --refresh --commit
"""
from __future__ import annotations

import argparse
import datetime
import re

from case_sheet_writer import _svc, aggregate_month, write_settlement
from create_week_tabs import BRANCH_SHEETS, STD_TPL

# 결산 직접입력 셀 (write_settlement 키 → 주간탭 읽기 셀)
SETTLE_CELLS = {"매출": "D158", "환불": "D159", "총내원": "D166", "신환내원": "D167"}


def _num(x) -> float:
    return float(re.sub(r"[^0-9.\-]", "", str(x)) or 0)


def sum_weekly_settlement(sh, sid: str, weeks: list) -> dict:
    """그 달 주간탭들의 결산(매출·환불·총내원·신환)을 합산. 월경계주는 이미 월별로
    클램프돼 있어(collect_settlement 기간결산) 이중계상 없음. 월간=주간합 검증필(<0.05%)."""
    keys = list(SETTLE_CELLS.keys())
    ranges = [f"'{w}'!{SETTLE_CELLS[k]}" for w in weeks for k in keys]
    vr = sh.values().batchGet(spreadsheetId=sid, ranges=ranges).execute().get("valueRanges", [])
    total = {k: 0.0 for k in keys}
    for i in range(len(weeks)):
        for j, k in enumerate(keys):
            cell = vr[len(keys) * i + j].get("values", [])
            total[k] += _num(cell[0][0] if cell and cell[0] else "")
    return {k: round(v) for k, v in total.items()}


def monthly_settlement_empty(sh, sid: str, month_tab: str) -> bool:
    """월간탭 매출(D158)이 비어있으면 True(=주간합으로 채울 대상). 기존 OKTAS 월말값 보호."""
    v = sh.values().get(spreadsheetId=sid, range=f"'{month_tab}'!D158").execute().get("values", [])
    return _num(v[0][0] if v and v[0] else "") == 0


def prev_month_prefix(today: datetime.date) -> str:
    """오늘 기준 '지난달' 접두어 'YY-MM'."""
    y, m = today.year, today.month
    m -= 1
    if m == 0:
        y -= 1
        m = 12
    return f"{y % 100:02d}-{m:02d}"


def ensure_month_tab(sh, sid, month_tab: str, titles_ids: dict, dry_run: bool) -> tuple:
    """월간탭 없으면 _표준양식_주간 복제로 생성. return (상태문구, created_bool)."""
    if month_tab in titles_ids:
        return "이미 있음", False
    if STD_TPL not in titles_ids:
        return f"⚠️ 표준템플릿({STD_TPL}) 없음 — 생성 불가", False
    if dry_run:
        return f"생성 예정 (← {STD_TPL})", True
    sh.batchUpdate(spreadsheetId=sid, body={"requests": [
        {"duplicateSheet": {"sourceSheetId": titles_ids[STD_TPL], "newSheetName": month_tab}}
    ]}).execute()
    return "✅ 생성됨", True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true")
    ap.add_argument("--month", help="대상 월 'YY-MM' (기본=지난달)")
    ap.add_argument("--refresh", action="store_true",
                    help="이미 있는 월간탭도 재집계(덮어씀)")
    args = ap.parse_args()

    month = args.month or prev_month_prefix(datetime.date.today())
    month_tab = f"{month}월"
    dry = not args.commit
    print(f"대상 월간탭: {month_tab}  ({'DRY-RUN' if dry else '실생성'}"
          f"{', 기존도 재집계' if args.refresh else ''})\n")

    sh = _svc()
    for name, sid in BRANCH_SHEETS.items():
        try:
            meta = sh.get(spreadsheetId=sid,
                          fields="sheets(properties(title,sheetId))").execute()
            titles_ids = {s["properties"]["title"]: s["properties"]["sheetId"]
                          for s in meta["sheets"]}
            status, created = ensure_month_tab(sh, sid, month_tab, titles_ids, dry)
            exists = created or (month_tab in titles_ids)

            all_tabs = list(titles_ids.keys())
            if month_tab not in all_tabs:
                all_tabs.append(month_tab)   # 방금 생성분
            weeks = [t for t in all_tabs if t.startswith(month + "-") and "주" in t]

            # ① 환자명단 집계: 방금 만든 탭 or (--refresh 로) 기존 탭
            do_agg = created or (args.refresh and month_tab in titles_ids)
            notes = []
            if do_agg:
                if not weeks:
                    notes.append("⚠️ 주간탭 없어 집계 생략")
                elif dry:
                    notes.append(f"집계 예정 (주간 {len(weeks)}개)")
                else:
                    res = aggregate_month(sid, month, all_tabs, dry_run=False)
                    notes.append(f"집계완료 초진 {res['초진합계']} · 문의 {res['문의합계']}")

            # ② 결산 채우기: 월간 결산이 비어있으면 주간합으로(기존 OKTAS 월말값은 보존)
            if exists and weeks:
                if dry:
                    if created:
                        notes.append("결산 주간합 예정")
                    elif monthly_settlement_empty(sh, sid, month_tab):
                        notes.append("결산 비어있음 → 주간합 예정")
                elif monthly_settlement_empty(sh, sid, month_tab):
                    s = sum_weekly_settlement(sh, sid, weeks)
                    write_settlement(sid, month_tab, s, dry_run=False)
                    notes.append(f"결산기록 매출 {s['매출']:,} · 신환 {s['신환내원']}")

            print(f"[{name}] {month_tab}: {status}"
                  + (" · " + " · ".join(notes) if notes else ""))
        except Exception as e:
            print(f"[{name}] {month_tab} 실패: {type(e).__name__}: {e}")


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    main()
