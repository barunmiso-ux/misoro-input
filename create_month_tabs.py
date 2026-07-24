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

from case_sheet_writer import _svc, aggregate_month
from create_week_tabs import BRANCH_SHEETS, STD_TPL


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

            # 집계 대상: 방금 만든 탭 or (--refresh 로) 기존 탭
            do_agg = created or (args.refresh and month_tab in titles_ids)
            if not do_agg:
                print(f"[{name}] {month_tab}: {status}")
                continue

            all_tabs = list(titles_ids.keys())
            if month_tab not in all_tabs:
                all_tabs.append(month_tab)   # 방금 생성분
            weeks = [t for t in all_tabs if t.startswith(month + "-") and "주" in t]
            if not weeks:
                print(f"[{name}] {month_tab}: {status} · ⚠️ 주간탭 없어 집계 생략")
                continue
            if dry:
                print(f"[{name}] {month_tab}: {status} · 집계 예정 (주간 {len(weeks)}개)")
                continue
            res = aggregate_month(sid, month, all_tabs, dry_run=False)
            print(f"[{name}] {month_tab}: {status} · 집계완료 "
                  f"초진 {res['초진합계']} · 문의 {res['문의합계']} (주간 {len(res['weeks'])}개)")
        except Exception as e:
            print(f"[{name}] {month_tab} 실패: {type(e).__name__}: {e}")


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    main()
