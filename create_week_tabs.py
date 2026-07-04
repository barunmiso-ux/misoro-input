"""주차 탭 자동 생성 — 인간 개입 최소화.

각 지점 시트에서 '빈 주간 템플릿'을 복사(DuplicateSheet)해 새 주차 탭을 만든다.
지점이 매주 수동으로 템플릿 복사하던 걸 대체 → export만 올리면 됨.

템플릿 = 이름에 '신규' 있고 '월간' 없고, 초진 C5가 비어있는 탭(예 '<26.05> 신규 통계').
매주 월요일 스케줄로 그 주 탭을 전 지점에 생성(이미 있으면 skip).

  python create_week_tabs.py                    # 이번주 dry-run
  python create_week_tabs.py --commit
  python create_week_tabs.py --date 2026-07-06  # 그 날이 속한 주
"""
from __future__ import annotations

import argparse
import datetime

from case_sheet_writer import _svc


def week_of_month(day: int) -> int:
    """월 내 주차: 1주=1~7, 2주=8~14, 3주=15~21, 4주=22~28, 5주=29~말일.
    (export_parser.week_of_month 과 동일 — CI 경량화 위해 인라인, pandas 의존 제거.)"""
    return (day - 1) // 7 + 1

BRANCH_SHEETS = {
    "분당": "1GScJEpb2frMwFpRlbw-2OtXLWe9RXIioGfCRw9mUnfI",
    "천안": "1f8pef-FQ58e5eEzzWOkWcx_FJcqZCEdzSquHQLjTg9U",
    "전주": "16mnhArKOUhG0oMztsaN5bDotESBKpRDoCWmWrpElSIw",
    "안산": "1t7iWIYiUQPKMAB9XO345HONc8XbLH0MgeGClIcSRkko",
    "수원": "1X13A_jJ5ejeStgeg9A7pt5DXLqiHgm6mwpphxhlcKZk",
    "평택": "1pOHJZ8_3eQyXXEGQ31MOwt5kWtcbqAAv9oQxMDDEaVI",
    "인천": "1KW1KjaBZ8RZJtoANDwqb21NB1uNOere4iRaVxsE21ko",
    "영등포": "1cq_34IETdHTyD-FLe_lHJdK52o88rEEQuOOvp7uWk4I",
    "대전": "1rBCCeB8S-NJ9z2ApHGyXhVK_NScoYHauSYZBOl1KQ0c",
}


def week_tab_name(d: datetime.date) -> str:
    """날짜 d 의 주차 탭명 'YY-MM-N주'.

    앱(export_parser._week_tab_of)과 **동일 규칙**: 그 날짜의 '월 + week_of_month(일)'.
    월요일 기준 아님 — 앱이 레코드를 날짜별로 이 이름의 탭에 넣으므로 탭명이 반드시 일치해야 한다.
    (예: 7/5 → 26-07-1주. 월요일 기준이면 6/29→26-06-5주가 되어 앱과 불일치했음.)
    """
    return f"{d.year % 100:02d}-{d.month:02d}-{week_of_month(d.day)}주"


import re

_TPL_BAD = ("월간", "샘플", "체크", "양식", "디버그")
_WEEK_RE = re.compile(r"-(\d+)주$")


STD_TPL = "_표준양식_주간"  # 표준화로 전 지점에 심은 57열 캐논 템플릿(2026-07-01)


def _source_tab(sh, sid, titles_ids: dict) -> tuple:
    """복사 원본 (title, gid, mode). 표준템플릿 최우선 → 신규템플릿 → 최신주차+클리어."""
    # ⓪ 표준화 템플릿 _표준양식_주간 (전 지점 통일, 57열 blank)
    if STD_TPL in titles_ids:
        return STD_TPL, titles_ids[STD_TPL], "template"
    # ① 깨끗한 템플릿: '신규' 있고 나쁜키워드 없고 C5 비어있음
    tpl = [(t, g) for t, g in titles_ids.items()
           if "신규" in t and not any(b in t for b in _TPL_BAD)]
    for title, gid in sorted(tpl, reverse=True):
        c5 = sh.values().get(spreadsheetId=sid, range=f"'{title}'!C5").execute().get("values", [])
        if not c5 or not str(c5[0][0]).strip():
            return title, gid, "template"
    # ② 폴백: 최신 'N주' 탭 복사 후 데이터 클리어
    weeks = [(t, g) for t, g in titles_ids.items() if _WEEK_RE.search(t)]
    if weeks:
        title, gid = sorted(weeks, reverse=True)[0]
        return title, gid, "copy_clear"
    return None, None, None


def _clear_new_tab(sh, sid, tab: str):
    """복사한 탭의 데이터 클리어(수식 보존): 초진 C5:Y154·문의 AA5:AK139·결산입력."""
    for rng in (f"'{tab}'!C5:Y154", f"'{tab}'!AA5:AK139",
                f"'{tab}'!D158", f"'{tab}'!D159", f"'{tab}'!D166", f"'{tab}'!D167"):
        sh.values().clear(spreadsheetId=sid, range=rng).execute()


def ensure_week_tab(sh, sid, week_tab: str, dry_run: bool = True) -> str:
    meta = sh.get(spreadsheetId=sid, fields="sheets(properties(title,sheetId))").execute()
    titles_ids = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}
    if week_tab in titles_ids:
        return "이미 있음"
    src_title, src_gid, mode = _source_tab(sh, sid, titles_ids)
    if src_gid is None:
        return "⚠️ 원본(템플릿/주차탭) 없음"
    tag = "템플릿" if mode == "template" else "최신주차복사+클리어"
    if dry_run:
        return f"생성 예정 (← '{src_title}', {tag})"
    sh.batchUpdate(spreadsheetId=sid, body={"requests": [
        {"duplicateSheet": {"sourceSheetId": src_gid, "newSheetName": week_tab}}
    ]}).execute()
    if mode == "copy_clear":
        _clear_new_tab(sh, sid, week_tab)
    return f"✅ 생성됨 (← '{src_title}', {tag})"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true")
    ap.add_argument("--date", help="대상 주에 속한 날짜 YYYY-MM-DD (기본=오늘)")
    args = ap.parse_args()

    d = datetime.date.fromisoformat(args.date) if args.date else datetime.date.today()
    week_tab = week_tab_name(d)
    print(f"대상 주차 탭: {week_tab}  ({'실생성' if args.commit else 'DRY-RUN'})\n")

    sh = _svc()
    for name, sid in BRANCH_SHEETS.items():
        try:
            print(f"[{name}] {ensure_week_tab(sh, sid, week_tab, dry_run=not args.commit)}")
        except Exception as e:
            print(f"[{name}] 실패: {type(e).__name__}: {e}")


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    main()
