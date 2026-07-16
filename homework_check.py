# -*- coding: utf-8 -*-
"""카페 숙제 3주기 컴플라이언스 체크 → 지점 잔디 촉구.

주기:
  daily   — 게시글 1 + 댓글 2 미달 (그날)
  weekly  — 전문가 칼럼 주 3회 미달 (완료된 지난주 월~일)
  monthly — 상담실 답변 월 2회 미달 (완료된 지난달)

알림은 START_DATE(다음주 월요일)부터 유효 — 그 전 기간/날짜는 조용(롤아웃 유예).
웹훅·발송은 submission_check 재사용. dry-run 기본, --commit 로 실제 발송.

사용:
  python homework_check.py daily   [--commit] [--date YYYY-MM-DD]
  python homework_check.py weekly  [--commit]
  python homework_check.py monthly [--commit]
"""
import sys
import datetime
sys.stdout.reconfigure(encoding="utf-8")

import config
from submission_check import load_webhooks, send_jandi, _svc, CONFIG_SHEET

HOMEWORK_TAB = config.HOMEWORK_TAB
APP_URL = "https://misoro-input.streamlit.app"
BRANCHES = config.ALL_BRANCHES                       # 9지점(부산 제외)
START_DATE = datetime.date(2026, 7, 20)              # 알림 시작(다음주 월요일)

# 카페숙제체크 컬럼: 0날짜 1지점 2칼럼카페 3칼럼홈페 4게시글 5댓글수 6댓글URL 7입력자 8기록시각 9상담답변
C_DATE, C_BRANCH, C_COL_CAFE, C_COL_HOME, C_POST, C_CMTN, C_CONSULT = 0, 1, 2, 3, 4, 5, 9


def _kst_today():
    return datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).date()


def _cell(r, i):
    return r[i].strip() if i < len(r) and r[i] is not None else ""


def _rows():
    sh = _svc()
    return sh.values().get(spreadsheetId=CONFIG_SHEET,
                           range=f"'{HOMEWORK_TAB}'!A:J").execute().get("values", [])


def _int(s):
    try:
        return int(str(s).strip() or 0)
    except ValueError:
        return 0


def _send(name, msg, webhooks, commit):
    url = webhooks.get(name)
    print(f"[{name}] 미달 → 잔디 " + ("발송" if commit and url else "대상"))
    print(f"    {msg.splitlines()[0]}")
    if commit and url:
        send_jandi(url, msg)
    elif commit and not url:
        print(f"    ⚠️ {name} 웹훅 URL 없음(지점설정 탭)")


def check_daily(commit=False, target=None):
    target = target or _kst_today()
    if target < START_DATE:
        print(f"daily {target}: 알림 시작({START_DATE}) 전 → 조용"); return
    rows = _rows()
    webhooks = load_webhooks()
    ds = target.isoformat()
    by = {_cell(r, C_BRANCH): r for r in rows[1:] if _cell(r, C_DATE) == ds}
    for name in BRANCHES:
        r = by.get(name)
        post = 1 if (r and _cell(r, C_POST)) else 0
        cmt = _int(r[C_CMTN]) if r else 0
        if post >= 1 and cmt >= 2:
            print(f"[{name}] 게시글 {post}/1 · 댓글 {cmt}/2 ✅"); continue
        msg = (f"🔔 [{name}] {target:%m/%d} 카페 숙제 미완료\n"
               f"· 게시글 {post}/1 · 댓글 {cmt}/2\n"
               f"→ 일일보고에서 올려주세요: {APP_URL}/?branch={name}")
        _send(name, msg, webhooks, commit)


def check_weekly(commit=False, today=None):
    today = today or _kst_today()
    wk_start = today - datetime.timedelta(days=today.weekday() + 7)   # 지난주 월요일
    wk_end = wk_start + datetime.timedelta(days=6)
    if wk_start < START_DATE:
        print(f"weekly {wk_start}~{wk_end}: 알림 시작 전 → 조용"); return
    rows = _rows()
    webhooks = load_webhooks()
    cnt = {b: 0 for b in BRANCHES}
    for r in rows[1:]:
        b = _cell(r, C_BRANCH)
        if b not in cnt:
            continue
        try:
            rd = datetime.date.fromisoformat(_cell(r, C_DATE))
        except ValueError:
            continue
        if wk_start <= rd <= wk_end and (_cell(r, C_COL_CAFE) or _cell(r, C_COL_HOME)):
            cnt[b] += 1
    for name in BRANCHES:
        n = cnt[name]
        if n >= 3:
            print(f"[{name}] 전문가 칼럼 {n}/3회 ✅"); continue
        msg = (f"📅 [{name}] 지난주({wk_start:%m/%d}~{wk_end:%m/%d}) 전문가 칼럼 {n}/3회\n"
               f"· {3 - n}회 더 필요 (카페+홈페이지 병행)")
        _send(name, msg, webhooks, commit)


def check_monthly(commit=False, today=None):
    today = today or _kst_today()
    first_this = today.replace(day=1)
    last_month_end = first_this - datetime.timedelta(days=1)
    ym = (last_month_end.year, last_month_end.month)
    if last_month_end.replace(day=1) < START_DATE:
        print(f"monthly {ym}: 알림 시작 전 → 조용"); return
    rows = _rows()
    webhooks = load_webhooks()
    cnt = {b: 0 for b in BRANCHES}
    for r in rows[1:]:
        b = _cell(r, C_BRANCH)
        if b not in cnt:
            continue
        try:
            rd = datetime.date.fromisoformat(_cell(r, C_DATE))
        except ValueError:
            continue
        if (rd.year, rd.month) == ym and _cell(r, C_CONSULT):
            cnt[b] += 1
    for name in BRANCHES:
        n = cnt[name]
        if n >= 2:
            print(f"[{name}] 상담실 답변 {n}/2회 ✅"); continue
        msg = (f"🗓️ [{name}] 지난달({ym[0]}-{ym[1]:02d}) 상담실 답변 {n}/2회\n"
               f"· {2 - n}회 더 필요 (지정일 작성)")
        _send(name, msg, webhooks, commit)


def main():
    args = sys.argv[1:]
    mode = args[0] if args else "daily"
    commit = "--commit" in args
    target = None
    if "--date" in args:
        target = datetime.date.fromisoformat(args[args.index("--date") + 1])
    print(f"=== 숙제 체크: {mode} {'(발송)' if commit else '(dry-run)'} · 시작일 {START_DATE} ===")
    if mode == "daily":
        check_daily(commit, target)
    elif mode == "weekly":
        check_weekly(commit)
    elif mode == "monthly":
        check_monthly(commit)
    else:
        print("모드: daily | weekly | monthly")


if __name__ == "__main__":
    main()
