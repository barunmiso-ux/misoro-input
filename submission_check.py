"""초진/문의 제출현황 체커 → 미제출 지점 잔디 촉구 (기존 인프라 재활용).

기존 일일보고 알림 시스템(앱스스크립트)이 쓰는 것을 그대로 재활용:
- 웹훅 URL: 집계시트(1uTkikV)의 '지점설정' 탭 (지점별 잔디 Incoming Webhook)
- 휴진: '휴진관리' 탭 (정기 매주/날짜) — 단 주간 체크는 OKTAS 신환=0 으로 대부분 커버

판정: OKTAS 신환수(주간탭 D167, collect_settlement가 채움) = "올려야 할 초진 수".
  신환 > 0 인데 초진 테이블 비어있음 → 미제출 → 잔디 촉구.
  신환 = 0 → 올릴 것 없음(휴진/한산) → 조용.
  신환 미집계(결산 아직) → 보류.

  python submission_check.py                       # 지난주 dry-run
  python submission_check.py --commit              # 실제 잔디 발송
  python submission_check.py --date 2026-06-22
"""
from __future__ import annotations

import argparse
import datetime
import re
import time

import requests

from case_sheet_writer import _svc
from export_parser import week_label
import week_rule

CONFIG_SHEET = "1uTkikVDCUfVry6l-GX2L8yvlhRSVFdUtEwlubrhpZLk"  # 지점설정·휴진관리(일일보고 앱스스크립트와 동일)
APP_URL = "https://misoro-stats.streamlit.app"  # 주간통계 자동입력 앱 (배포 시 이 슬러그로)

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


def load_webhooks() -> dict:
    """지점설정 탭 → {지점: 잔디웹훅URL}."""
    sh = _svc()
    vals = sh.values().get(spreadsheetId=CONFIG_SHEET, range="'지점설정'!A2:B").execute().get("values", [])
    out = {}
    for r in vals:
        if len(r) >= 2 and r[0].strip() and r[1].strip().startswith("http"):
            out[r[0].strip()] = r[1].strip()
    return out


def send_jandi(url: str, message: str, color: str = "#FAC11B"):
    """웹훅 URL 형식 자동 판별: 구글챗이면 챗 형식({text}), 아니면 잔디 형식.
    (잔디→구글챗 이전 중 — 지점설정 탭 URL만 바꾸면 형식은 여기서 알아서 처리)"""
    if "chat.googleapis.com" in url:
        requests.post(url, json={"text": message},
                      headers={"Content-Type": "application/json"}, timeout=10)
    else:
        requests.post(url, json={"body": message, "connectColor": color},
                      headers={"Accept": "application/vnd.tosslab.jandi-v2+json"}, timeout=10)


def _has_tab(sh, sid, tab):
    meta = sh.get(spreadsheetId=sid, fields="sheets(properties(title))").execute()
    return tab in [s["properties"]["title"] for s in meta.get("sheets", [])]


def _nonempty(vals):
    return sum(1 for r in vals if r and str(r[0]).strip())


def _num(cell):
    return int("".join(ch for ch in str(cell) if ch.isdigit()) or 0)


def _reg_date(s):
    m = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", str(s))
    if not m:
        return None
    try:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def _filter_monthly_held(sh, sid, existing, misses):
    """miss 주차 중 '초진 미제출'이지만 그 주 환자가 월간탭에 등록돼 있으면(=월경계 조각주·수기월간)
    진짜 누락 아님 → 제외. 반환 (진짜누락[], 월간보유[(주,인원)])."""
    month_of = {wt: wt.rsplit("-", 1)[0] + "월" for wt, _, mm in misses if "초진" in mm}
    mtabs = sorted({m for m in month_of.values() if m in existing})
    if not mtabs:
        return misses, []
    rngs = []
    for mt in mtabs:
        rngs += [f"'{mt}'!C5:C154", f"'{mt}'!N5:N154"]
    vr = sh.values().batchGet(spreadsheetId=sid, ranges=rngs).execute().get("valueRanges", [])
    monthly = {}   # 월간탭 -> [등록일]
    for j, mt in enumerate(mtabs):
        names = vr[j*2].get("values", [])
        regs = vr[j*2+1].get("values", [])
        dts = []
        for k in range(len(names)):
            if names[k] and str(names[k][0]).strip():
                d = _reg_date(regs[k][0]) if k < len(regs) and regs[k] else None
                if d:
                    dts.append(d)
        monthly[mt] = dts
    real, held = [], []
    for wt, exp, mm in misses:
        mt = month_of.get(wt)
        if mt and mt in monthly:
            lo, hi = week_rule.range_for_tab(wt)
            cnt = sum(1 for d in monthly[mt] if lo <= d <= hi)
            if cnt > 0:
                held.append((wt, cnt))
                continue
        real.append((wt, exp, mm))
    return real, held


def check(week_tabs, commit: bool = False):
    """week_tabs = 검사할 주차 목록(최근 N주). 지점당 1콜 배치로 읽고 누락 주차 모아 통합 알림."""
    if isinstance(week_tabs, str):
        week_tabs = [week_tabs]
    sh = _svc()
    webhooks = load_webhooks()
    print(f"대상 주: {', '.join(week_tabs)}  ({'실발송' if commit else 'DRY-RUN'})\n")

    for name, sid in BRANCH_SHEETS.items():
        try:
            meta = sh.get(spreadsheetId=sid, fields="sheets(properties(title))").execute()
            existing = {s["properties"]["title"] for s in meta.get("sheets", [])}
            present = [wt for wt in week_tabs if wt in existing]
            misses = []   # [(week_tab, exp, missing)]
            if present:
                ranges = []
                for wt in present:
                    ranges += [f"'{wt}'!D167", f"'{wt}'!C5:C150", f"'{wt}'!AC5:AC139"]
                vr = sh.values().batchGet(spreadsheetId=sid, ranges=ranges).execute().get("valueRanges", [])
                for i, wt in enumerate(present):
                    exp = _num((vr[i*3].get("values") or [[""]])[0][0] if vr[i*3].get("values") else "")
                    if exp == 0:                      # 신환0 = 올릴 것 없음(조용)
                        continue
                    chojin = _nonempty(vr[i*3+1].get("values", []))
                    munui = _nonempty(vr[i*3+2].get("values", []))
                    mm = (["초진"] if chojin == 0 else []) + (["문의"] if munui == 0 else [])
                    if mm:
                        misses.append((wt, exp, mm))
            time.sleep(1)                             # 레이트리밋 여유
            # 월간 교차확인: 주간탭 비었어도 그 주 환자가 월간탭에 등록돼 있으면 알림 제외
            if misses:
                misses, held = _filter_monthly_held(sh, sid, existing, misses)
                for wt, cnt in held:
                    print(f"    ℹ️ {name} {wt}: 주간탭 비었으나 월간탭에 {cnt}명 등록됨 → 알림 제외")
                time.sleep(1)
            if not misses:
                print(f"[{name}] 최근 {len(week_tabs)}주 제출 정상 ✅ (월간보유/신환0 포함)")
                continue
            lines = [f"· {week_label(wt)}" for wt, _, _ in misses]
            msg = (f"[{name}] 업로드 기록이 없는 주간 통계입니다.\n"
                   + "\n".join(lines)
                   + f"\n'주간통계 자동입력'에 명단 올린 뒤 기록 버튼까지 눌러주세요.\n"
                   + f"{APP_URL}/?branch={name}")
            url = webhooks.get(name)
            print(f"[{name}] 미제출 {len(misses)}주 {[w for w,_,_ in misses]} → 잔디 "
                  + ("발송" if commit and url else "대상"))
            if commit and url:
                send_jandi(url, msg)
            elif not url:
                print(f"    ⚠️ {name} 웹훅 URL 없음(지점설정 탭)")
        except Exception as e:
            print(f"[{name}] 확인 실패: {type(e).__name__}: {e}")


def recent_week_tabs(n_weeks: int) -> list:
    """최근 N개 완료주가 걸치는 주차탭(월경계 분할 포함)."""
    today = datetime.date.today()
    this_mon = today - datetime.timedelta(days=today.weekday())
    tabs = []
    for i in range(1, n_weeks + 1):
        base = this_mon - datetime.timedelta(days=7 * i)
        for off in range(7):
            t = week_rule.tab_name(base + datetime.timedelta(days=off))
            if t not in tabs:
                tabs.append(t)
    return tabs


def last_monday_week_tab():
    today = datetime.date.today()
    mon = today - datetime.timedelta(days=today.weekday() + 7)
    return week_rule.tab_name(mon)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true")
    ap.add_argument("--date", help="특정 주 월요일 YYYY-MM-DD만 검사 (기본=최근 N주)")
    ap.add_argument("--weeks", type=int, default=4, help="최근 몇 주 검사(기본 4) — 놓친 옛주도 잡음")
    args = ap.parse_args()
    if args.date:
        tabs = [week_rule.tab_name(datetime.date.fromisoformat(args.date))]
    else:
        tabs = recent_week_tabs(args.weeks)
    check(tabs, args.commit)


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    main()
