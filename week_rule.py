"""공용 주차 규칙 (사용자 확정 · 전 앱 단일 기준).

규칙:
  - 주 = 월~일.
  - 'N주' = 그 달을 월요일 시작 달력으로 깔았을 때의 주 행 번호(첫 주 = 1일~첫 일요일).
    week_of_month(d) = (d.day + 그달1일의요일 - 1)//7 + 1   (월=0)
  - 주가 두 달에 걸치면 월 경계에서 분할한다.
      예) 6/29(월)~7/5(일)  →  6월5주 = 6/29~6/30,  7월1주 = 7/1~7/5
  - 탭명 형식: 'YY-MM-N주'  (예 '26-07-1주')

이 모듈이 유일한 주차 기준. collect_settlement / export_parser / weekly_inquiry /
check_inquiry_consistency 가 전부 여기에 물린다.
"""
import calendar
import datetime
import re


def _first_weekday(y: int, m: int) -> int:
    """그 달 1일의 요일 (월=0 … 일=6)."""
    return datetime.date(y, m, 1).weekday()


def week_of_month(d: datetime.date) -> int:
    """날짜 → 그 달의 주차 번호 (달력 주행, 월요일 시작)."""
    return (d.day + _first_weekday(d.year, d.month) - 1) // 7 + 1


def tab_name(d: datetime.date) -> str:
    """날짜 → 주차 탭명 'YY-MM-N주'."""
    return f"{d.year % 100:02d}-{d.month:02d}-{week_of_month(d)}주"


def parse_tab(tab: str):
    """'YY-MM-N주' → (year, month, n). 주차 탭 아니면 None."""
    m = re.match(r"\s*(\d{2})-(\d{2})-(\d+)주\s*$", tab)
    if not m:
        return None
    return 2000 + int(m.group(1)), int(m.group(2)), int(m.group(3))


def range_for_tab(tab: str):
    """'YY-MM-N주' → (start_date, end_date). 월 경계로 클램프. 그 달에 없는 주차/형식오류면 None."""
    p = parse_tab(tab)
    if not p:
        return None
    y, m, n = p
    fw = _first_weekday(y, m)
    last = calendar.monthrange(y, m)[1]
    start_day = max(1, (n - 1) * 7 - fw + 1)
    end_day = min(last, n * 7 - fw)
    if start_day > last or end_day < 1 or start_day > end_day:
        return None
    return datetime.date(y, m, start_day), datetime.date(y, m, end_day)


def tabs_for_range(start: datetime.date, end: datetime.date):
    """[start,end] 날짜를 주차 탭별로 그룹 → [(tab, (t_start, t_end)), ...] (시작일 오름차순).
    경계 넘는 물리적 주는 여기서 2개 탭으로 자동 분할된다."""
    groups = {}
    d = start
    while d <= end:
        t = tab_name(d)
        if t not in groups:
            groups[t] = [d, d]
        else:
            groups[t][1] = d
        d += datetime.timedelta(days=1)
    return [(t, (v[0], v[1])) for t, v in sorted(groups.items(), key=lambda kv: kv[1][0])]


def last_completed_week(today: datetime.date = None):
    """직전 완료된 물리적 주(월~일)의 (월요일, 일요일)."""
    today = today or datetime.date.today()
    this_mon = today - datetime.timedelta(days=today.weekday())
    mon = this_mon - datetime.timedelta(days=7)
    return mon, mon + datetime.timedelta(days=6)


def completed_week_tabs(today: datetime.date = None):
    """직전 완료 주를 주차 탭(들)로 분할. 경계주면 2개. → [(tab, (start,end)), ...]."""
    mon, sun = last_completed_week(today)
    return tabs_for_range(mon, sun)
