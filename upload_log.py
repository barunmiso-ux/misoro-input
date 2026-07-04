"""중앙 업로드 로그 — 지점이 앱으로 기록에 성공할 때마다 캐시시트에 한 줄 append.

- 저장 위치: 대시보드 캐시시트(_NOSHOW 등과 같은 곳)의 `_업로드로그` 탭. 전 지점 통합.
- 기존 구조 무변경(새 탭 추가만) → 블래스트반경 0.
- 로그는 부가기능 → 실패해도 예외를 삼켜 '기록' 자체는 막지 않는다.
"""
from datetime import datetime, timezone, timedelta

from case_sheet_writer import _svc

# 중앙 로그 위치 = 전용 시트 '미소로 업로드로그'(사용자 생성, SA 편집자 공유). 전 지점 통합.
LOG_SHEET = "1xB0Ehb8KAUuV077UFJ09fiT3mlC4_7ZMXm2wSuyo4eo"
LOG_TAB = "_업로드로그"
HEADERS = ["시각", "지점", "구분", "주차", "추가", "갱신", "탭총계", "기록자"]
_KST = timezone(timedelta(hours=9))


def _now_kst() -> str:
    return datetime.now(_KST).strftime("%Y-%m-%d %H:%M")


def log_upload(branch: str, kind: str, week_tab: str, added, updated,
               total="", writer: str = "app", now: str = "") -> bool:
    """업로드 1건(지점×구분×주차) 로그. 성공 True / 실패(무시) False."""
    try:
        sh = _svc()
        meta = sh.get(spreadsheetId=LOG_SHEET,
                      fields="sheets(properties(title))").execute()
        titles = {s["properties"]["title"] for s in meta["sheets"]}
        if LOG_TAB not in titles:
            sh.batchUpdate(spreadsheetId=LOG_SHEET, body={"requests": [
                {"addSheet": {"properties": {"title": LOG_TAB}}}]}).execute()
            sh.values().update(spreadsheetId=LOG_SHEET, range=f"'{LOG_TAB}'!A1",
                               valueInputOption="RAW",
                               body={"values": [HEADERS]}).execute()
        row = [now or _now_kst(), branch, kind, week_tab, added, updated, total, writer]
        sh.values().append(
            spreadsheetId=LOG_SHEET, range=f"'{LOG_TAB}'!A:H",
            valueInputOption="RAW", insertDataOption="INSERT_ROWS",
            body={"values": [row]}).execute()
        return True
    except Exception:
        return False


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    ok = log_upload("테스트", "초진", "26-07-1주", 0, 0, 0, writer="cli-test")
    print("로그 append:", "성공" if ok else "실패")
