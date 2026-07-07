"""구글 시트 읽기/쓰기 — 구글 시트 values API 직접 호출.

이 시트는 탭이 68개라 gspread open_by_key(전체 탭 메타데이터 fetch)가 ~5초 걸린다.
우리가 쓰는 탭만 values API 로 직접 읽고/쓰면 메타데이터 fetch 없이 ~1~2초로 끝난다.

시트 구조 (2026-06 확인):
- 상단 요약(1~15행)은 수식 → 직접 쓰지 않는다.
- "Daily Report" 블록에 (지점 열그룹 × 날짜 행) 으로 일별 카운트 기록 → 요약/합계 자동 갱신.
  · 블록1: 대전·부산·분당·수원·안산 / 블록2: 영등포·인천·전주·천안·평택
  · 분당 = O~U열 (행=날짜 'YYYY-MM-DD')
- 카페숙제체크: (날짜,지점) upsert.
- 초진상세: (날짜,지점) 교체(읽고-필터-재기록, sheetId 불필요).
"""

from __future__ import annotations

import os
from datetime import datetime, timedelta, timezone

from google.oauth2.service_account import Credentials

import config

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
_LOCAL_KEY_FILE = "service_account.json"
_KST = timezone(timedelta(hours=9))

_SVC = None  # Sheets API 서비스 (프로세스 내 캐시)


# ──────────────────────────────────────────────────────────────────
# 인증 / values API 래퍼
# ──────────────────────────────────────────────────────────────────
def _creds():
    if os.path.exists(_LOCAL_KEY_FILE):
        return Credentials.from_service_account_file(_LOCAL_KEY_FILE, scopes=SCOPES)
    import streamlit as st

    return Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )


def _service():
    global _SVC
    if _SVC is None:
        from googleapiclient.discovery import build

        _SVC = build("sheets", "v4", credentials=_creds(), cache_discovery=False)
    return _SVC


def _q(title: str, a1: str) -> str:
    return f"'{title}'!{a1}"


def _vget(rng: str):
    return (
        _service().spreadsheets().values()
        .get(spreadsheetId=config.SHEET_ID, range=rng).execute(num_retries=3)
        .get("values", [])
    )


def _vbatch_get(ranges: list):
    res = (
        _service().spreadsheets().values()
        .batchGet(spreadsheetId=config.SHEET_ID, ranges=ranges).execute(num_retries=3)
    )
    return [vr.get("values", []) for vr in res.get("valueRanges", [])]


def _vupdate(rng: str, values: list):
    _service().spreadsheets().values().update(
        spreadsheetId=config.SHEET_ID, range=rng,
        valueInputOption="USER_ENTERED", body={"values": values},
    ).execute(num_retries=3)


def _vappend(rng: str, values: list):
    _service().spreadsheets().values().append(
        spreadsheetId=config.SHEET_ID, range=rng,
        valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS",
        body={"values": values},
    ).execute(num_retries=3)


def _vclear(rng: str):
    _service().spreadsheets().values().clear(
        spreadsheetId=config.SHEET_ID, range=rng
    ).execute(num_retries=3)


def _create_tab(title: str, headers: list):
    _service().spreadsheets().batchUpdate(
        spreadsheetId=config.SHEET_ID,
        body={"requests": [{"addSheet": {"properties": {"title": title}}}]},
    ).execute(num_retries=3)
    _vupdate(_q(title, "A1"), [headers])


# ──────────────────────────────────────────────────────────────────
# 위치 찾기 (순수 함수: 그리드 입력)
# ──────────────────────────────────────────────────────────────────
def _col_letter(idx0: int) -> str:
    s = ""
    n = idx0 + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _cell(grid, r, c) -> str:
    if r < len(grid) and c < len(grid[r]):
        return grid[r][c].strip()
    return ""


def build_column_map(grid, branch: str):
    """Daily Report 블록에서 branch 의 항목 컬럼 위치. 반환 (item_header_row, {라벨:col}) (0-based)."""
    for ri, row in enumerate(grid):
        if _cell(grid, ri, 1) != "지점":
            continue
        # 블록 경계 = '지점' 헤더행(C열~)의 '비지 않은 셀 전부'.
        # ALL_BRANCHES 로만 잡으면 탈퇴지점(부산 등, 템플릿엔 열이 남음)이 경계에서
        # 빠져 → 그 앞 지점 블록이 탈퇴지점 열까지 넓어지고 colmap(dict)이 같은 라벨을
        # 그 열로 덮어써 데이터가 탈퇴지점 칸에 기록되는 버그가 남. 남은 열도 경계로 인식.
        branch_cols = [ci for ci in range(2, len(row)) if row[ci].strip()]
        if not branch_cols:
            continue
        for k, ci in enumerate(branch_cols):
            if row[ci].strip() != branch:
                continue
            start = ci
            hdr_row = ri + 1
            if k + 1 < len(branch_cols):
                end = branch_cols[k + 1]
            else:
                end = start
                while _cell(grid, hdr_row, end):
                    end += 1
            colmap = {}
            for c in range(start, end):
                label = _cell(grid, hdr_row, c)
                if label:
                    colmap[label] = c
            return hdr_row, colmap
    raise ValueError(f"지점 '{branch}' 를 일일 시트의 Daily Report 에서 찾지 못했습니다.")


def find_date_row(grid, hdr_row_idx: int, report_date) -> int:
    target = report_date.isoformat()
    for ri in range(hdr_row_idx + 1, len(grid)):
        b = _cell(grid, ri, 1)
        if b == target:
            return ri
        if "합계" in b or b == "지점":
            break
    raise ValueError(f"날짜 '{target}' 행을 일일 시트에서 찾지 못했습니다.")


def _now_str() -> str:
    return datetime.now(_KST).strftime("%Y-%m-%d %H:%M:%S")


# ──────────────────────────────────────────────────────────────────
# 읽기 (수정 모드 prefill)
# ──────────────────────────────────────────────────────────────────
def read_existing(branch: str, report_date) -> dict | None:
    """(branch, report_date) 기존 기록. 없으면 None. daily 한 번 + 카페숙제 한 번(batchGet)."""
    daily_name = config.daily_sheet_name(report_date)
    grid, hw = _vbatch_get([_q(daily_name, "A:AG"), _q(config.HOMEWORK_TAB, "A:I")])

    out = {"daily": {}, "activity": None}
    try:
        hdr_row, colmap = build_column_map(grid, branch)
        date_row = find_date_row(grid, hdr_row, report_date)
        for label, ci in colmap.items():
            out["daily"][label] = _cell(grid, date_row, ci)
    except ValueError:
        pass

    d = report_date.isoformat()
    for ri in range(1, len(hw)):
        if _cell(hw, ri, 0) == d and _cell(hw, ri, 1) == branch:
            comments_raw = hw[ri][6] if len(hw[ri]) > 6 else ""
            out["activity"] = {
                "health_cafe": hw[ri][2] if len(hw[ri]) > 2 else "",
                "health_home": hw[ri][3] if len(hw[ri]) > 3 else "",
                "daily_post": hw[ri][4] if len(hw[ri]) > 4 else "",
                "comments": [c for c in comments_raw.split("\n") if c.strip()],
            }
            break

    has_any = bool(out["activity"]) or any(v.strip() for v in out["daily"].values())
    return out if has_any else None


# ──────────────────────────────────────────────────────────────────
# 쓰기
# ──────────────────────────────────────────────────────────────────
def write_submission(payload: dict) -> dict:
    branch = payload["branch"]
    d = payload["date"]
    daily_name = config.daily_sheet_name(d)

    # 1) 일일 시트: 지점 열그룹 × 날짜 행에 소계 기록
    #    월별 탭이 아직 없거나(월 초 미생성) 구조 문제면 일일탭만 건너뛰고 계속 진행
    #    (카페숙제·초진상세는 기록 → 제출이 실패하지 않음)
    a1 = None
    try:
        grid = _vget(_q(daily_name, "A:AG"))
        hdr_row, colmap = build_column_map(grid, branch)
        date_row = find_date_row(grid, hdr_row, d)

        values_by_label = {}
        for sec in config.INQUIRY_SECTIONS:
            values_by_label[sec["sheet_col"]] = payload["sections"][sec["key"]]["subtotal"]

        cols = sorted(colmap.values())
        c_start, c_end = cols[0], cols[-1]
        inv = {ci: label for label, ci in colmap.items()}
        row_vals = [values_by_label.get(inv.get(c), "") for c in range(c_start, c_end + 1)]
        a1 = _q(daily_name, f"{_col_letter(c_start)}{date_row + 1}:{_col_letter(c_end)}{date_row + 1}")
        _vupdate(a1, [row_vals])
    except Exception:  # noqa: BLE001  (월별 탭 미존재 등)
        a1 = None

    # 2) 카페숙제체크 upsert
    _upsert_homework(payload)

    # 3) 초진상세 (날짜,지점) 교체
    detail_n = _replace_detail(payload)

    return {"ok": True, "daily_range": a1, "detail_rows": detail_n}


def _upsert_homework(payload: dict):
    vals = _vget(_q(config.HOMEWORK_TAB, "A:I"))
    d = payload["date"].isoformat()
    branch = payload["branch"]
    a = payload["activity"]
    count = a.get("comment_count", len(a["comments"]))
    record = [
        d, branch, a["health_cafe"], a["health_home"], a["daily_post"],
        str(count), "\n".join(a["comments"]), payload["writer"], _now_str(),
    ]

    target = None
    for ri in range(1, len(vals)):
        if _cell(vals, ri, 0) == d and _cell(vals, ri, 1) == branch:
            target = ri
            break
    if target is not None:
        _vupdate(_q(config.HOMEWORK_TAB, f"A{target + 1}:I{target + 1}"), [record])
    else:
        _vappend(_q(config.HOMEWORK_TAB, "A:I"), [record])


def _replace_detail(payload: dict) -> int:
    """초진상세에서 (날짜,지점) 기존 행을 빼고 현재 제출 상세로 재기록 (sheetId 불필요)."""
    from googleapiclient.errors import HttpError

    rng = _q(config.DETAIL_TAB, "A:J")
    try:
        vals = _vget(rng)
    except HttpError:
        _create_tab(config.DETAIL_TAB, config.DETAIL_HEADERS)
        vals = [config.DETAIL_HEADERS]
    if not vals:
        vals = [config.DETAIL_HEADERS]

    header = vals[0]
    d = payload["date"].isoformat()
    branch = payload["branch"]
    kept = [
        r for r in vals[1:]
        if not (len(r) > 1 and r[0].strip() == d and r[1].strip() == branch)
    ]

    now = _now_str()
    new_rows = []
    for sec in config.INQUIRY_SECTIONS:
        for it in payload["sections"][sec["key"]]["items"]:
            if not it.get("name"):
                continue
            route = it.get("route", "")
            new_rows.append([
                d, branch, sec["label"],
                config.DISEASE_CATEGORY.get(it["name"], ""), it["name"],
                config.ROUTE_CATEGORY.get(route, ""), route,
                it["count"], payload["writer"], now,
            ])

    final = [header] + kept + new_rows
    _vupdate(_q(config.DETAIL_TAB, "A1"), final)
    if len(final) < len(vals):  # 이전보다 행이 줄면 남은 뒷행 비우기
        _vclear(_q(config.DETAIL_TAB, f"A{len(final) + 1}:J{len(vals)}"))
    return len(new_rows)
