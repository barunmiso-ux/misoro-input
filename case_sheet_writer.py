"""주간시트 환자행 쓰기 — 케이스 추적기 P1 벽돌(b).

(주의) 같은 폴더 sheet_writer.py 는 '일일앱'의 일일시트 I/O 모듈로 별개다. 이 파일은
케이스 추적기가 '지점 주간통계 시트'의 초진 환자테이블에 환자행을 기록하는 전용 모듈.

검증됨(2026-06-30, 분당 사본): export 23열을 B5:Y 에 위치기록하면
시트 기존 수식이 예약율·한약결제율(특화=피부+호흡기)·질환군별을 자동계산.

안전계약(절대 위반 금지):
- **환자테이블(B{FIRST}:Y{LAST})만** 건드린다.
- 상담테이블(Z~AK)·집계/수식(157행↓)·앵커 라벨은 손대지 않는다 → 다운스트림(misoro-dashboard·캐시·우리보고서) 무변경.
- dry_run 기본 True. 실제 쓰기는 호출자가 명시적으로 dry_run=False.

사용:
  from export_parser import rows_for_sheet
  from case_sheet_writer import write_patients
  rows = rows_for_sheet("환자검색결과.xls")
  write_patients(SPREADSHEET_ID, "26-06-3주", rows)               # dry_run: 계획만
  write_patients(SPREADSHEET_ID, "26-06-3주", rows, dry_run=False)  # 실제 기록
"""
from __future__ import annotations

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

PATIENT_FIRST_ROW = 5
PATIENT_LAST_ROW = 154    # 시트 수식 COUNTIF 범위 상한(>=153). 그 아래는 절대 미변경.
EXPECTED_COLS = 23        # export 23열(= 시트 C..Y)

# ②상담테이블(문의): 번호 Z + AA~AK(11열). 수식 COUNTIF 범위 상한 ~139.
INQ_FIRST_ROW = 5
INQ_LAST_ROW = 139
INQ_EXPECTED_COLS = 11    # 상담내역 11열(= 시트 AA..AK)
DEFAULT_KEY = r"D:\Data\클로드 코드\미소로광고분석\service_account.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def _creds(key_path: str = DEFAULT_KEY):
    """인증 우선순위: ① GCP_SA_KEY env(JSON 문자열, GitHub Actions용) →
    ② 로컬 service_account.json/DEFAULT_KEY → ③ Streamlit secrets(배포)."""
    import os
    key = os.environ.get("GCP_SA_KEY")
    if key:
        import json
        return Credentials.from_service_account_info(json.loads(key), scopes=SCOPES)
    for p in ("service_account.json", key_path):
        if p and os.path.exists(p):
            return Credentials.from_service_account_file(p, scopes=SCOPES)
    import streamlit as st  # 배포 환경
    return Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES)


def _svc(key_path: str = DEFAULT_KEY):
    return build("sheets", "v4", credentials=_creds(key_path),
                 cache_discovery=False).spreadsheets()


def _tabs(sh, sid) -> list:
    meta = sh.get(spreadsheetId=sid, fields="sheets(properties(title))").execute()
    return [s["properties"]["title"] for s in meta.get("sheets", [])]


def write_patients(spreadsheet_id: str, tab: str, rows23: list, *,
                   key_path: str = DEFAULT_KEY, dry_run: bool = True,
                   verify: bool = True, merge: bool = False) -> dict:
    """rows23: 23열 행 리스트(export 순서). B{FIRST}:Y{LAST} 클리어 후 환자행 기록.

    merge=False → 탭 통째 교체(기존 환자행 삭제 후 새로 기록).
    merge=True  → **차트번호 기준 upsert**(기존 유지 + 같은 차트는 갱신 + 새 차트는 추가).
                  여러 번/여러 주에 나눠 올려도 데이터 안 지워짐. 차트 없으면 이름으로 키.
    dry_run=True → 계획만 반환(쓰기 안 함). False → 실제 기록 후 verify.
    """
    # ── 입력 검증(새 행)
    if not rows23:
        raise ValueError("기록할 환자행이 없습니다.")
    for i, r in enumerate(rows23):
        if len(r) != EXPECTED_COLS:
            raise ValueError(f"{i}행 열수 오류: {len(r)} (기대 {EXPECTED_COLS})")

    sh = _svc(key_path)
    titles = _tabs(sh, spreadsheet_id)
    if tab not in titles:
        raise ValueError(f"탭 '{tab}' 없음. 가능: {titles[:12]}")

    merged_added = merged_updated = 0
    if merge:
        # 기존 C..Y 읽어 차트(=D, C:Y기준 index1) 키로 upsert. 없으면 이름 키.
        existing = sh.values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab}'!C{PATIENT_FIRST_ROW}:Y{PATIENT_LAST_ROW}"
        ).execute().get("values", [])

        def _keyof(r):
            chart = str(r[1]).strip() if len(r) > 1 else ""
            name = str(r[0]).strip() if len(r) > 0 else ""
            return chart or ("이름:" + name)

        by_key, order = {}, []
        for r in existing:
            r = (list(r) + [""] * EXPECTED_COLS)[:EXPECTED_COLS]
            if not str(r[0]).strip() and not str(r[1]).strip():
                continue
            k = _keyof(r)
            if k not in by_key:
                order.append(k)
            by_key[k] = r
        for r in rows23:
            k = _keyof(r)
            if k in by_key:
                merged_updated += 1
            else:
                order.append(k)
                merged_added += 1
            by_key[k] = list(r)
        rows23 = [by_key[k] for k in order]

    n = len(rows23)
    cap = PATIENT_LAST_ROW - PATIENT_FIRST_ROW + 1
    if n > cap:
        raise ValueError(f"환자 {n}명이 테이블 용량({cap}) 초과 — 양식 확인 필요")

    clear_range = f"'{tab}'!B{PATIENT_FIRST_ROW}:Y{PATIENT_LAST_ROW}"
    write_anchor = f"'{tab}'!B{PATIENT_FIRST_ROW}"
    values = [[str(i + 1)] + list(rows23[i]) for i in range(n)]  # B=번호 + C..Y

    plan = {
        "spreadsheet_id": spreadsheet_id, "tab": tab,
        "clear_range": clear_range, "write_anchor": write_anchor, "rows": n,
        "보존": "Z+ 상담테이블·157행↓ 수식·앵커 미변경",
        "merge": merge, "추가": merged_added, "갱신": merged_updated,
        "dry_run": dry_run,
    }
    if dry_run:
        return plan

    # ── 실제 기록: 환자영역 클리어 → 기록
    sh.values().clear(spreadsheetId=spreadsheet_id, range=clear_range).execute()
    sh.values().update(spreadsheetId=spreadsheet_id, range=write_anchor,
                       valueInputOption="USER_ENTERED", body={"values": values}).execute()
    plan["written"] = True

    if verify:
        got = sh.values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab}'!B{PATIENT_FIRST_ROW}:B{PATIENT_FIRST_ROW + n - 1}"
        ).execute().get("values", [])
        plan["verify_rows"] = len(got)
        plan["verify_ok"] = (len(got) == n)
    return plan


# 결산(매출/내원) 직접입력 셀 — 주간탭·월간탭 동일 템플릿(검증 2026-07-01).
SETTLE_CELLS = {"매출": "D158", "환불": "D159", "총내원": "D166", "신환내원": "D167"}


def write_settlement(spreadsheet_id: str, tab: str, values: dict, *,
                     key_path: str = DEFAULT_KEY, dry_run: bool = True) -> dict:
    """OKTAS 결산값을 탭의 직접입력 셀에 기록. values 키: 매출/환불/총내원/신환내원.

    수식·구조는 안 건드리고 해당 단일 셀만 갱신(D158/D159/D166/D167) → 비율 자동 재계산.
    """
    data = []
    plan_writes = {}
    for k, cell in SETTLE_CELLS.items():
        v = values.get(k)
        if v is not None:
            data.append({"range": f"'{tab}'!{cell}", "values": [[v]]})
            plan_writes[cell] = v
    if not data:
        raise ValueError("기록할 결산값이 없습니다.")

    plan = {"tab": tab, "writes": plan_writes, "dry_run": dry_run}
    if dry_run:
        return plan

    sh = _svc(key_path)
    if tab not in _tabs(sh, spreadsheet_id):
        raise ValueError(f"탭 '{tab}' 없음")
    sh.values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": data}).execute()
    plan["written"] = True
    return plan


def write_inquiries(spreadsheet_id: str, tab: str, rows11: list, *,
                    key_path: str = DEFAULT_KEY, dry_run: bool = True,
                    verify: bool = True, merge: bool = False) -> dict:
    """rows11: 11열(AA~AK) 행. Z{FIRST}:AK{LAST} 클리어 후 번호+상담행 기록.

    merge=True → **차트번호(없으면 성명|상담시각) 기준 upsert**(기존 유지+갱신+추가).
    안전계약: **Z5:AK139만**(①초진테이블 B~Y·수식·앵커 미변경). dry_run 기본 True.
    """
    if not rows11:
        raise ValueError("기록할 문의행이 없습니다.")
    for i, r in enumerate(rows11):
        if len(r) != INQ_EXPECTED_COLS:
            raise ValueError(f"{i}행 열수 오류: {len(r)} (기대 {INQ_EXPECTED_COLS})")

    sh = _svc(key_path)
    titles = _tabs(sh, spreadsheet_id)
    if tab not in titles:
        raise ValueError(f"탭 '{tab}' 없음. 가능: {titles[:12]}")

    merged_added = merged_updated = 0
    if merge:
        # 기존 AA..AK 읽어 upsert. AA:AK 기준 index 0=상담시각·1=차트·2=성명.
        existing = sh.values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab}'!AA{INQ_FIRST_ROW}:AK{INQ_LAST_ROW}"
        ).execute().get("values", [])

        def _keyof(r):
            chart = str(r[1]).strip() if len(r) > 1 else ""
            if chart:
                return chart
            name = str(r[2]).strip() if len(r) > 2 else ""
            tm = str(r[0]).strip() if len(r) > 0 else ""
            return f"{name}|{tm}"

        by_key, order = {}, []
        for r in existing:
            r = (list(r) + [""] * INQ_EXPECTED_COLS)[:INQ_EXPECTED_COLS]
            if not any(str(x).strip() for x in r):
                continue
            k = _keyof(r)
            if k not in by_key:
                order.append(k)
            by_key[k] = r
        for r in rows11:
            k = _keyof(r)
            if k in by_key:
                merged_updated += 1
            else:
                order.append(k)
                merged_added += 1
            by_key[k] = list(r)
        rows11 = [by_key[k] for k in order]

    n = len(rows11)
    cap = INQ_LAST_ROW - INQ_FIRST_ROW + 1
    if n > cap:
        raise ValueError(f"문의 {n}건이 테이블 용량({cap}) 초과 — 양식 확인 필요")

    clear_range = f"'{tab}'!Z{INQ_FIRST_ROW}:AK{INQ_LAST_ROW}"
    write_anchor = f"'{tab}'!Z{INQ_FIRST_ROW}"
    values = [[str(i + 1)] + list(rows11[i]) for i in range(n)]  # Z=번호 + AA..AK

    plan = {
        "spreadsheet_id": spreadsheet_id, "tab": tab,
        "clear_range": clear_range, "write_anchor": write_anchor, "rows": n,
        "보존": "B~Y 초진테이블·157행↓ 수식·앵커 미변경",
        "merge": merge, "추가": merged_added, "갱신": merged_updated,
        "dry_run": dry_run,
    }
    if dry_run:
        return plan

    sh.values().clear(spreadsheetId=spreadsheet_id, range=clear_range).execute()
    sh.values().update(spreadsheetId=spreadsheet_id, range=write_anchor,
                       valueInputOption="USER_ENTERED", body={"values": values}).execute()
    plan["written"] = True
    if verify:
        got = sh.values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab}'!Z{INQ_FIRST_ROW}:Z{INQ_FIRST_ROW + n - 1}"
        ).execute().get("values", [])
        plan["verify_rows"] = len(got)
        plan["verify_ok"] = (len(got) == n)
    return plan


def _week_num(tab: str) -> int:
    import re
    m = re.search(r"-(\d+)주", tab)
    return int(m.group(1)) if m else 99


def aggregate_month(sid: str, month_prefix: str, all_tabs: list, *,
                    key_path: str = DEFAULT_KEY, dry_run: bool = True) -> dict:
    """그 달 주간탭들의 초진/문의 환자행을 합쳐 월간탭에 기록.

    month_prefix='26-06' → 주간 '26-06-N주' 합쳐 → '26-06월' 에 기록.
    매출·총내원(157행↓ 직접입력)은 안 건드림(OKTAS 월말결산 몫). 비율류는 월간탭 수식이 재계산.
    """
    month_tab = f"{month_prefix}월"
    weeks = sorted([t for t in all_tabs if t.startswith(month_prefix + "-") and "주" in t],
                   key=_week_num)
    if month_tab not in all_tabs:
        raise ValueError(f"월간탭 '{month_tab}' 없음")
    if not weeks:
        raise ValueError(f"'{month_prefix}' 주간탭이 없습니다")

    sh = _svc(key_path)
    ranges = []
    for w in weeks:
        ranges += [f"'{w}'!B{PATIENT_FIRST_ROW}:Y{PATIENT_LAST_ROW}",
                   f"'{w}'!Z{INQ_FIRST_ROW}:AK{INQ_LAST_ROW}"]
    vrs = sh.values().batchGet(spreadsheetId=sid, ranges=ranges).execute().get("valueRanges", [])

    chojin, inq = [], []
    for k in range(len(weeks)):
        for row in vrs[2 * k].get("values", []):
            r = (list(row) + [""] * 24)[:24]
            cy = [str(x).strip() for x in r[1:24]]          # C..Y (23열)
            if any(cy):
                chojin.append(cy)
        for row in vrs[2 * k + 1].get("values", []):
            r = (list(row) + [""] * 12)[:12]
            aa = [str(x).strip() for x in r[1:12]]          # AA..AK (11열)
            if any(aa):
                inq.append(aa)

    plan = {"month_tab": month_tab, "weeks": weeks,
            "초진합계": len(chojin), "문의합계": len(inq), "dry_run": dry_run}
    if dry_run:
        return plan

    if chojin:
        write_patients(sid, month_tab, chojin, key_path=key_path, dry_run=False)
    if inq:
        write_inquiries(sid, month_tab, inq, key_path=key_path, dry_run=False)
    plan["written"] = True
    return plan


if __name__ == "__main__":
    import sys
    from export_parser import rows_for_sheet
    if len(sys.argv) < 4:
        print('사용: python case_sheet_writer.py <spreadsheet_id> <tab> <export.xls> [--commit]')
        sys.exit()
    sid, tab, path = sys.argv[1], sys.argv[2], sys.argv[3]
    commit = "--commit" in sys.argv
    rows = rows_for_sheet(path)
    res = write_patients(sid, tab, rows, dry_run=not commit)
    print("[실제기록]" if commit else "[DRY-RUN 계획만]")
    for k, v in res.items():
        print(f"  {k}: {v}")
