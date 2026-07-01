# -*- coding: utf-8 -*-
"""다음 달(또는 인자로 준 YYYY-MM) '[YY.MM] 지점별 주간 문의내역' 탭을 자동 생성.

숨김 템플릿 `_TEMPLATE_문의내역`(빈 31일 구조)을 복제한 뒤 해당 월의 날짜/라벨을 채운다.
GitHub Actions(월 25일)에서 실행하거나 로컬에서 인자로 테스트한다.

인증: 환경변수 GCP_SA_KEY(서비스계정 JSON 문자열) 우선, 없으면 로컬 service_account.json.
"""
import os
import sys
import json
import calendar
from datetime import date

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SHEET_ID = "1uTkikVDCUfVry6l-GX2L8yvlhRSVFdUtEwlubrhpZLk"
TEMPLATE = "_TEMPLATE_문의내역"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 31일 템플릿 기준 블록 위치
B1_START, B1_END = 21, 51   # block1 날짜행
B2_START, B2_END = 57, 87   # block2 날짜행


def _creds():
    key = os.environ.get("GCP_SA_KEY")
    if key:
        return Credentials.from_service_account_info(json.loads(key), scopes=SCOPES)
    return Credentials.from_service_account_file("service_account.json", scopes=SCOPES)


def main():
    if len(sys.argv) > 1:               # 테스트: YYYY-MM
        y, m = map(int, sys.argv[1].split("-"))
    else:                               # 기본: 다음 달
        t = date.today()
        y, m = (t.year + 1, 1) if t.month == 12 else (t.year, t.month + 1)

    title = f"[{y % 100:02d}.{m:02d}] 지점별 주간 문의내역"
    svc = build("sheets", "v4", credentials=_creds(), cache_discovery=False)

    meta = svc.spreadsheets().get(spreadsheetId=SHEET_ID,
        fields="sheets(properties(sheetId,title))").execute(num_retries=3)
    titles = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}

    if title in titles:
        print(f"이미 존재: {title} — 건너뜀")
        return
    if TEMPLATE not in titles:
        print(f"템플릿 '{TEMPLATE}' 없음 — 중단")
        sys.exit(1)

    # 복제
    r = svc.spreadsheets().batchUpdate(spreadsheetId=SHEET_ID, body={"requests": [
        {"duplicateSheet": {"sourceSheetId": titles[TEMPLATE], "newSheetName": title}}]}).execute(num_retries=3)
    newid = r["replies"][0]["duplicateSheet"]["properties"]["sheetId"]
    # 템플릿은 숨김 → 복제본은 보이게
    svc.spreadsheets().batchUpdate(spreadsheetId=SHEET_ID, body={"requests": [
        {"updateSheetProperties": {"properties": {"sheetId": newid, "hidden": False},
         "fields": "hidden"}}]}).execute(num_retries=3)

    days = calendar.monthrange(y, m)[1]
    dates = [[f"{y}-{m:02d}-{d:02d}"] for d in range(1, days + 1)]

    def upd(rng, vals):
        svc.spreadsheets().values().update(spreadsheetId=SHEET_ID, range=f"'{title}'!{rng}",
            valueInputOption="USER_ENTERED", body={"values": vals}).execute(num_retries=3)

    def clr(rng):
        svc.spreadsheets().values().clear(spreadsheetId=SHEET_ID,
            range=f"'{title}'!{rng}").execute(num_retries=3)

    upd(f"B{B1_START}:B{B1_START + days - 1}", dates)
    if days < 31:
        clr(f"B{B1_START + days}:B{B1_END}")
    upd(f"B{B2_START}:B{B2_START + days - 1}", dates)
    if days < 31:
        clr(f"B{B2_START + days}:B{B2_END}")

    upd("B3", [[f"{y % 100:02d}.{m:02d}월 초진/문의 집계"]])
    upd("B8", [[f"{y % 100:02d}.{m:02d}월 초진/문의 내역별 집계"]])
    upd("B52", [[f"{m}월 합계"]])
    upd("B88", [[f"{m}월 합계"]])

    print(f"생성 완료: {title} ({days}일)")


if __name__ == "__main__":
    main()
