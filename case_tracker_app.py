"""케이스 추적기 — 주간통계 초진·문의 기록 앱 (P1 골격).

흐름:  ?branch= 지점확인 → [초진 탭] 환자검색결과.xls / [문의 탭] 상담내역.xls 업로드
       → 파서 미리보기·검증 → 주차 자동감지 → 기록(시트 수식 자동계산).

로컬 실행:  streamlit run case_tracker_app.py
배포:       misoro-input 과 동일(Streamlit Cloud, ?branch= URL, st.secrets[gcp_service_account]).
"""
from __future__ import annotations

import hmac
import tempfile

import streamlit as st

from datetime import datetime, timezone, timedelta

from export_parser import (parse_export, rows_for_sheet, infer_week_tab,
                           parse_inquiries, inquiry_rows_for_sheet, inquiry_week_tab,
                           week_label, rows_for_sheet_by_week, inquiry_rows_by_week,
                           to_standard_treatment, normalize_result)
from case_sheet_writer import write_patients, write_inquiries, aggregate_month, _svc
from upload_log import log_upload
from noshow_matcher import match_inquiries, set_override

_KST = timezone(timedelta(hours=9))

# 지점 → 주간통계 스프레드시트 ID (2026-06-30 Drive 확인분; 미확인 지점은 추후 추가)
BRANCH_SHEETS = {
    "분당(테스트사본)": "1ocYot2i8NsM-pmV2kHCOYCngIEsx5SccgkpfZ0HU9M8",  # 시험용 — 라이브 아님
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
PILOT_DEFAULT = "분당(테스트사본)"  # 시험 단계 기본값(라이브 아님). 운영 전환 시 "분당"으로.


@st.cache_data(ttl=300, show_spinner=False)
def _list_week_tabs(sid: str) -> list:
    sh = _svc()
    meta = sh.get(spreadsheetId=sid, fields="sheets(properties(title))").execute()
    titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
    return [t for t in titles if "주" in t]  # 26-06-3주 류


def _branch_from_url() -> str:
    try:
        b = st.query_params.get("branch", "")
    except Exception:
        b = st.experimental_get_query_params().get("branch", [""])[0]
    return (b or "").strip()


def _save_temp(up) -> str:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tf:
        tf.write(up.getbuffer())
        return tf.name


def _week_selectbox(tabs: list, inferred, winfo: dict, key: str) -> str:
    """주차 자동감지 안내 + 선택 + 불일치 경고. 반환 선택된 탭."""
    default_idx = tabs.index(inferred) if inferred in tabs else 0
    if inferred and inferred in tabs:
        st.success(f"📅 초진/상담 날짜 기준 **자동감지: {week_label(inferred)}** (직접 바꿀 수 있어요)")
    elif inferred:
        st.warning(f"자동감지 주차 '{week_label(inferred)}' 탭이 시트에 없습니다 — 탭을 먼저 만들어야 합니다.")
    if winfo.get("multi"):
        st.warning(f"⚠️ 날짜가 여러 주에 걸쳐 있습니다: {winfo['counts']} — 주차를 꼭 확인하세요.")
    tab = st.selectbox("기록할 주차 탭", tabs, index=default_idx, key=key,
                       format_func=week_label)   # '26-06-4주 (6/22~6/28)' 로 표시
    if inferred and tab != inferred:
        st.warning(f"⚠️ 자동감지({week_label(inferred)})와 다른 주차({week_label(tab)})를 선택했습니다. 맞는지 확인하세요.")
    return tab


# ──────────────────────────────────────────────────────────────────
# 초진 탭
# ──────────────────────────────────────────────────────────────────
def render_chojin(sid: str, tabs: list, branch: str = ""):
    up = st.file_uploader("초진 export — 환자검색결과 (.xls/.xlsx)", type=["xls", "xlsx"], key="up_chojin")
    if not up:
        st.caption("차트에서 '환자검색결과'를 .xls로 내려받아 올리세요.")
        return
    try:
        path = _save_temp(up)
        parsed = parse_export(path)
        rows = rows_for_sheet(path)
    except Exception as e:
        st.error(f"파일을 읽지 못했습니다: {e}")
        return

    s = parsed["summary"]
    st.subheader("미리보기")
    tr = s["특화전환율"]
    rate = f"{tr * 100:.0f}%" if tr is not None else "-"
    seen = s["특화초진"] - s["특화진료전이탈"]
    c1, c2, c3 = st.columns(3)
    c1.metric("초진수", s["초진수"])
    c2.metric("특화 결제/진료", f"{s['특화결제']} / {seen}")
    c3.metric("특화 전환율", rate)
    bd = s["진료결과별"]
    st.write("**진료결과:** "
             f"한약결제 {bd.get('한약결제', 0)} · 약침결제 {bd.get('약침결제', 0)} · "
             f"첩약보험 {bd.get('첩약보험', 0)} · 특화치료 {bd.get('특화치료', 0)} · "
             f"일반치료 {bd.get('일반치료', 0)} · 상담만 {bd.get('상담만', 0)} · "
             f"**그냥감 {bd.get('그냥감', 0)}** · 미정 {bd.get('미정', 0)}"
             + (f" · ⚠️기타 {bd.get('기타', 0)}" if bd.get("기타") else ""))
    st.caption("**전환율 = (한약+약침+첩약보험) / 특화 진료(피부·호흡기, 그냥감 제외)**. "
               f"그냥감(진료 안 봄) **{s['특화진료전이탈']}명은 분모서 제외**(진료 전 이탈 = 접촉이지 결제판단 아님). "
               "특화치료(1회)·일반치료·상담만은 진료봤으나 미결제.")
    st.write("**질환군별:**", s["질환군별"])

    if parsed["missing_headers"]:
        st.warning(f"못 찾은 헤더: {parsed['missing_headers']}")
    if parsed["unmapped_diseases"]:
        st.warning(f"분류표에 없는 질환명(기타 처리): {parsed['unmapped_diseases']}")
    comp = parsed["completeness"]
    _FIELD = {"유입경로": "유입경로 칸", "진행치료": "진행치료 칸",
              "결제여부": "EMail 칸(결제@ / 결제안함@)", "상담자": "직업 칸"}
    # ── 기록 차단 판정: 미완성(특화 필수칸 빔) · 엉뚱(진행치료 자동변환 불가) ──
    미완성 = comp["미완성목록"]
    엉뚱, 자동교정 = [], []
    for p in parsed["patients"]:
        raw = (p.treatment_raw or "").strip()
        if not raw:
            continue
        std, conv = to_standard_treatment(raw)
        if std is None:
            엉뚱.append((p.chart_no, (p.name[:1] + "*") if p.name else "", raw))
        elif conv:
            자동교정.append((p.chart_no, raw, std))

    # ── 구조 이상: 질환이 대량 공란 = export 컬럼이 표준과 안 맞음(천안 케이스) → 전체 차단
    pts = parsed["patients"]
    질환공란 = [p for p in pts if not (p.disease or "").strip()]
    구조이상 = len(pts) >= 2 and len(질환공란) >= len(pts) * 0.5

    blocked = bool(미완성) or bool(엉뚱) or 구조이상

    if 구조이상:
        st.error(f"🚫 **export 구조가 표준과 안 맞습니다 — 초진 {len(pts)}명 중 "
                 f"{len(질환공란)}명의 질환을 못 읽었어요.** 차트 export 형식이 다른 것 같습니다 "
                 "(질환 칸이 비정상 위치). **이대로는 기록할 수 없어요** — 관리자에게 문의하세요.")
    if 미완성:
        st.error(f"🚫 **특화질환 초진 {len(미완성)}명 — 필수칸 비어있음. 차트에서 채워야 기록됩니다.**")
        for chart, nm, dis, miss in 미완성:
            spots = " · ".join(f"**{m}**({_FIELD.get(m, m)})" for m in miss)
            st.write(f"- 차트 **{chart}** {nm} ({dis}) → 채울 곳: {spots}")
    if 엉뚱:
        st.error(f"🚫 **진행치료 값 {len(엉뚱)}건이 표준값 아님 + 자동변환 불가 — 차트에서 고쳐야 기록됩니다.**")
        for chart, nm, raw in 엉뚱:
            st.write(f"- 차트 **{chart}** {nm}: `{raw}` → 표준값"
                     "(한약N달·약침패키지·첩약보험·특화치료·일반치료·상담만·그냥감) 중 하나로")
    if 자동교정:
        with st.expander(f"✏️ 표준값으로 자동 변환되어 기록될 진행치료 {len(자동교정)}건 (조치 불필요)"):
            for chart, raw, std in 자동교정:
                st.write(f"- [{chart}] `{raw}` → **`{std}`** 로 기록")
    if not blocked:
        st.success("✅ 완전성·표준값 OK — 기록 가능")

    with st.expander(f"환자 {len(parsed['patients'])}명 (PII 마스킹)"):
        st.table([
            {"차트": p.chart_no, "이름": (p.name[:1] + "*") if p.name else "",
             "질환": f"{p.disease}({p.disease_group})", "결제여부": p.booking,
             "진료결과": p.outcome, "유입": p.inflow, "상담": p.counselor}
            for p in parsed["patients"]
        ])

    st.subheader("초진 시트에 기록")
    st.caption("환자별 **등록일 기준으로 각 주차 탭에 자동 분리 + 병합(차트번호 기준)** 기록합니다. "
               "한 파일에 여러 주·여러 달이 섞여 있어도 알아서 나눠 넣어요. 상담테이블·수식·집계는 보존.")
    by_week = rows_for_sheet_by_week(path)
    unknown = by_week.pop("_미분류", None)
    known = {wk: r for wk, r in by_week.items() if wk in tabs}
    absent = {wk: r for wk, r in by_week.items() if wk not in tabs}
    for wk in sorted(known):
        st.write(f"- **{week_label(wk)}** → 초진 **{len(known[wk])}명** 기록")
    for wk in sorted(absent):
        st.warning(f"⚠️ '{week_label(wk)}' 초진 {len(absent[wk])}명 — 시트에 그 주차 탭이 없어 못 넣어요 (탭 먼저 생성 필요)")
    if unknown:
        st.warning(f"⚠️ 등록일을 못 읽은 {len(unknown)}명은 제외됩니다.")
    if not known:
        st.error("기록할 주차가 없어요 (해당 주차 탭이 시트에 없음).")
        return
    total = sum(len(v) for v in known.values())
    if blocked:
        st.warning("🚫 위 🚫 항목(미완성·표준값 아님)을 차트에서 고치고 다시 올려야 기록할 수 있어요.")
    confirm = st.checkbox(f"위 {len(known)}개 주차에 초진 {total}명 병합 기록", key="cf_chojin",
                          disabled=blocked)
    if st.button("📝 초진 기록하기", type="primary",
                 disabled=(not confirm or blocked), key="bt_chojin"):
        oks = 0
        for wk in sorted(known):
            try:
                res = write_patients(sid, wk, known[wk], dry_run=False, merge=True)
                st.success(f"✅ {week_label(wk)} — 추가 {res['추가']} · 갱신 {res['갱신']} (탭 총 {res['rows']}명)")
                log_upload(branch, "초진", wk, res["추가"], res["갱신"], res["rows"])
                oks += 1
            except Exception as e:
                st.error(f"❌ {week_label(wk)} 기록 실패: {e}")
        if oks == len(known):
            st.balloons()


# ──────────────────────────────────────────────────────────────────
# 문의 탭
# ──────────────────────────────────────────────────────────────────
def render_munui(sid: str, tabs: list, branch: str = ""):
    up = st.file_uploader("문의 export — 상담내역 (.xls/.xlsx)", type=["xls", "xlsx"], key="up_munui")
    if not up:
        st.caption("차트에서 '상담내역'을 .xls로 내려받아 올리세요.")
        return
    try:
        path = _save_temp(up)
        parsed = parse_inquiries(path)
        rows = inquiry_rows_for_sheet(path)
    except Exception as e:
        st.error(f"파일을 읽지 못했습니다: {e}")
        return

    s = parsed["summary"]
    st.subheader("미리보기")
    rate = f"{s['예약전환율'] * 100:.0f}%"
    c1, c2, c3 = st.columns(3)
    c1.metric("문의수", s["문의수"])
    c2.metric("예약완료 / 예약안함", f"{s['예약완료']} / {s['예약안함']}")
    c3.metric("예약전환율", rate)
    st.caption(f"예약전환율 = 예약완료 / 문의수. 진행중(재통화필요 등) {s['진행중']}건은 미전환으로 집계.")
    st.write("**질환군별:**", s["질환군별"])
    st.write("**상담구분별:**", s["상담구분별"])

    if parsed["unmapped_diseases"]:
        st.warning(f"분류표에 없는 질환명(기타 처리): {parsed['unmapped_diseases']}")
    _STD_R = ("예약완료", "예약안함", "재통화필요")
    nonstd_r = s.get("비표준결과", [])
    자동_r = [(nm, r, normalize_result(r)) for nm, r in nonstd_r if normalize_result(r) in _STD_R]
    엉뚱_r = [(nm, r) for nm, r in nonstd_r if normalize_result(r) not in _STD_R]
    blocked_m = bool(엉뚱_r)
    if 엉뚱_r:
        st.error(f"🚫 **상담결과 {len(엉뚱_r)}건이 표준값 아님 + 자동변환 불가 — 차트에서 고쳐야 기록됩니다.** "
                 + ", ".join(f"{nm} `{r}`" for nm, r in 엉뚱_r[:8]))
    if 자동_r:
        with st.expander(f"✏️ 표준값으로 자동 변환되어 기록될 상담결과 {len(자동_r)}건 (조치 불필요)"):
            for nm, r, std in 자동_r:
                st.write(f"- {nm}: `{r}` → **`{std}`**")

    with st.expander(f"문의 {len(parsed['inquiries'])}건 (PII 마스킹)"):
        st.table([
            {"시각": i.consult_time[5:16], "성명": (i.name[:1] + "*") if i.name else "",
             "질환": f"{i.disease}({i.disease_group})", "상담구분": i.channel,
             "결과": i.result, "상담자": i.counselor}
            for i in parsed["inquiries"]
        ])

    st.subheader("문의 시트에 기록")
    st.caption("**상담시각 기준으로 각 주차 탭에 자동 분리 + 병합(차트/성명 기준)** 기록합니다. "
               "여러 주·여러 달 섞여도 알아서 나눠 넣어요. 초진테이블·수식·집계는 보존.")
    by_week = inquiry_rows_by_week(path)
    unknown = by_week.pop("_미분류", None)
    known = {wk: r for wk, r in by_week.items() if wk in tabs}
    absent = {wk: r for wk, r in by_week.items() if wk not in tabs}
    for wk in sorted(known):
        st.write(f"- **{week_label(wk)}** → 문의 **{len(known[wk])}건** 기록")
    for wk in sorted(absent):
        st.warning(f"⚠️ '{week_label(wk)}' 문의 {len(absent[wk])}건 — 시트에 그 주차 탭이 없어 못 넣어요 (탭 먼저 생성 필요)")
    if unknown:
        st.warning(f"⚠️ 상담시각을 못 읽은 {len(unknown)}건은 제외됩니다.")
    if not known:
        st.error("기록할 주차가 없어요 (해당 주차 탭이 시트에 없음).")
        return
    total = sum(len(v) for v in known.values())
    if blocked_m:
        st.warning("🚫 위 🚫 상담결과(표준값 아님)를 차트에서 고치고 다시 올려야 기록할 수 있어요.")
    confirm = st.checkbox(f"위 {len(known)}개 주차에 문의 {total}건 병합 기록", key="cf_munui",
                          disabled=blocked_m)
    if st.button("📝 문의 기록하기", type="primary",
                 disabled=(not confirm or blocked_m), key="bt_munui"):
        oks = 0
        for wk in sorted(known):
            try:
                res = write_inquiries(sid, wk, known[wk], dry_run=False, merge=True)
                st.success(f"✅ {week_label(wk)} — 추가 {res['추가']} · 갱신 {res['갱신']} (탭 총 {res['rows']}건)")
                log_upload(branch, "문의", wk, res["추가"], res["갱신"], res["rows"])
                oks += 1
            except Exception as e:
                st.error(f"❌ {week_label(wk)} 기록 실패: {e}")
        if oks == len(known):
            st.balloons()


@st.cache_data(ttl=300, show_spinner="문의↔초진 매칭 중…")
def _run_match(sid: str, asof: str, tabs_tuple: tuple) -> dict:
    return match_inquiries(sid, asof, list(tabs_tuple))


def _mask(nm: str) -> str:
    return (nm[:1] + "*") if nm else "무명"


def render_noshow(sid: str, tabs: list):
    asof = datetime.now(_KST).date().isoformat()
    r = _run_match(sid, asof, tuple(tabs))
    c = r["counts"]
    st.caption(f"기준일 **{r['asof']}** · 윈도우 {r['window_days']}일 · 예약완료 {r['예약완료수']}건 "
               "(문의일+2주 안에 초진 나타났는지 자동 대조)")
    m = st.columns(4)
    m[0].metric("✅ 전환", c["전환"])
    m[1].metric("⏳ 내원대기", c["내원대기"])
    m[2].metric("❌ 노쇼", c["노쇼"])
    rate = r["확정전환율"]
    m[3].metric("확정전환율", f"{rate * 100:.0f}%" if rate is not None else "-")

    def _tbl(items):
        return [{"주차": x["week"], "성명": _mask(x["name"]), "질환": x["disease"],
                 "문의일": str(x["time"]), "비고": x["note"]} for x in items]

    waiting = [x for x in r["rows"] if x["status"] == "내원대기"]
    st.subheader(f"⏳ 내원대기 {len(waiting)}건 — 리마인드 콜 대상")
    st.caption("예약완료했지만 아직 안 옴. 마감(D-day) 전이라 노쇼 아님 — 전화해서 챙기면 전환 가능.")
    if waiting:
        st.table(_tbl(sorted(waiting, key=lambda v: v["time"] or "")))
    else:
        st.info("내원대기 없음")

    noshow = [x for x in r["rows"] if x["status"] == "노쇼"]
    with st.expander(f"❌ 노쇼 확정 {len(noshow)}건 (마감 지나도 미내원)"):
        st.table(_tbl(noshow) or [{"-": "없음"}])

    dw = [x for x in r["rows"] if x["status"] == "데이터대기"]
    if dw:
        st.caption(f"🕒 데이터대기 {len(dw)}건 — 해당 주 초진이 업로드되면 자동 재판정 (지금은 노쇼로 단정 안 함)")

    # ── 수기보정 에디터
    st.divider()
    st.subheader("✏️ 상태 수기보정")
    st.caption("매칭이 틀린 경우(무명·전화 다름·타지점 등) '보정' 열을 고치고 저장하세요. "
               "(자동)=매처 판정 그대로. 보정은 자동판정 위에 덮어씁니다.")
    st.caption("💡 한 달 뒤처럼 먼 날짜로 예약한 경우 → '보정'은 그대로 두고 **예약일**칸에 "
               "예약일(예: 8/15 또는 2026-08-15)만 적으면, 그 날까지 내원대기 → 오면 전환 / 지나면 노쇼로 자동판정됩니다.")
    rows = r["rows"]
    keys = [x["key"] for x in rows]
    # 보정 열: 예약일(until)만 있고 강제상태가 없으면 (자동) 유지 — until 은 자동판정에 반영됨
    init = ["(자동)" if (not x["overridden"] or not x.get("forced")) else x["status"] for x in rows]
    init_until = [x.get("until", "") for x in rows]
    disp = [{"주차": x["week"], "성명": _mask(x["name"]), "질환": x["disease"],
             "문의일": str(x["time"]), "자동판정": x["auto_status"],
             "보정": init[i], "예약일": init_until[i]}
            for i, x in enumerate(rows)]
    edited = st.data_editor(
        disp, hide_index=True, use_container_width=True, key="noshow_editor",
        column_config={
            "보정": st.column_config.SelectboxColumn(
                "보정", options=["(자동)", "전환", "노쇼", "내원대기"], required=True),
            "예약일": st.column_config.TextColumn(
                "예약일", help="먼 예약일(예: 8/15). 이 날까지 내원대기 → 오면 전환/지나면 노쇼")},
        disabled=["주차", "성명", "질환", "문의일", "자동판정"])
    if st.button("💾 보정 저장", key="bt_override"):
        now = r["asof"]
        changed = 0
        for i, row in enumerate(edited):
            new_until = (row.get("예약일") or "").strip()
            if row["보정"] != init[i] or new_until != init_until[i]:
                set_override(sid, keys[i], "" if row["보정"] == "(자동)" else row["보정"],
                             "app", now=now, until=new_until)
                changed += 1
        _run_match.clear()  # 매칭 캐시 무효화 → 재계산
        st.success(f"{changed}건 보정 저장됨. 위 카드가 갱신됩니다." if changed else "변경 없음")
        st.rerun()


@st.cache_data(ttl=300, show_spinner=False)
def _all_tabs(sid: str) -> list:
    sh = _svc()
    meta = sh.get(spreadsheetId=sid, fields="sheets(properties(title))").execute()
    return [s["properties"]["title"] for s in meta.get("sheets", [])]


def render_monthly(sid: str):
    st.caption("그 달 주간탭들을 합쳐 **월간탭**을 자동 생성합니다 (대시보드 월간뷰가 읽는 탭). "
               "지점은 월간을 따로 입력할 필요 없음.")
    allt = _all_tabs(sid)
    months = sorted([t for t in allt if t.endswith("월") and t[:2].isdigit() and "-" in t],
                    reverse=True)
    if not months:
        st.info("월간 탭이 없습니다.")
        return
    month_tab = st.selectbox("월간탭 선택", months)
    prefix = month_tab[:-1]  # '26-06월' → '26-06'
    try:
        plan = aggregate_month(sid, prefix, allt, dry_run=True)
    except Exception as e:
        st.error(str(e))
        return

    st.write(f"**{month_tab}** ← 주간 {len(plan['weeks'])}개 합침: {', '.join(plan['weeks'])}")
    st.info(f"초진 **{plan['초진합계']}명** + 문의 **{plan['문의합계']}건** 을 월간탭에 기록")
    st.caption("매출·총내원(직접입력)은 안 건드림 — 차트 월말결산 몫. 비율은 월간탭 수식이 월 단위로 재계산.")
    confirm = st.checkbox(f"{month_tab} 자동생성 확인 (기존 월간탭 환자/문의 교체)", key="cf_month")
    if st.button("🗓️ 월간탭 자동생성", type="primary", disabled=not confirm, key="bt_month"):
        try:
            res = aggregate_month(sid, prefix, allt, dry_run=False)
            st.success(f"✅ {month_tab} 생성: 초진 {res['초진합계']} + 문의 {res['문의합계']}")
            st.balloons()
        except Exception as e:
            st.error(f"실패: {e}")


def _check_password() -> bool:
    """비밀번호 게이트(공개 배포 보호). secrets에 app_password 없으면 통과(로컬 개발).
    통과 시 세션 유지. 대시보드와 동일 방식."""
    try:
        expected = st.secrets.get("app_password")
    except Exception:
        expected = None
    if not expected:
        return True  # 로컬 개발 모드
    if st.session_state.get("pw_ok", False):
        return True
    st.title("📊 주간통계 자동입력")
    with st.form("login"):
        pw = st.text_input("비밀번호", type="password", placeholder="담당자에게 받은 비밀번호")
        if st.form_submit_button("입장", use_container_width=True):
            if hmac.compare_digest(pw or "", str(expected)):
                st.session_state["pw_ok"] = True
                st.rerun()
            else:
                st.error("❌ 비밀번호가 일치하지 않습니다")
    return False


def main():
    st.set_page_config(page_title="주간통계 자동입력", page_icon="📊", layout="centered")
    if not _check_password():
        st.stop()
    st.title("📊 주간통계 자동입력")
    st.caption("차트 명단(초진·문의)을 올리면 주간통계가 자동으로 채워집니다 "
               "— 손으로 입력할 필요 없어요. (예약율·결제율·전환율 자동계산)")

    branch = _branch_from_url()
    if not branch:
        branch = st.selectbox("지점 선택", list(BRANCH_SHEETS.keys()),
                              index=list(BRANCH_SHEETS.keys()).index(PILOT_DEFAULT))
    if branch not in BRANCH_SHEETS:
        st.error(f"'{branch}' 지점 시트가 아직 등록되지 않았습니다. 관리자에게 문의하세요.")
        st.stop()
    sid = BRANCH_SHEETS[branch]
    st.info(f"지점: **{branch}**")

    tabs = _list_week_tabs(sid)
    if not tabs:
        st.error("주차 탭을 찾지 못했습니다.")
        st.stop()

    t1, t2, t3, t4 = st.tabs(["🧑‍⚕️ 초진 기록", "📞 문의 기록", "📊 노쇼/전환", "🗓️ 월간 집계"])
    with t1:
        render_chojin(sid, tabs, branch)
    with t2:
        render_munui(sid, tabs, branch)
    with t3:
        render_noshow(sid, tabs)
    with t4:
        render_monthly(sid)


if __name__ == "__main__":
    main()
