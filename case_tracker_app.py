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
                           parse_inquiries, inquiry_rows_for_sheet, inquiry_week_tab)
from case_sheet_writer import write_patients, write_inquiries, aggregate_month, _svc
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
        st.success(f"📅 초진/상담 날짜 기준 **자동감지: {inferred}** (직접 바꿀 수 있어요)")
    elif inferred:
        st.warning(f"자동감지 주차 '{inferred}' 탭이 시트에 없습니다 — 탭을 먼저 만들어야 합니다.")
    if winfo.get("multi"):
        st.warning(f"⚠️ 날짜가 여러 주에 걸쳐 있습니다: {winfo['counts']} — 주차를 꼭 확인하세요.")
    tab = st.selectbox("기록할 주차 탭", tabs, index=default_idx, key=key)
    if inferred and tab != inferred:
        st.warning(f"⚠️ 자동감지({inferred})와 다른 주차({tab})를 선택했습니다. 맞는지 확인하세요.")
    return tab


# ──────────────────────────────────────────────────────────────────
# 초진 탭
# ──────────────────────────────────────────────────────────────────
def render_chojin(sid: str, tabs: list):
    up = st.file_uploader("초진 export — 환자검색결과 (.xls/.xlsx)", type=["xls", "xlsx"], key="up_chojin")
    if not up:
        st.caption("OKTAS '환자검색결과'를 .xls로 export 해서 올리세요.")
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
    # 각 필수항목이 OKTAS 차트의 어느 칸인지
    _FIELD = {"유입경로": "유입경로 칸", "진행치료": "진행치료 칸",
              "예약여부": "EMail 칸(예약@ / 예약안함@)", "상담자": "직업 칸"}
    if comp["미완성_환자수"]:
        st.error(f"⚠️ **특화질환 초진 {comp['미완성_환자수']}명 — 아래 칸을 OKTAS 차트에서 채우고 다시 올려주세요**")
        for chart, nm, dis, miss in comp["미완성목록"]:
            spots = " · ".join(f"**{m}**({_FIELD.get(m, m)})" for m in miss)
            st.write(f"- 차트 **{chart}** {nm} ({dis}) → 채울 곳: {spots}")
    else:
        st.success("완전성 OK — 특화질환 필수항목 다 채워짐")

    nonstd = comp.get("비표준목록", [])
    if nonstd:
        st.warning(f"✏️ **비표준 진행치료값 {len(nonstd)}건 — 표준값으로 교정 필요** "
                   "(집계는 되지만 차트를 표준값으로 고쳐주세요)")
        for chart, nm, raw, sug in nonstd:
            st.write(f"- [{chart}] {nm}: `{raw}` → 표준 **`{sug}`** 로 바꿔주세요")

    with st.expander(f"환자 {len(parsed['patients'])}명 (PII 마스킹)"):
        st.table([
            {"차트": p.chart_no, "이름": (p.name[:1] + "*") if p.name else "",
             "질환": f"{p.disease}({p.disease_group})", "예약": p.booking,
             "진료결과": p.outcome, "유입": p.inflow, "상담": p.counselor}
            for p in parsed["patients"]
        ])

    st.subheader("초진 시트에 기록")
    inferred, winfo = infer_week_tab(parsed["patients"])
    tab = _week_selectbox(tabs, inferred, winfo, "wk_chojin")
    st.caption(f"`{tab}` 의 **초진 환자테이블(B5:Y)**을 이 export로 교체합니다. 상담테이블·수식·집계는 보존.")
    confirm = st.checkbox(f"'{tab}' 에 초진 {s['초진수']}명 기록 확인", key="cf_chojin")
    if st.button("📝 초진 기록하기", type="primary", disabled=not confirm, key="bt_chojin"):
        try:
            res = write_patients(sid, tab, rows, dry_run=False)
            if res.get("verify_ok"):
                st.success(f"✅ 초진 {res['rows']}명 기록 완료 — {tab}")
                st.balloons()
            else:
                st.error(f"기록됐으나 검증 불일치: {res}")
        except Exception as e:
            st.error(f"기록 실패: {e}")


# ──────────────────────────────────────────────────────────────────
# 문의 탭
# ──────────────────────────────────────────────────────────────────
def render_munui(sid: str, tabs: list):
    up = st.file_uploader("문의 export — 상담내역 (.xls/.xlsx)", type=["xls", "xlsx"], key="up_munui")
    if not up:
        st.caption("OKTAS '상담내역'을 .xls로 export 해서 올리세요.")
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
    nonstd_r = s.get("비표준결과", [])
    if nonstd_r:
        st.warning(f"✏️ **비표준 상담결과 {len(nonstd_r)}건 — 표준값(예약완료/예약안함/재통화필요)으로 교정**: "
                   + ", ".join(f"{nm} `{r}`" for nm, r in nonstd_r[:8]))

    with st.expander(f"문의 {len(parsed['inquiries'])}건 (PII 마스킹)"):
        st.table([
            {"시각": i.consult_time[5:16], "성명": (i.name[:1] + "*") if i.name else "",
             "질환": f"{i.disease}({i.disease_group})", "상담구분": i.channel,
             "결과": i.result, "상담자": i.counselor}
            for i in parsed["inquiries"]
        ])

    st.subheader("문의 시트에 기록")
    inferred, winfo = inquiry_week_tab(path)
    tab = _week_selectbox(tabs, inferred, winfo, "wk_munui")
    st.caption(f"`{tab}` 의 **상담테이블(Z5:AK)**을 이 export로 교체합니다. 초진테이블·수식·집계는 보존.")
    confirm = st.checkbox(f"'{tab}' 에 문의 {s['문의수']}건 기록 확인", key="cf_munui")
    if st.button("📝 문의 기록하기", type="primary", disabled=not confirm, key="bt_munui"):
        try:
            res = write_inquiries(sid, tab, rows, dry_run=False)
            if res.get("verify_ok"):
                st.success(f"✅ 문의 {res['rows']}건 기록 완료 — {tab}")
                st.balloons()
            else:
                st.error(f"기록됐으나 검증 불일치: {res}")
        except Exception as e:
            st.error(f"기록 실패: {e}")


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
    rows = r["rows"]
    keys = [x["key"] for x in rows]
    init = ["(자동)" if not x["overridden"] else x["status"] for x in rows]
    disp = [{"주차": x["week"], "성명": _mask(x["name"]), "질환": x["disease"],
             "문의일": str(x["time"]), "자동판정": x["auto_status"], "보정": init[i]}
            for i, x in enumerate(rows)]
    edited = st.data_editor(
        disp, hide_index=True, use_container_width=True, key="noshow_editor",
        column_config={"보정": st.column_config.SelectboxColumn(
            "보정", options=["(자동)", "전환", "노쇼", "내원대기"], required=True)},
        disabled=["주차", "성명", "질환", "문의일", "자동판정"])
    if st.button("💾 보정 저장", key="bt_override"):
        now = r["asof"]
        changed = 0
        for i, row in enumerate(edited):
            if row["보정"] != init[i]:
                set_override(sid, keys[i], "" if row["보정"] == "(자동)" else row["보정"],
                             "app", now=now)
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
    st.caption("매출·총내원(직접입력)은 안 건드림 — OKTAS 월말결산 몫. 비율은 월간탭 수식이 월 단위로 재계산.")
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
    st.caption("OKTAS 명단(초진·문의)을 올리면 주간통계가 자동으로 채워집니다 "
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
        render_chojin(sid, tabs)
    with t2:
        render_munui(sid, tabs)
    with t3:
        render_noshow(sid, tabs)
    with t4:
        render_monthly(sid)


if __name__ == "__main__":
    main()
