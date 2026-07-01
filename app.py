"""미소로한의원 일일 보고 입력 — Streamlit 앱.

Phase 1: 폼 UI + 검증 + 요약. 시트 쓰기는 mock (USE_MOCK=True).
Phase 2: USE_MOCK=False 로 sheet_writer.write_submission 연결.

실행: streamlit run app.py
지점 지정: URL 에 ?branch=분당
"""

from __future__ import annotations

from datetime import date, datetime, timedelta

import streamlit as st

import config

# 실제 시트 쓰기 사용. (시트 접근 불가 환경에서 UI만 보려면 True)
USE_MOCK = False

st.set_page_config(page_title="미소로 일일 보고", page_icon="🩺", layout="centered")

# 기본 여백이 넓어 입력 폼이 늘어져 보임 → 밀도 높이는 CSS
st.markdown(
    """
    <style>
      .block-container {padding-top: 2.2rem; padding-bottom: 3rem;}
      /* 세로 요소 간격 축소 */
      div[data-testid="stVerticalBlock"] {gap: 0.45rem;}
      /* 동적 항목 한 줄(가로 컬럼) 간격 축소 */
      div[data-testid="stHorizontalBlock"] {gap: 0.35rem;}
      /* 위젯 라벨 여백 축소 */
      div[data-testid="stWidgetLabel"] {margin-bottom: 0.1rem;}
      /* 구분선 얇게 */
      hr {margin: 0.5rem 0;}
      /* 제목/소제목 위아래 여백 축소 */
      h1 {padding-top: 0; margin-bottom: 0.3rem; font-size: 1.6rem;}
      h2, h3 {margin-top: 0.4rem; margin-bottom: 0.2rem;}
      /* 입력 위젯 자체 높이 약간 축소 */
      div[data-testid="stTextInput"] input,
      div[data-baseweb="select"] > div {min-height: 2.1rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

WEEKDAYS_KO = ["월", "화", "수", "목", "금", "토", "일"]


# ──────────────────────────────────────────────────────────────────
# 지점 결정 (URL 파라미터 우선)
# ──────────────────────────────────────────────────────────────────
def resolve_branch() -> str | None:
    qp = st.query_params.get("branch")
    if qp in config.ALL_BRANCHES:
        return qp
    # 파라미터 없거나 잘못된 경우: 선택 박스
    st.info("지점이 URL 에 지정되지 않았습니다. 지점을 선택하세요. (예: `?branch=분당`)")
    choice = st.selectbox("지점", ["선택..."] + config.ALL_BRANCHES, index=0)
    return choice if choice in config.ALL_BRANCHES else None


# ──────────────────────────────────────────────────────────────────
# 세션 상태 초기화 (동적 리스트)
# ──────────────────────────────────────────────────────────────────
def init_state():
    for sec in config.INQUIRY_SECTIONS:
        key = f"items_{sec['key']}"
        if key not in st.session_state:
            st.session_state[key] = [_empty_item()]
    if "comments" not in st.session_state:
        st.session_state["comments"] = [""]


def _empty_item():
    return {"name": config.PLACEHOLDER, "route": config.PLACEHOLDER, "count": 0}


def _sel_index(options, value):
    return options.index(value) if value in options else 0


# ──────────────────────────────────────────────────────────────────
# 기존 제출 조회 + prefill (수정 모드)
# ──────────────────────────────────────────────────────────────────
@st.cache_data(ttl=30, show_spinner=False)
def _cached_existing(branch: str, date_iso: str):
    """(branch, 날짜) 기존 기록 조회. 30초 캐시로 API 호출 폭주 방지."""
    if USE_MOCK:
        return None
    import sheet_writer

    from datetime import date as _date

    return sheet_writer.read_existing(branch, _date.fromisoformat(date_iso))


def apply_prefill(branch, report_date):
    """branch/날짜가 바뀌면 기존 기록을 불러와 활동 칸을 채운다. 반환: existing dict|None."""
    tag = (branch, report_date.isoformat())
    if st.session_state.get("loaded_for") == tag:
        return st.session_state.get("existing_info")

    try:
        existing = _cached_existing(branch, report_date.isoformat())
    except Exception:  # noqa: BLE001
        existing = None
        st.warning("기존 기록을 불러오지 못했습니다 (일시적 오류일 수 있어요). 입력·제출에는 지장 없습니다.")
    act = existing["activity"] if existing else None

    # 동적 댓글 위젯 키 정리 (날짜 변경 시 잔상 제거)
    for k in [x for x in st.session_state if x.startswith("comment_")]:
        del st.session_state[k]

    st.session_state["health_cafe"] = act["health_cafe"] if act else ""
    st.session_state["health_home"] = act["health_home"] if act else ""
    st.session_state["daily_post"] = act["daily_post"] if act else ""
    st.session_state["comments"] = (act["comments"] or [""]) if act else [""]

    st.session_state["loaded_for"] = tag
    st.session_state["existing_info"] = existing
    return existing


# ──────────────────────────────────────────────────────────────────
# 초진·문의 섹션 렌더링
# ──────────────────────────────────────────────────────────────────
def render_section(sec: dict):
    key = f"items_{sec['key']}"
    items = st.session_state[key]

    st.markdown(f"**{sec['label']}**")
    subtotal = 0
    remove_idx = None

    for i, item in enumerate(items):
        c_route = None
        if sec["allow_route"]:      # 초진: 병명 / 경로 / 인원 / ✕
            c_name, c_route, c_cnt, c_del = st.columns([4, 4, 3, 1])
        else:                       # 그 외: 병명 / 인원 / ✕
            c_name, c_cnt, c_del = st.columns([5, 3, 1])

        # 병명: 드롭다운 (자유 텍스트 불가, 모든 채널 공통 세부 병명)
        item["name"] = c_name.selectbox(
            "병명", config.DISEASE_OPTIONS,
            index=_sel_index(config.DISEASE_OPTIONS, item.get("name", config.PLACEHOLDER)),
            key=f"{key}_name_{i}", label_visibility="collapsed",
        )
        # 경로: 드롭다운 (초진 한정)
        if c_route is not None:
            item["route"] = c_route.selectbox(
                "경로", config.ROUTE_OPTIONS,
                index=_sel_index(config.ROUTE_OPTIONS, item.get("route", config.PLACEHOLDER)),
                key=f"{key}_route_{i}", label_visibility="collapsed",
            )
        # 인원: 숫자만 (텍스트 입력 불가 → 카운트 폭주 원천 차단)
        item["count"] = c_cnt.number_input(
            "인원", min_value=0, max_value=999, step=1, value=int(item["count"]),
            key=f"{key}_cnt_{i}", label_visibility="collapsed",
        )
        if c_del.button("✕", key=f"{key}_del_{i}", help="이 항목 삭제"):
            remove_idx = i

        subtotal += int(item["count"])

    cols = st.columns([3, 2])
    if cols[0].button("➕ 항목 추가", key=f"{key}_add"):
        items.append(_empty_item())
        st.rerun()

    cols[1].markdown(f"소계: **{subtotal}명**")

    if remove_idx is not None:
        items.pop(remove_idx)
        if not items:
            items.append(_empty_item())
        st.rerun()

    st.divider()
    return subtotal


# ──────────────────────────────────────────────────────────────────
# 활동 섹션 렌더링
# ──────────────────────────────────────────────────────────────────
def render_activity():
    st.subheader("활동")

    st.markdown("**건강 칼럼**")
    health_cafe = st.text_input("카페 URL", key="health_cafe", placeholder="https://...")
    health_home = st.text_input("홈페이지 URL", key="health_home", placeholder="https://...")

    daily_post = st.text_input(
        "카페 일상 게시글 URL", key="daily_post", placeholder="https://...",
    )

    st.markdown("**카페 댓글**")
    comments = st.session_state["comments"]
    remove_idx = None
    for i, _ in enumerate(comments):
        c_url, c_del = st.columns([8, 1])
        comments[i] = c_url.text_input(
            f"댓글 {i + 1}", value=comments[i], key=f"comment_{i}",
            label_visibility="collapsed", placeholder=f"댓글 {i + 1} URL",
        )
        if c_del.button("✕", key=f"comment_del_{i}", help="댓글 삭제"):
            remove_idx = i

    if st.button("➕ 댓글 추가", key="comment_add"):
        comments.append("")
        st.rerun()

    if remove_idx is not None:
        comments.pop(remove_idx)
        if not comments:
            comments.append("")
        st.rerun()

    # 댓글 캡처 이미지 (사진으로 보내는 지점용) → 공유 드라이브 업로드
    import drive_uploader

    images = st.file_uploader(
        "댓글 캡처 이미지 (선택, 여러 장 가능)", type=config.IMAGE_TYPES,
        accept_multiple_files=True, key="comment_images",
    )
    image_comment_count = 0
    if images:
        if not drive_uploader.enabled():
            st.caption("⚠️ 공유 드라이브 폴더가 아직 설정되지 않아 이미지는 저장되지 않습니다.")
        image_comment_count = st.number_input(
            "이미지 속 댓글 수", min_value=0, max_value=999, step=1,
            value=len(images), key="img_comment_count",
            help="캡처 안에 들어있는 실제 댓글 개수를 적어주세요.",
        )

    return {
        "health_cafe": health_cafe.strip(),
        "health_home": health_home.strip(),
        "daily_post": daily_post.strip(),
        "comments": [c.strip() for c in comments if c.strip()],
        "images": images or [],
        "image_comment_count": int(image_comment_count),
    }


# ──────────────────────────────────────────────────────────────────
# 검증 + payload 빌드
# ──────────────────────────────────────────────────────────────────
def validate_url(label: str, url: str, errors: list):
    if url and not url.startswith("https://"):
        errors.append(f"{label} 은(는) https:// 로 시작해야 합니다.")


def build_payload(branch, report_date, activity) -> tuple[dict, list]:
    errors: list[str] = []
    sections = {}
    for sec in config.INQUIRY_SECTIONS:
        items = []
        for it in st.session_state[f"items_{sec['key']}"]:
            name = it.get("name", config.PLACEHOLDER)
            route = it.get("route", config.PLACEHOLDER)
            cnt = int(it["count"])
            name_ok = name != config.PLACEHOLDER and name not in config.GROUP_DIVIDERS
            route_ok = route != config.PLACEHOLDER and route not in config.GROUP_DIVIDERS
            if cnt > 0 and not name_ok:
                errors.append(f"[{sec['label']}] 인원이 입력된 항목의 병명을 선택하세요.")
            if sec["allow_route"] and (name_ok or cnt > 0) and not route_ok:
                errors.append(f"[{sec['label']}] 유입경로를 선택하세요 (필수). 병명: {name if name_ok else '미선택'}")
            if name_ok or cnt > 0:
                entry = {"name": name if name_ok else "", "count": cnt}
                if sec["allow_route"]:
                    entry["route"] = route if route_ok else ""
                items.append(entry)
        subtotal = sum(it["count"] for it in items)
        sections[sec["key"]] = {"items": items, "subtotal": subtotal}

    validate_url("건강 카페 URL", activity["health_cafe"], errors)
    validate_url("건강 홈페이지 URL", activity["health_home"], errors)
    validate_url("카페 일상 게시글 URL", activity["daily_post"], errors)
    for i, c in enumerate(activity["comments"], 1):
        validate_url(f"댓글 {i} URL", c, errors)

    payload = {
        "branch": branch,
        "date": report_date,
        "writer": "",
        "sections": sections,
        "activity": activity,
    }
    return payload, errors


def summarize(payload: dict) -> str:
    s = payload["sections"]
    parts = []
    for sec in config.INQUIRY_SECTIONS:
        st_ = s[sec["key"]]["subtotal"]
        if st_:
            parts.append(f"{sec['label']}{st_}")
    counts = " ".join(parts) if parts else "(문의 없음)"
    a = payload["activity"]
    mark = lambda v: "✓" if v else "✗"
    comment_n = a.get("comment_count", len(a["comments"]))
    act = (
        f"건강[카페{mark(a['health_cafe'])} 홈{mark(a['health_home'])}] "
        f"일상{mark(a['daily_post'])} "
        f"댓글{comment_n}건"
    )
    return (
        f"{payload['branch']} {payload['date']:%m.%d} 기록 완료: "
        f"{counts} | 활동: {act}"
    )


# ──────────────────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────────────────
def main():
    branch = resolve_branch()
    if not branch:
        st.stop()

    today = date.today()
    st.title(f"미소로한의원 일일 보고 — {branch}")

    report_date = st.date_input(
        "보고 날짜",
        value=today,
        min_value=today - timedelta(days=config.PAST_DAYS_ALLOWED),
        max_value=today,
        format="YYYY.MM.DD",
    )
    wd = WEEKDAYS_KO[report_date.weekday()]
    st.caption(f"{report_date:%Y.%m.%d} ({wd})")

    existing = apply_prefill(branch, report_date)
    if existing:
        daily = existing.get("daily") or {}
        nz = ", ".join(f"{k}{v}" for k, v in daily.items() if v.strip() and v.strip() != "0")
        has_act = existing.get("activity") is not None
        st.warning(
            f"이미 이 날짜의 보고가 있습니다 (일일: {nz or '카운트 없음'}"
            f"{' · 활동 기록 있음' if has_act else ''}). "
            "다시 제출하면 덮어쓰기됩니다. 활동 칸은 기존 값으로 채워두었습니다."
        )

    init_state()

    st.subheader("초진·문의")
    for sec in config.INQUIRY_SECTIONS:
        render_section(sec)

    activity = render_activity()

    st.divider()
    if st.button("제출하기", type="primary", use_container_width=True):
        payload, errors = build_payload(branch, report_date, activity)
        if errors:
            for e in errors:
                st.error(e)
            st.stop()

        # 댓글 이미지 업로드 → 링크를 댓글 목록에 합치고 댓글수 보정
        act = payload["activity"]
        url_comments = act["comments"]
        image_links = []
        if act["images"]:
            import drive_uploader

            if not drive_uploader.enabled():
                st.error("댓글 이미지를 저장할 공유 드라이브 폴더가 설정되지 않았습니다.")
                st.stop()
            try:
                with st.spinner("이미지 업로드 중..."):
                    for f in act["images"]:
                        image_links.append(drive_uploader.upload_image(f.getvalue(), f.name, f.type))
            except Exception as ex:  # noqa: BLE001
                st.error(f"이미지 업로드 실패: {ex}")
                st.stop()
        act["comments"] = url_comments + image_links
        act["comment_count"] = len(url_comments) + act.get("image_comment_count", 0)
        act.pop("images", None)  # 기록/직렬화 불필요

        if USE_MOCK:
            st.success("✅ (mock) 제출 완료 — 시트 쓰기는 Phase 2 에서 연결됩니다.")
            st.info(summarize(payload))
            with st.expander("제출 payload 확인 (mock)"):
                st.json(_jsonable(payload))
        else:
            import sheet_writer

            try:
                res = sheet_writer.write_submission(payload)
                st.success("✅ " + summarize(payload))
                if not res.get("daily_range"):
                    st.info("이번 달 일일집계 탭이 아직 준비 안 돼 상세·활동만 기록됐어요. 관리자에게 알려주세요 (곧 반영됩니다).")
                _cached_existing.clear()           # 캐시 무효화
                st.session_state["loaded_for"] = None  # 배너 갱신
            except Exception as ex:  # noqa: BLE001
                st.error(f"시트 기록 실패: {ex}")


def _jsonable(payload: dict) -> dict:
    p = dict(payload)
    p["date"] = payload["date"].isoformat()
    return p


if __name__ == "__main__":
    main()
