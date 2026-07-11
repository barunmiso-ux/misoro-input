"""댓글 캡처 이미지를 공유 드라이브(Shared Drive)에 업로드.

서비스 계정은 개인 드라이브 용량이 없어 일반 폴더엔 업로드 불가
→ 공유 드라이브 폴더(서비스 계정을 콘텐츠 관리자로 추가)에 올린다.
폴더 ID는 config.DRIVE_FOLDER_ID 또는 st.secrets["drive_folder_id"] 로 설정.
저장된 파일은 공유 드라이브 권한(조직 내)으로 접근 — 별도 공개 설정은 하지 않는다.
"""

from __future__ import annotations

import os

from google.oauth2.service_account import Credentials

import config

DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]
_LOCAL_KEY_FILE = "service_account.json"


def _folder_id() -> str:
    if config.DRIVE_FOLDER_ID:
        return config.DRIVE_FOLDER_ID
    try:
        import streamlit as st

        return st.secrets.get("drive_folder_id", "")
    except Exception:  # noqa: BLE001
        return ""


def enabled() -> bool:
    return bool(_folder_id())


_SVC = None   # 싱글톤: 이미지·제출마다 build() 새로 안 하도록 재사용(커넥션 누적 방지)


def _service():
    global _SVC
    if _SVC is None:
        from googleapiclient.discovery import build

        if os.path.exists(_LOCAL_KEY_FILE):
            creds = Credentials.from_service_account_file(_LOCAL_KEY_FILE, scopes=DRIVE_SCOPES)
        else:
            import streamlit as st

            creds = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]), scopes=DRIVE_SCOPES
            )
        _SVC = build("drive", "v3", credentials=creds, cache_discovery=False)
    return _SVC


def _shrink(data: bytes, mimetype: str | None):
    """큰 이미지만 최대 2000px + JPEG로 축소(텍스트 가독성 유지). 작으면·실패하면 원본 그대로.
    메모리·드라이브 용량 절감용. Pillow 없거나 오류 나면 안전하게 원본 반환."""
    try:
        from io import BytesIO
        from PIL import Image
        if max(len(data), 0) <= 1_200_000:      # ~1.2MB 이하면 손 안 댐
            return data, mimetype
        im = Image.open(BytesIO(data))
        w, h = im.size
        if max(w, h) > 2000:
            r = 2000 / max(w, h)
            im = im.resize((int(w * r), int(h * r)))
        if im.mode in ("RGBA", "P", "LA"):
            im = im.convert("RGB")
        out = BytesIO()
        im.save(out, format="JPEG", quality=85, optimize=True)
        return out.getvalue(), "image/jpeg"
    except Exception:  # noqa: BLE001
        return data, mimetype


def upload_image(data: bytes, filename: str, mimetype: str | None) -> str:
    """이미지 바이트를 공유 드라이브 폴더에 업로드하고 보기 링크(webViewLink) 반환."""
    from googleapiclient.http import MediaInMemoryUpload

    data, mimetype = _shrink(data, mimetype)
    svc = _service()
    meta = {"name": filename, "parents": [_folder_id()]}
    media = MediaInMemoryUpload(data, mimetype=mimetype or "application/octet-stream")
    f = svc.files().create(
        body=meta, media_body=media,
        fields="id,webViewLink", supportsAllDrives=True,
    ).execute(num_retries=3)
    return f.get("webViewLink", "")
