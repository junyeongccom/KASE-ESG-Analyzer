"""Google Drive에서 템플릿 파일을 동기화하는 모듈."""
from __future__ import annotations

import logging
from pathlib import Path

import gspread

from config import GSHEET_CREDENTIALS_FILE, TEMPLATES_DIR

logger = logging.getLogger(__name__)

# Google Drive 폴더명 (이 이름의 폴더를 서비스 계정에 공유해야 함)
DRIVE_FOLDER_NAME = "KASE_templates"


def _get_drive_client():
    """Google Drive 접근이 가능한 gspread 클라이언트를 반환한다."""
    # 1차: st.secrets
    try:
        import streamlit as st
        if "gsheet" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gsheet"]))
    except Exception:
        pass

    # 2차: 로컬 파일
    cred_path = Path(GSHEET_CREDENTIALS_FILE)
    if cred_path.exists():
        try:
            return gspread.service_account(filename=str(cred_path))
        except Exception as e:
            logger.error("Drive 인증 실패: %s", e)
    return None


def sync_templates_from_drive() -> list[str]:
    """Google Drive의 KASE_templates 폴더에서 xlsx 파일을 templates/로 동기화한다.

    Returns: 동기화된 파일명 리스트
    """
    client = _get_drive_client()
    if client is None:
        logger.info("Google Drive 인증 없음 — 템플릿 동기화 건너뜀")
        return []

    try:
        # Drive API로 폴더 내 파일 목록 조회
        response = client.http_client.request(
            "get",
            "https://www.googleapis.com/drive/v3/files",
            params={
                "q": f"name='{DRIVE_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
                "fields": "files(id,name)",
            },
        )
        folders = response.json().get("files", [])
        if not folders:
            logger.info("Drive에 '%s' 폴더가 없습니다.", DRIVE_FOLDER_NAME)
            return []

        folder_id = folders[0]["id"]

        # 폴더 내 xlsx 파일 목록
        response = client.http_client.request(
            "get",
            "https://www.googleapis.com/drive/v3/files",
            params={
                "q": f"'{folder_id}' in parents and name contains '.xlsx' and trashed=false",
                "fields": "files(id,name,modifiedTime)",
            },
        )
        files = response.json().get("files", [])

        if not files:
            logger.info("Drive 폴더에 xlsx 파일이 없습니다.")
            return []

        synced = []
        TEMPLATES_DIR.mkdir(exist_ok=True)

        for file_info in files:
            file_id = file_info["id"]
            file_name = file_info["name"]
            local_path = TEMPLATES_DIR / file_name

            # 파일 다운로드
            response = client.http_client.request(
                "get",
                f"https://www.googleapis.com/drive/v3/files/{file_id}",
                params={"alt": "media"},
            )

            local_path.write_bytes(response.content)
            synced.append(file_name)
            logger.info("템플릿 동기화: %s", file_name)

        return synced

    except Exception as e:
        logger.error("Drive 템플릿 동기화 실패: %s", e)
        return []


def is_drive_configured() -> bool:
    """Google Drive 동기화가 가능한 상태인지 확인한다."""
    return _get_drive_client() is not None
