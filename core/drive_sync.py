"""Google Drive의 평가지표 폴더(산업별 하위폴더)에서 템플릿 xlsx를 동기화하는 모듈.

Drive 구조:
    <평가지표 폴더>/
        평가지표_메타데이터_인벤토리.xlsx   ← 루트 파일은 건너뜀
        식품/  KASE_평가지표_식품_v4.xlsx ...
        건설/  KASE_평가지표_건설_v3.xlsx ...

산업 = 하위폴더명, 버전 = 파일명의 v토큰(없으면 정제된 stem).
다운로드 시 로컬 `templates/{산업}_{버전}.xlsx` 규칙으로 저장한다 →
다운스트림(config의 산업/버전 파싱·선택 UI)은 수정 없이 그대로 동작한다.

전제: 평가지표 폴더(또는 상위 폴더)가 서비스계정 이메일에 공유돼 있어야 한다.
"""
from __future__ import annotations

import logging
import re
import unicodedata
from pathlib import Path

import gspread

from config import GSHEET_CREDENTIALS_FILE, TEMPLATES_DIR, DRIVE_INDICATORS_FOLDER_ID

logger = logging.getLogger(__name__)

_DRIVE_FILES_URL = "https://www.googleapis.com/drive/v3/files"
_FOLDER_MIME = "application/vnd.google-apps.folder"


def _get_drive_client():
    """Google Drive 접근이 가능한 gspread 클라이언트를 반환한다."""
    # 1차: st.secrets (Streamlit Cloud)
    try:
        import streamlit as st
        if "gsheet" in st.secrets:
            return gspread.service_account_from_dict(dict(st.secrets["gsheet"]))
    except Exception:
        pass

    # 2차: 로컬 서비스계정 파일
    cred_path = Path(GSHEET_CREDENTIALS_FILE)
    if cred_path.exists():
        try:
            return gspread.service_account(filename=str(cred_path))
        except Exception as e:
            logger.error("Drive 인증 실패: %s", e)
    return None


def _list_children(client, folder_id: str) -> list[dict]:
    """폴더 직속 항목(파일+폴더) 목록을 반환한다."""
    resp = client.http_client.request(
        "get",
        _DRIVE_FILES_URL,
        params={
            "q": f"'{folder_id}' in parents and trashed=false",
            "fields": "files(id,name,mimeType)",
            "supportsAllDrives": True,
            "includeItemsFromAllDrives": True,
            "orderBy": "name",
            "pageSize": 200,
        },
    )
    data = resp.json()
    if "error" in data:
        logger.error("Drive 목록 오류: %s", data["error"])
        return []
    return data.get("files", [])


def _derive_version(stem: str, industry: str) -> str:
    """파일명 stem에서 버전 라벨을 뽑는다. v토큰(v3/v4/v5.2) 우선, 없으면 정제 stem."""
    # macOS/Drive 파일명은 NFD(자모분리)로 올 수 있어 NFC로 통일해야 한글 노이즈 제거가 먹는다.
    stem = unicodedata.normalize("NFC", stem)
    industry = unicodedata.normalize("NFC", industry)
    m = re.search(r"[vV]\d+(?:\.\d+)*", stem)
    if m:
        return m.group(0).lower()
    s = stem
    for noise in ("KASE", "평가지표", "(5기 평가지표)", "2차고도화", "통일", industry):
        s = s.replace(noise, "")
    s = re.sub(r"[\s_]+", "", s).strip("_-· ")
    return s or "v1"


def _download(client, file_id: str) -> bytes:
    resp = client.http_client.request(
        "get",
        f"{_DRIVE_FILES_URL}/{file_id}",
        params={"alt": "media", "supportsAllDrives": True},
    )
    return resp.content


def sync_templates_from_drive() -> list[str]:
    """평가지표 폴더의 산업 하위폴더 xlsx를 templates/{산업}_{버전}.xlsx 로 동기화한다.

    Returns: 동기화된 로컬 파일명 리스트.
    """
    client = _get_drive_client()
    if client is None:
        logger.info("Google Drive 인증 없음 — 템플릿 동기화 건너뜀")
        return []
    if not DRIVE_INDICATORS_FOLDER_ID:
        logger.info("DRIVE_INDICATORS_FOLDER_ID 미설정 — 동기화 건너뜀")
        return []

    try:
        children = _list_children(client, DRIVE_INDICATORS_FOLDER_ID)
        if not children:
            logger.info("평가지표 폴더가 비었거나 접근 불가 — 공유/폴더ID 확인.")
            return []

        TEMPLATES_DIR.mkdir(exist_ok=True)
        synced: list[str] = []

        for child in children:
            # 산업 = 하위폴더. 루트의 파일(인벤토리 등)은 동기화 대상이 아니다.
            if child["mimeType"] != _FOLDER_MIME:
                logger.info("루트 파일 건너뜀: %s", child["name"])
                continue

            industry = unicodedata.normalize("NFC", child["name"]).strip()
            for f in _list_children(client, child["id"]):
                name = f["name"]
                if f["mimeType"] == _FOLDER_MIME or not name.lower().endswith(".xlsx"):
                    continue
                if name.startswith("~$"):  # 엑셀 임시잠금 파일
                    continue

                version = _derive_version(Path(name).stem, industry)
                local_name = f"{industry}_{version}.xlsx"
                local_path = TEMPLATES_DIR / local_name
                try:
                    local_path.write_bytes(_download(client, f["id"]))
                    synced.append(local_name)
                    logger.info("동기화: %s/%s → %s", industry, name, local_name)
                except Exception as e:
                    logger.error("다운로드 실패 %s: %s", name, e)

        return synced

    except Exception as e:
        logger.error("Drive 평가지표 동기화 실패: %s", e)
        return []


def is_drive_configured() -> bool:
    """Google Drive 동기화가 가능한 상태인지 확인한다."""
    return _get_drive_client() is not None
