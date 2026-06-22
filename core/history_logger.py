"""분석 이력을 Google Sheets에 영구 저장하는 모듈."""
from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path

import gspread

from config import (
    GSHEET_CREDENTIALS_FILE,
    GSHEET_SPREADSHEET_NAME,
    GSHEET_SPREADSHEET_ID,
)

logger = logging.getLogger(__name__)

# 헤더 행 (스프레드시트 첫 행)
HEADERS = [
    "분석일시",
    "기업명",
    "산업분류",
    "보고서연도",
    "지표버전",
    "모델",
    "분석시트",
    "시트수",
    "소요시간",
    "Input 토큰",
    "Output 토큰",
    "총 토큰",
    "Input 비용($)",
    "Output 비용($)",
    "총 비용($)",
    "결과파일",
]


def _get_client() -> gspread.Client | None:
    """서비스 계정으로 gspread 클라이언트를 생성한다.
    1차: st.secrets["gsheet"] (Streamlit Cloud)
    2차: 로컬 JSON 파일
    """
    # 1차: Streamlit Cloud secrets
    try:
        import streamlit as st
        if "gsheet" in st.secrets:
            creds = dict(st.secrets["gsheet"])
            return gspread.service_account_from_dict(creds)
    except Exception:
        pass

    # 2차: 로컬 파일
    cred_path = Path(GSHEET_CREDENTIALS_FILE)
    if not cred_path.exists():
        logger.warning("Google Sheets 인증 파일이 없습니다: %s", cred_path)
        return None
    try:
        return gspread.service_account(filename=str(cred_path))
    except Exception as e:
        logger.error("Google Sheets 인증 실패: %s", e)
        return None


def _open_spreadsheet(client: gspread.Client) -> gspread.Spreadsheet:
    """이력 스프레드시트를 연다.
    - GSHEET_SPREADSHEET_ID가 있으면 그 시트를 직접 연다(권장, KASE 공용 시트).
    - 없으면 이름으로 열고, 없으면 새로 생성(구 동작 — 서비스계정 소유).
    """
    if GSHEET_SPREADSHEET_ID:
        return client.open_by_key(GSHEET_SPREADSHEET_ID)
    try:
        return client.open(GSHEET_SPREADSHEET_NAME)
    except gspread.SpreadsheetNotFound:
        ss = client.create(GSHEET_SPREADSHEET_NAME)
        logger.info("새 스프레드시트 생성: %s", GSHEET_SPREADSHEET_NAME)
        return ss


def _get_or_create_sheet(client: gspread.Client) -> gspread.Worksheet:
    """스프레드시트를 열거나, 없으면 헤더를 포함해 첫 시트를 초기화한다."""
    spreadsheet = _open_spreadsheet(client)
    worksheet = spreadsheet.sheet1

    # 헤더가 없으면 추가
    existing = worksheet.row_values(1)
    if not existing or existing[0] != HEADERS[0]:
        worksheet.update("A1", [HEADERS])
        worksheet.format("A1:P1", {"textFormat": {"bold": True}})
        logger.info("헤더 행 초기화 완료")

    return worksheet


def _extract_report_year(pdf_filename: str) -> str:
    """파일명에서 보고서 연도를 추출. 예: 2024_CJ제일제당_... → 2024"""
    import re
    m = re.search(r"(20\d{2})", pdf_filename)
    return m.group(1) if m else "N/A"


def log_analysis(
    pdf_filename: str,
    company_name: str,
    industry: str,
    model_name: str,
    selected_sheets: list[str],
    cost_summary: dict,
    template_version: str = "",
    result_filename: str = "",
) -> bool:
    """분석 이력 1건을 Google Sheets에 추가한다.

    Returns True if successful, False otherwise.
    """
    client = _get_client()
    if client is None:
        logger.warning("Google Sheets 이력 저장 건너뜀 (인증 없음)")
        return False

    try:
        worksheet = _get_or_create_sheet(client)

        elapsed = cost_summary.get("total_elapsed_sec", 0)
        m, s = divmod(int(elapsed), 60)
        elapsed_str = f"{m}분 {s}초"

        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            company_name,
            industry,
            _extract_report_year(pdf_filename),
            template_version,
            model_name,
            ", ".join(selected_sheets),
            len(selected_sheets),
            elapsed_str,
            cost_summary.get("total_input_tokens", 0),
            cost_summary.get("total_output_tokens", 0),
            cost_summary.get("total_tokens", 0),
            round(cost_summary.get("total_input_cost", 0), 4),
            round(cost_summary.get("total_output_cost", 0), 4),
            round(cost_summary.get("total_cost", 0), 4),
            result_filename,
        ]

        worksheet.append_row(row, value_input_option="USER_ENTERED")
        logger.info("Google Sheets 이력 저장 완료: %s", company_name)
        return True

    except Exception as e:
        logger.error("Google Sheets 이력 저장 실패: %s", e)
        return False


def is_gsheet_configured() -> bool:
    """Google Sheets 인증이 설정되어 있는지 확인한다."""
    # st.secrets 확인
    try:
        import streamlit as st
        if "gsheet" in st.secrets:
            return True
    except Exception:
        pass
    # 로컬 파일 확인
    return Path(GSHEET_CREDENTIALS_FILE).exists()
