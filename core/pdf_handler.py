"""PDF 처리: 기업명 추출 + API별 변환."""
from __future__ import annotations

import base64
import io
import re
from pathlib import Path

import fitz  # PyMuPDF


# ── 기업명 추출 ──


def extract_company_name(pdf_filename: str) -> str:
    """PDF 파일명에서 기업명을 추출한다.

    지원 패턴:
      - 2024_CJ제일제당_sustainability_report_ko.pdf  → CJ제일제당
      - CJ제일제당_2024_SR.pdf                       → CJ제일제당
      - 그 외 → 확장자 제거 후 전체 반환
    """
    stem = Path(pdf_filename).stem

    # 패턴 1: {연도}_{기업명}_{나머지}
    m = re.match(r"\d{4}[_\-](.+?)[_\-](?:sustainability|sr|report|ESG)", stem, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # 패턴 2: {기업명}_{연도}_{나머지}
    m = re.match(r"(.+?)[_\-]\d{4}[_\-]", stem)
    if m:
        return m.group(1).strip()

    # 패턴 3: 언더스코어로 분리, 첫 한글 포함 부분
    parts = re.split(r"[_\-]", stem)
    for part in parts:
        if re.search(r"[가-힣]", part):
            return part.strip()

    return stem


# ── Claude용: PDF를 base64 인코딩 ──


def prepare_for_claude(pdf_bytes: bytes) -> str:
    """PDF 바이트를 base64 문자열로 반환 (Claude document type)."""
    return base64.standard_b64encode(pdf_bytes).decode("utf-8")


# ── OpenAI용: PDF 페이지를 이미지로 변환 ──

MAX_IMAGE_TOTAL_MB = 45  # OpenAI 제한 50MB보다 여유 확보


def prepare_for_openai(pdf_bytes: bytes, dpi: int = 100) -> list[str]:
    """PDF를 페이지별 JPEG 이미지 base64 리스트로 변환.

    총 이미지 크기가 MAX_IMAGE_TOTAL_MB를 초과하지 않도록 자동 조절.
    """
    from PIL import Image

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images: list[str] = []
    total_bytes = 0
    max_bytes = MAX_IMAGE_TOTAL_MB * 1024 * 1024

    for i in range(len(doc)):
        try:
            page = doc.load_page(i)
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat)

            # PyMuPDF → PIL → JPEG (더 안정적)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=75)
            img_bytes = buf.getvalue()

            total_bytes += len(img_bytes)
            if total_bytes > max_bytes:
                break

            b64 = base64.standard_b64encode(img_bytes).decode("utf-8")
            images.append(b64)
        except Exception:
            continue  # 변환 실패한 페이지는 건너뜀

    doc.close()
    return images


# ── Gemini용: PDF를 바이트 그대로 ──


def prepare_for_gemini(pdf_bytes: bytes) -> bytes:
    """Gemini는 inline PDF를 지원하므로 바이트 그대로 반환."""
    return pdf_bytes


# ── PDF 페이지 수 확인 ──


def get_page_count(pdf_bytes: bytes) -> int:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    count = len(doc)
    doc.close()
    return count
