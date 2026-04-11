"""LLM 프로바이더 추상 기반 클래스."""
from __future__ import annotations

import json
import logging
import re
from abc import ABC, abstractmethod

from json_repair import repair_json

logger = logging.getLogger(__name__)


class LLMProvider(ABC):
    """모든 LLM 프로바이더의 추상 기반 클래스."""

    def __init__(self, api_key: str, model: str | None = None):
        self.api_key = api_key
        self.model = model

    @property
    @abstractmethod
    def provider_name(self) -> str:
        ...

    @abstractmethod
    async def analyze_sheet(
        self,
        pdf_bytes: bytes,
        system_prompt: str,
        user_prompt: str,
    ) -> tuple[list[dict], dict]:
        """PDF와 프롬프트를 받아 (분석 결과, 토큰 사용량)을 반환한다."""
        ...

    def parse_response(self, raw_text: str) -> list[dict]:
        """LLM 응답에서 JSON을 추출하고 파싱한다."""
        # JSON 블록 추출
        m = re.search(r"```json\s*\n?(.*?)```", raw_text, re.DOTALL)
        if m:
            json_str = m.group(1).strip()
        else:
            m = re.search(r"[\{\[].*[\}\]]", raw_text, re.DOTALL)
            if m:
                json_str = m.group(0).strip()
            else:
                raise ValueError("JSON을 찾을 수 없습니다.")

        # 1차: 표준 파싱
        try:
            data = json.loads(json_str)
        except json.JSONDecodeError:
            # 2차: json-repair로 자동 복구
            logger.info("JSON 자동 복구 시도 중...")
            repaired = repair_json(json_str)
            data = json.loads(repaired)
            logger.info("JSON 자동 복구 성공")

        if isinstance(data, dict) and "results" in data:
            results = data["results"]
        elif isinstance(data, list):
            results = data
        else:
            raise ValueError(f"예상치 못한 JSON 구조: {type(data)}")

        validated: list[dict] = []
        for item in results:
            if not isinstance(item, dict):
                continue
            validated.append(
                {
                    "indicator_number": str(item.get("indicator_number", "")),
                    "e_content": str(item.get("e_content", "")),
                    "f_score": item.get("f_score", 0),
                    "g_review": str(item.get("g_review", "")),
                }
            )
        return validated
