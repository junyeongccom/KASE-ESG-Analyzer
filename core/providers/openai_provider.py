"""OpenAI GPT 프로바이더."""
from __future__ import annotations

import logging

import openai

from core.pdf_handler import get_page_count, prepare_for_openai
from .base import LLMProvider

logger = logging.getLogger(__name__)

DEFAULT_MODEL = "gpt-5.1"
MAX_PAGES_VISION = 80


class OpenAIProvider(LLMProvider):
    """OpenAI GPT API (vision)를 사용하는 프로바이더."""

    def __init__(self, api_key: str, model: str | None = None):
        super().__init__(api_key, model or DEFAULT_MODEL)
        self.client = openai.OpenAI(api_key=api_key)

    @property
    def provider_name(self) -> str:
        return "GPT"

    async def analyze_sheet(
        self,
        pdf_bytes: bytes,
        system_prompt: str,
        user_prompt: str,
        max_output_tokens: int = 16000,
    ) -> tuple[list[dict], dict]:
        page_count = get_page_count(pdf_bytes)
        logger.info("[GPT] PDF pages: %d (limit: %d)", page_count, MAX_PAGES_VISION)

        page_images = prepare_for_openai(pdf_bytes)

        # 유저 메시지 구성: 텍스트 + PDF 페이지 이미지들
        content: list[dict] = [{"type": "text", "text": user_prompt}]

        for i, b64_img in enumerate(page_images):
            content.append(
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/jpeg;base64,{b64_img}",
                        "detail": "low",  # 토큰 절약
                    },
                }
            )

        response = self.client.chat.completions.create(
            model=self.model,
            max_completion_tokens=max_output_tokens,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": content},
            ],
        )

        raw_text = response.choices[0].message.content
        resp_usage = response.usage
        usage = {
            "input_tokens": resp_usage.prompt_tokens if resp_usage else 0,
            "output_tokens": resp_usage.completion_tokens if resp_usage else 0,
        }
        logger.info("[GPT] tokens: input=%d, output=%d", usage["input_tokens"], usage["output_tokens"])
        return self.parse_response(raw_text), usage
