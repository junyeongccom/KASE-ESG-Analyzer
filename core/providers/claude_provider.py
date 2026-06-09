"""Anthropic Claude 프로바이더."""
from __future__ import annotations

import logging

import anthropic

from core.pdf_handler import prepare_for_claude
from .base import LLMProvider

logger = logging.getLogger(__name__)

DEFAULT_MODEL = "claude-sonnet-4-20250514"


class ClaudeProvider(LLMProvider):
    """Anthropic Claude API를 사용하는 프로바이더."""

    def __init__(self, api_key: str, model: str | None = None):
        super().__init__(api_key, model or DEFAULT_MODEL)
        self.client = anthropic.Anthropic(api_key=api_key)

    @property
    def provider_name(self) -> str:
        return "Claude"

    async def analyze_sheet(
        self,
        pdf_bytes: bytes,
        system_prompt: str,
        user_prompt: str,
        max_output_tokens: int = 16000,
    ) -> tuple[list[dict], dict]:
        pdf_b64 = prepare_for_claude(pdf_bytes)

        message = self.client.messages.create(
            model=self.model,
            max_tokens=max_output_tokens,
            system=system_prompt,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "document",
                            "source": {
                                "type": "base64",
                                "media_type": "application/pdf",
                                "data": pdf_b64,
                            },
                        },
                        {
                            "type": "text",
                            "text": user_prompt,
                        },
                    ],
                }
            ],
        )

        raw_text = message.content[0].text
        usage = {
            "input_tokens": message.usage.input_tokens,
            "output_tokens": message.usage.output_tokens,
        }
        logger.info("[Claude] tokens: input=%d, output=%d", usage["input_tokens"], usage["output_tokens"])
        return self.parse_response(raw_text), usage
