"""Google Gemini 프로바이더 (google-genai SDK)."""
from __future__ import annotations

import logging

from google import genai
from google.genai import types

from .base import LLMProvider

logger = logging.getLogger(__name__)

DEFAULT_MODEL = "gemini-2.5-flash"


class GeminiProvider(LLMProvider):
    """Google Gemini API를 사용하는 프로바이더."""

    def __init__(self, api_key: str, model: str | None = None):
        super().__init__(api_key, model or DEFAULT_MODEL)
        self.client = genai.Client(api_key=api_key)

    @property
    def provider_name(self) -> str:
        return "Gemini"

    async def analyze_sheet(
        self,
        pdf_bytes: bytes,
        system_prompt: str,
        user_prompt: str,
    ) -> tuple[list[dict], dict]:
        combined_prompt = f"{system_prompt}\n\n---\n\n{user_prompt}"

        response = self.client.models.generate_content(
            model=self.model,
            contents=[
                types.Content(
                    parts=[
                        types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                        types.Part.from_text(text=combined_prompt),
                    ]
                )
            ],
            config=types.GenerateContentConfig(
                max_output_tokens=16000,
                temperature=0.1,
                response_mime_type="application/json",
            ),
        )

        raw_text = response.text
        meta = response.usage_metadata
        usage = {
            "input_tokens": meta.prompt_token_count if meta else 0,
            "output_tokens": meta.candidates_token_count if meta else 0,
        }
        logger.info("[Gemini] tokens: input=%d, output=%d", usage["input_tokens"], usage["output_tokens"])
        return self.parse_response(raw_text), usage
