from .base import LLMProvider
from .claude_provider import ClaudeProvider
from .openai_provider import OpenAIProvider
from .gemini_provider import GeminiProvider

PROVIDER_MAP = {
    "claude": ClaudeProvider,
    "openai": OpenAIProvider,
    "gemini": GeminiProvider,
}


def create_provider(provider_key: str, api_key: str, model: str | None = None) -> LLMProvider:
    cls = PROVIDER_MAP.get(provider_key)
    if cls is None:
        raise ValueError(f"Unknown provider: {provider_key}")
    return cls(api_key=api_key, model=model)
