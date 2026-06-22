import os
import unicodedata
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()


def _get_secret(key: str) -> str | None:
    """st.secrets (Streamlit Cloud) → 환경변수 → .env 순으로 값을 찾는다."""
    try:
        import streamlit as st
        val = st.secrets.get(key)
        if val:
            return str(val)
    except Exception:
        pass
    return os.environ.get(key)

# ── 경로 설정 ──
BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"

TEMPLATES_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)


# ── 템플릿 관리 ──
# 파일명 규칙: {산업분류}_{버전}.xlsx  예: 식품_v5.2.xlsx, 건설_v5.3.xlsx

def get_templates_for_industry(industry: str) -> list[Path]:
    """해당 산업의 템플릿 파일 목록을 반환한다 (최신순 정렬). 한글 정규화(NFC/NFD) 무관."""
    target = unicodedata.normalize("NFC", industry)
    files = [
        f for f in TEMPLATES_DIR.glob("*.xlsx")
        if unicodedata.normalize("NFC", f.stem).split("_", 1)[0] == target
    ]
    return sorted(files, reverse=True)


def get_all_industries_with_templates() -> list[str]:
    """templates/ 폴더에 있는 산업 분류 목록을 반환한다 (NFC 정규화로 중복 제거)."""
    industries = set()
    for f in TEMPLATES_DIR.glob("*.xlsx"):
        parts = f.stem.split("_", 1)
        if parts:
            industries.add(unicodedata.normalize("NFC", parts[0]))
    return sorted(industries)


def get_template_version(template_path: Path) -> str:
    """템플릿 파일명에서 버전을 추출한다. 예: 식품_v5.2.xlsx → v5.2"""
    stem = template_path.stem
    parts = stem.split("_", 1)
    return parts[1] if len(parts) > 1 else stem

# ── 플레이스홀더 (덮어쓰기 대상) ──
PLACEHOLDERS = [
    "— AI 분석 대기 —",
    "분석 전",
    "(응답 누락)",
    "(API 오류)",
]

# ── Row 1 타이틀 치환용 ──
TITLE_PATTERN_PLACEHOLDER = "[기업명]"
TITLE_PATTERN_HARDCODED = "CJ제일제당"

# ── 산업 분류 ──
INDUSTRIES = ["식품", "건설", "화학", "제약/바이오", "IT/전자", "금융", "에너지", "기타"]

# ── 병렬 처리 ──
MAX_CONCURRENT = 5

# ── LLM 모델 설정 ──
MODELS = {
    "Claude": {
        "env_key": "ANTHROPIC_API_KEY",
        "default_model": "claude-sonnet-4-20250514",
        "provider": "claude",
    },
    "GPT": {
        "env_key": "OPENAI_API_KEY",
        "default_model": "gpt-5.1",
        "provider": "openai",
    },
    "Gemini": {
        "env_key": "GOOGLE_API_KEY",
        "default_model": "gemini-2.5-flash",
        "provider": "gemini",
    },
}


# ── 토큰 단가 ($/1M tokens) ──
PRICING = {
    # Claude
    "claude-sonnet-4-20250514": {"input": 3.00, "output": 15.00},
    # OpenAI
    "gpt-5.1":  {"input": 1.25, "output": 10.00},
    "gpt-5":    {"input": 1.25, "output": 10.00},
    "gpt-5.2":  {"input": 1.75, "output": 14.00},
    "gpt-4.1":  {"input": 2.00, "output": 8.00},
    # Gemini
    "gemini-2.5-flash": {"input": 0.30, "output": 2.50},
    "gemini-2.5-pro": {"input": 1.25, "output": 10.00},  # 판정 폴백. >200k 입력 시 실제 단가 더 높음
    "gemini-3.1-pro-preview": {"input": 2.00, "output": 12.00},  # 판정자(thinking). >200k 입력 시 $4/$18
}


def calc_cost(model: str, input_tokens: int, output_tokens: int) -> dict:
    """토큰 수와 모델명으로 비용(USD)을 계산한다."""
    prices = PRICING.get(model, {"input": 0, "output": 0})
    input_cost = input_tokens / 1_000_000 * prices["input"]
    output_cost = output_tokens / 1_000_000 * prices["output"]
    return {
        "input_tokens": input_tokens,
        "output_tokens": output_tokens,
        "input_cost": input_cost,
        "output_cost": output_cost,
        "total_cost": input_cost + output_cost,
    }


# ── Google Sheets 이력 저장 ──
GSHEET_CREDENTIALS_FILE = _get_secret("GSHEET_CREDENTIALS_FILE") or str(
    BASE_DIR / "credentials" / "gsheet_service_account.json"
)
GSHEET_SPREADSHEET_NAME = _get_secret("GSHEET_SPREADSHEET_NAME") or "KASE ESG 분석 이력"
# 이력을 쌓을 스프레드시트 ID(권장). 설정 시 이름 대신 이 시트를 직접 연다.
# 시트 URL `.../spreadsheets/d/<여기>/edit` 부분. 서비스계정 이메일에 '편집자'로 공유 필요.
GSHEET_SPREADSHEET_ID = _get_secret("GSHEET_SPREADSHEET_ID") or "1TDL5f0W6lkOKtE9Mjmf6al_m9w9GziNpkwYiLy-bHP4"

# ── Google Drive 평가지표 동기화 ──
# 산업별 하위폴더(식품/, 건설/ …)를 담은 "평가지표" 폴더의 ID.
# Drive에서 그 폴더를 열었을 때 URL `.../folders/<여기>` 부분. 비우면 동기화를 건너뛴다.
# 이 폴더(또는 상위 폴더)를 서비스계정 이메일에 공유해야 앱이 읽을 수 있다.
DRIVE_INDICATORS_FOLDER_ID = _get_secret("DRIVE_INDICATORS_FOLDER_ID") or "1-pErqxOsqGcY5D9S_IHq3f_WXMgywTLo"


def get_api_key(model_name: str) -> str | None:
    cfg = MODELS.get(model_name)
    if cfg is None:
        return None
    return _get_secret(cfg["env_key"])


def get_available_models() -> list[str]:
    return [name for name in MODELS if get_api_key(name)]
