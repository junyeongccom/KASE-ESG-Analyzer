"""Phase 1: v3·v4 전체 시트 CJ SR 분석 (앱과 동일한 run_analysis 경로).
단계별 로그를 stdout + logs/phase1_run.log 에 남긴다. 언제든 재실행 가능."""
import sys, time, logging, traceback
from pathlib import Path

BASE = Path("/Users/junyeongc/KASE/AI도입/KASE_ESG_Analyzer")
sys.path.insert(0, str(BASE))

LOGDIR = BASE / "logs"
LOGDIR.mkdir(exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOGDIR / "phase1_run.log", mode="w", encoding="utf-8"),
    ],
)
log = logging.getLogger("run_full")

from dotenv import dotenv_values
from core.providers import create_provider
from core.analyzer import run_analysis_sync
from core.excel_handler import get_sheet_names
from core.schemas import detect_schema

ENV = dotenv_values(str(BASE / ".env"))
PDF = Path("/Users/junyeongc/KASE/AI도입/03_원본보고서_PDF/2024_CJ제일제당_sustainability_report_ko.pdf").read_bytes()
log.info("PDF %.1fMB 로드", len(PDF) / 1024 / 1024)
prov = create_provider(provider_key="gemini", api_key=ENV["GOOGLE_API_KEY"], model="gemini-2.5-flash")

CASES = [
    ("v3", BASE / "templates/식품_v3.xlsx", ["식품-E", "식품-S"]),
    ("v4", BASE / "templates/식품_v4.xlsx", ["식품_환경", "식품_사회"]),
]
results = {}
for tag, tpl, sheets in CASES:
    log.info("\n############### %s 분석(E/S) │ schema=%s │ 시트=%s ###############",
             tag, detect_schema(tpl), sheets)
    t = time.time()
    try:
        out, cost = run_analysis_sync(
            pdf_bytes=PDF, company_name=f"CJ_{tag}",
            selected_sheets=sheets, provider=prov, template_path=str(tpl))
        log.info("############### %s 완료 │ %.0f초 │ $%.4f │ %s ###############",
                 tag, time.time() - t, cost.get("total_cost", 0), Path(out).name)
        results[tag] = str(out)
    except Exception as e:
        log.error("############### %s 실패: %s ###############\n%s", tag, e, traceback.format_exc())

log.info("\n결과 파일: %s", results)
