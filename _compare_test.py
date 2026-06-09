"""GPT가 '해상도 문제'로 보고한 지표를 Claude/Gemini로 다시 돌려 비교."""
from __future__ import annotations
import asyncio
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from dotenv import dotenv_values
from core.excel_handler import load_template
from core.prompt_builder import SYSTEM_PROMPT, build_user_prompt
from core.providers.claude_provider import ClaudeProvider
from core.providers.gemini_provider import GeminiProvider

ENV = dotenv_values(".env")
PDF_PATH = "/Users/junyeongc/KASE/AI도입/2024_CJ제일제당_sustainability_report_ko.pdf"
TEMPLATE = "templates/식품_v3.xlsx"

# GPT가 '해상도 문제'로 표시한 row → indicator_number 추적
TARGET_ROWS = {
    8: "포장재 재생/재활용 (R8: 표/그래프 판독불가)",
    9: "LCA 도입 (R9: 실제 p.22에 명시)",
    17: "물 사용량 목표 (R17: 표 판독불가)",
    47: "재생에너지 사용률 (R47: 표 % 판독불가)",
    50: "Scope 1+2 온실가스 (R50: 표 수치 판독불가)",
}


def pick_indicators():
    tpl = load_template(TEMPLATE)
    e = tpl["식품-E"]
    selected = []
    for ind in e:
        if ind["row"] in TARGET_ROWS:
            ind["_label"] = TARGET_ROWS[ind["row"]]
            selected.append(ind)
    return selected


async def run_one(provider, pdf_bytes, indicators):
    user_prompt = build_user_prompt("식품-E (해상도 검증 부분집합)", indicators)
    t0 = time.time()
    try:
        results, usage = await provider.analyze_sheet(
            pdf_bytes=pdf_bytes,
            system_prompt=SYSTEM_PROMPT,
            user_prompt=user_prompt,
        )
        elapsed = time.time() - t0
        return {"results": results, "usage": usage, "elapsed": elapsed, "error": None}
    except Exception as e:
        return {"results": [], "usage": {}, "elapsed": time.time() - t0, "error": f"{type(e).__name__}: {e}"}


async def main():
    print("PDF 로드…")
    pdf_bytes = Path(PDF_PATH).read_bytes()
    print(f"  size: {len(pdf_bytes)/1024/1024:.1f}MB")

    indicators = pick_indicators()
    print(f"\n선정 지표 {len(indicators)}개:")
    for ind in indicators:
        print(f"  row {ind['row']} {ind['indicator_number']}: {ind['_label']}")
        print(f"    └ {ind['indicator'][:80]}")

    claude = ClaudeProvider(api_key=ENV["ANTHROPIC_API_KEY"], model="claude-sonnet-4-20250514")
    gemini = GeminiProvider(api_key=ENV["GOOGLE_API_KEY"], model="gemini-2.5-flash")

    print("\n=== Claude 분석 ===")
    cr = await run_one(claude, pdf_bytes, indicators)
    print(f"  elapsed: {cr['elapsed']:.1f}s, usage: {cr['usage']}, error: {cr['error']}")

    print("\n=== Gemini 분석 ===")
    gr = await run_one(gemini, pdf_bytes, indicators)
    print(f"  elapsed: {gr['elapsed']:.1f}s, usage: {gr['usage']}, error: {gr['error']}")

    # 결과 저장 + 요약 출력
    out = {"indicators": [{"row": i["row"], "num": i["indicator_number"], "label": i["_label"], "indicator": i["indicator"]} for i in indicators],
           "claude": cr, "gemini": gr}
    Path("_compare_result.json").write_text(json.dumps(out, ensure_ascii=False, indent=2))
    print("\n저장: _compare_result.json")

    # 핵심 비교: 각 지표별 e_content 끝부분에 "해상도/판독" 등 표현 있는지
    print("\n\n══════════════ 결과 요약 (지표별 e_content) ══════════════")
    for i, ind in enumerate(indicators):
        print(f"\n──── row {ind['row']} {ind['indicator_number']} {ind['_label']} ────")
        for name, res in [("Claude", cr), ("Gemini", gr)]:
            if res["error"]:
                print(f"  [{name}] ERROR: {res['error'][:200]}")
                continue
            if i < len(res["results"]):
                e = res["results"][i].get("e_content", "")
                f = res["results"][i].get("f_score", "?")
                e_short = e[:500].replace("\n", " | ")
                bad = any(k in e for k in ["해상도", "판독", "축소 화면", "캡처"])
                tag = " ⚠️ 해상도언급" if bad else ""
                print(f"  [{name}] score={f}{tag}\n    {e_short}{'...' if len(e)>500 else ''}")
            else:
                print(f"  [{name}] (응답 누락)")

asyncio.run(main())
