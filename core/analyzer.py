"""분석 오케스트레이터: 병렬 처리, 오류 복구, 진행 추적, 비용 계산."""
from __future__ import annotations

import asyncio
import json
import logging
import time
from pathlib import Path
from dataclasses import dataclass, field
from typing import Callable

from config import MAX_CONCURRENT, OUTPUT_DIR, calc_cost
from core.excel_handler import (
    load_template,
    prepare_output_file,
    write_sheet_results,
)
from core.prompt_builder import SYSTEM_PROMPT, build_user_prompt
from core.providers.base import LLMProvider

logger = logging.getLogger(__name__)


# ── 비용 추적 ──


@dataclass
class CostTracker:
    """분석 전체의 토큰 사용량과 비용을 누적 추적한다."""

    model: str = ""
    sheets: dict = field(default_factory=dict)  # sheet_name → cost_info
    total_input_tokens: int = 0
    total_output_tokens: int = 0
    total_input_cost: float = 0.0
    total_output_cost: float = 0.0
    total_cost: float = 0.0

    def add(self, sheet_name: str, input_tokens: int, output_tokens: int, elapsed_sec: float = 0) -> dict:
        cost = calc_cost(self.model, input_tokens, output_tokens)
        cost["elapsed_sec"] = elapsed_sec
        self.sheets[sheet_name] = cost
        self.total_input_tokens += input_tokens
        self.total_output_tokens += output_tokens
        self.total_input_cost += cost["input_cost"]
        self.total_output_cost += cost["output_cost"]
        self.total_cost += cost["total_cost"]
        return cost

    def summary(self) -> dict:
        total_elapsed = sum(s.get("elapsed_sec", 0) for s in self.sheets.values())
        return {
            "model": self.model,
            "total_input_tokens": self.total_input_tokens,
            "total_output_tokens": self.total_output_tokens,
            "total_tokens": self.total_input_tokens + self.total_output_tokens,
            "total_input_cost": self.total_input_cost,
            "total_output_cost": self.total_output_cost,
            "total_cost": self.total_cost,
            "total_elapsed_sec": total_elapsed,
            "sheets": self.sheets,
        }


# ── 단일 시트 분석 ──


async def _analyze_one_sheet(
    semaphore: asyncio.Semaphore,
    provider: LLMProvider,
    pdf_bytes: bytes,
    sheet_name: str,
    indicators: list[dict],
) -> tuple[str, list[dict] | str, dict]:
    """한 시트를 분석하고 (sheet_name, results|error, usage)를 반환한다."""
    async with semaphore:
        user_prompt = build_user_prompt(sheet_name, indicators)
        last_error = ""
        t_start = time.time()

        for attempt in range(2):
            try:
                results, usage = await provider.analyze_sheet(
                    pdf_bytes=pdf_bytes,
                    system_prompt=SYSTEM_PROMPT,
                    user_prompt=user_prompt,
                )
                usage["elapsed_sec"] = time.time() - t_start
                logger.info(
                    "[%s] %s 분석 완료 — %d개 지표 (%.1f초)",
                    provider.provider_name,
                    sheet_name,
                    len(results),
                    usage["elapsed_sec"],
                )
                return (sheet_name, results, usage)

            except json.JSONDecodeError as e:
                last_error = f"JSON 파싱 실패: {e}"
                logger.warning(
                    "[%s] %s JSON 파싱 실패 (시도 %d/2): %s",
                    provider.provider_name, sheet_name, attempt + 1, e,
                )
                if attempt == 0:
                    await asyncio.sleep(2)
                    continue

            except Exception as e:
                last_error = f"API 오류: {e}"
                logger.error(
                    "[%s] %s API 오류 (시도 %d/2): %s",
                    provider.provider_name, sheet_name, attempt + 1, e,
                )
                if attempt == 0:
                    await asyncio.sleep(3)
                    continue

        return (sheet_name, last_error, {"input_tokens": 0, "output_tokens": 0, "elapsed_sec": time.time() - t_start})


# ── 전체 분석 실행 ──


async def run_analysis(
    pdf_bytes: bytes,
    company_name: str,
    selected_sheets: list[str],
    provider: LLMProvider,
    template_path: str | Path = None,
    progress_callback: Callable[[str, str, float, dict | None], None] | None = None,
) -> tuple[Path, dict]:
    """한 PDF에 대해 선택된 시트들을 분석하고 결과 엑셀을 저장한다.

    Returns
    -------
    (결과 엑셀 경로, 비용 요약 dict)
    """
    template_data = load_template(template_path)
    output_path = prepare_output_file(template_path, OUTPUT_DIR, company_name)
    semaphore = asyncio.Semaphore(MAX_CONCURRENT)
    tracker = CostTracker(model=provider.model)

    # 처리할 시트와 지표 필터링
    sheet_tasks: list[tuple[str, list[dict]]] = []
    for sheet_name in selected_sheets:
        if sheet_name not in template_data:
            continue
        indicators = template_data[sheet_name]
        pending = [i for i in indicators if not i["has_existing_content"]]
        if pending:
            sheet_tasks.append((sheet_name, pending))

    total = len(sheet_tasks)
    if total == 0:
        logger.info("분석할 지표가 없습니다.")
        return output_path, tracker.summary()

    # 병렬 실행
    tasks = [
        _analyze_one_sheet(semaphore, provider, pdf_bytes, sname, inds)
        for sname, inds in sheet_tasks
    ]

    completed = 0
    for coro in asyncio.as_completed(tasks):
        sheet_name, result, usage = await coro
        completed += 1
        progress = completed / total

        # 비용·시간 누적
        sheet_cost = tracker.add(sheet_name, usage["input_tokens"], usage["output_tokens"], usage.get("elapsed_sec", 0))

        if isinstance(result, list):
            write_sheet_results(output_path, company_name, sheet_name, result)
            status = f"완료 ({len(result)}개 지표)"
            if progress_callback:
                progress_callback(sheet_name, status, progress, sheet_cost)
        else:
            error_results = _make_error_results(template_data[sheet_name], result)
            write_sheet_results(output_path, company_name, sheet_name, error_results)
            status = f"오류: {result}"
            if progress_callback:
                progress_callback(sheet_name, status, progress, sheet_cost)

    logger.info("분석 완료: %s → %s (총 비용: $%.4f)", company_name, output_path, tracker.total_cost)
    return output_path, tracker.summary()


def _make_error_results(indicators: list[dict], error_msg: str) -> list[dict]:
    """오류 발생 시 각 지표에 오류 메시지를 기입."""
    return [
        {
            "indicator_number": ind["indicator_number"],
            "e_content": f"(API 오류) {error_msg}",
            "f_score": 0,
            "g_review": f"※검토필요 — {error_msg}",
        }
        for ind in indicators
    ]


# ── 동기 래퍼 (Streamlit용) ──


def run_analysis_sync(
    pdf_bytes: bytes,
    company_name: str,
    selected_sheets: list[str],
    provider: LLMProvider,
    template_path: str | Path = None,
    progress_callback: Callable[[str, str, float, dict | None], None] | None = None,
) -> tuple[Path, dict]:
    """asyncio 이벤트 루프를 생성하여 분석을 실행하는 동기 래퍼."""
    loop = asyncio.new_event_loop()
    try:
        return loop.run_until_complete(
            run_analysis(pdf_bytes, company_name, selected_sheets, provider, template_path, progress_callback)
        )
    finally:
        loop.close()
