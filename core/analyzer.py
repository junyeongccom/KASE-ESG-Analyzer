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


# ── 단일 배치 분석 ──


@dataclass
class BatchItem:
    """배치 내 개별 지표. 원래 소속 시트를 추적한다."""
    sheet_name: str
    indicator: dict


async def _analyze_batch(
    semaphore: asyncio.Semaphore,
    provider: LLMProvider,
    pdf_bytes: bytes,
    batch_id: int,
    batch_items: list[BatchItem],
) -> tuple[int, list[tuple[str, dict]] | str, dict]:
    """배치를 분석하고 (batch_id, [(sheet_name, result_item)...] | error, usage)를 반환한다."""
    async with semaphore:
        # 프롬프트용 지표 리스트 구성
        indicators = [item.indicator for item in batch_items]
        # 배치에 포함된 시트명 목록
        sheet_names = list(dict.fromkeys(item.sheet_name for item in batch_items))
        label = " + ".join(sheet_names)

        user_prompt = build_user_prompt(label, indicators)
        expected = len(batch_items)
        MAX_ATTEMPTS = 3
        last_error = ""
        best_results: list[dict] | None = None
        best_usage: dict | None = None
        t_start = time.time()

        for attempt in range(MAX_ATTEMPTS):
            try:
                results, usage = await provider.analyze_sheet(
                    pdf_bytes=pdf_bytes,
                    system_prompt=SYSTEM_PROMPT,
                    user_prompt=user_prompt,
                )

                # 가장 완성도 높은 응답을 보존 (재시도가 더 적게 반환할 수도 있음)
                if best_results is None or len(results) > len(best_results):
                    best_results = results
                    best_usage = usage

                if len(results) >= expected:
                    break

                if attempt < MAX_ATTEMPTS - 1:
                    logger.warning(
                        "[%s] 배치%d (%s) 부분 응답 %d/%d — 재시도 (%d/%d)",
                        provider.provider_name, batch_id, label,
                        len(results), expected, attempt + 1, MAX_ATTEMPTS,
                    )
                    await asyncio.sleep(2)
                    continue
                break

            except json.JSONDecodeError as e:
                last_error = f"JSON 파싱 실패: {e}"
                logger.warning(
                    "[%s] 배치%d JSON 파싱 실패 (시도 %d/%d): %s",
                    provider.provider_name, batch_id, attempt + 1, MAX_ATTEMPTS, e,
                )
                if attempt < MAX_ATTEMPTS - 1:
                    await asyncio.sleep(2)
                    continue

            except Exception as e:
                last_error = f"API 오류: {e}"
                logger.error(
                    "[%s] 배치%d API 오류 (시도 %d/%d): %s",
                    provider.provider_name, batch_id, attempt + 1, MAX_ATTEMPTS, e,
                )
                if attempt < MAX_ATTEMPTS - 1:
                    await asyncio.sleep(3)
                    continue

        # 모든 시도가 예외로 끝났고 부분 결과도 없음
        if best_results is None:
            return (batch_id, last_error, {"input_tokens": 0, "output_tokens": 0, "elapsed_sec": time.time() - t_start})

        results = best_results
        usage = best_usage
        usage["elapsed_sec"] = time.time() - t_start

        if len(results) >= expected:
            logger.info(
                "[%s] 배치%d (%s) 분석 완료 — %d개 지표 (%.1f초)",
                provider.provider_name, batch_id, label,
                len(results), usage["elapsed_sec"],
            )
        else:
            logger.warning(
                "[%s] 배치%d (%s) 재시도 후에도 부분 응답 — %d/%d개만 수집 (%.1f초)",
                provider.provider_name, batch_id, label,
                len(results), expected, usage["elapsed_sec"],
            )

        # 결과를 원래 시트에 매핑
        mapped: list[tuple[str, dict]] = []
        for i, item in enumerate(batch_items):
            if i < len(results):
                mapped.append((item.sheet_name, results[i]))
            else:
                mapped.append((item.sheet_name, {
                    "indicator_number": item.indicator["indicator_number"],
                    "e_content": "(응답 누락)",
                    "f_score": 0,
                    "g_review": "※검토필요 — LLM 응답에서 누락됨",
                }))
        return (batch_id, mapped, usage)


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

    # 전체 지표를 시트 경계 없이 15개씩 배치로 묶기
    BATCH_SIZE = 15
    all_items: list[BatchItem] = []

    for sheet_name in selected_sheets:
        if sheet_name not in template_data:
            continue
        indicators = template_data[sheet_name]
        for ind in indicators:
            if not ind["has_existing_content"]:
                all_items.append(BatchItem(sheet_name=sheet_name, indicator=ind))

    if not all_items:
        logger.info("분석할 지표가 없습니다.")
        return output_path, tracker.summary()

    # 25개씩 배치 생성
    batches: list[list[BatchItem]] = []
    for i in range(0, len(all_items), BATCH_SIZE):
        batches.append(all_items[i:i + BATCH_SIZE])

    total = len(batches)
    logger.info("총 %d개 지표 → %d개 배치 (배치당 최대 %d개)", len(all_items), total, BATCH_SIZE)

    # 병렬 실행
    tasks = [
        _analyze_batch(semaphore, provider, pdf_bytes, bid, batch)
        for bid, batch in enumerate(batches, 1)
    ]

    completed = 0
    for coro in asyncio.as_completed(tasks):
        batch_id, result, usage = await coro
        completed += 1
        progress = completed / total

        # 배치 라벨
        batch = batches[batch_id - 1]
        batch_sheets = list(dict.fromkeys(item.sheet_name for item in batch))
        label = " + ".join(batch_sheets)

        # 비용·시간 누적
        batch_cost = tracker.add(f"배치{batch_id}", usage["input_tokens"], usage["output_tokens"], usage.get("elapsed_sec", 0))

        if isinstance(result, list):
            # 결과를 시트별로 분류하여 기입
            by_sheet: dict[str, list[dict]] = {}
            for sheet_name, item_result in result:
                by_sheet.setdefault(sheet_name, []).append(item_result)
            for sheet_name, sheet_results in by_sheet.items():
                write_sheet_results(output_path, company_name, sheet_name, sheet_results)
            status = f"배치{batch_id} 완료 ({label}, {len(result)}개)"
            if progress_callback:
                progress_callback(label, status, progress, batch_cost)
        else:
            # 오류 시 배치 내 모든 지표에 오류 기입
            by_sheet: dict[str, list[dict]] = {}
            for item in batch:
                err = {
                    "indicator_number": item.indicator["indicator_number"],
                    "e_content": f"(API 오류) {result}",
                    "f_score": 0,
                    "g_review": f"※검토필요 — {result}",
                }
                by_sheet.setdefault(item.sheet_name, []).append(err)
            for sheet_name, sheet_results in by_sheet.items():
                write_sheet_results(output_path, company_name, sheet_name, sheet_results)
            status = f"배치{batch_id} 오류: {result}"
            if progress_callback:
                progress_callback(label, status, progress, batch_cost)

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
