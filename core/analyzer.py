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
    write_results_v4,
)
from core.prompt_builder import (
    SYSTEM_PROMPT,
    build_user_prompt,
    SYSTEM_PROMPT_V4,
    build_user_prompt_v4,
)
from core.schemas import detect_schema, load_v4
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
    system_prompt: str,
    build_user_prompt_fn,
    max_output_tokens: int = 16000,
) -> tuple[int, list[tuple[str, dict]] | str, dict]:
    """배치를 분석하고 (batch_id, [(sheet_name, result_item)...] | error, usage)를 반환한다."""
    async with semaphore:
        # 프롬프트용 지표 리스트 구성
        indicators = [item.indicator for item in batch_items]
        # 배치에 포함된 시트명 목록
        sheet_names = list(dict.fromkeys(item.sheet_name for item in batch_items))
        label = " + ".join(sheet_names)

        user_prompt = build_user_prompt_fn(label, indicators)
        last_error = ""
        t_start = time.time()
        logger.info("│ 배치%d 시작 │ %s │ %d개 지표", batch_id, label, len(batch_items))

        # 시도 간 '최선(가장 많이 반환된)' 결과를 누적 유지 → 재시도가 더 나쁜 결과를 덮어쓰지 않게
        best_results = None
        best_usage = None
        for attempt in range(2):
            try:
                logger.info("│ 배치%d ②LLM호출 │ %s │ 시도 %d/2", batch_id, provider.provider_name, attempt + 1)
                results, usage = await provider.analyze_sheet(
                    pdf_bytes=pdf_bytes,
                    system_prompt=system_prompt,
                    user_prompt=user_prompt,
                    max_output_tokens=max_output_tokens,
                )
                usage["elapsed_sec"] = time.time() - t_start
                logger.info(
                    "│ 배치%d ③응답 │ 파싱 %d건 │ in=%d out=%d │ %.1f초",
                    batch_id, len(results), usage["input_tokens"], usage["output_tokens"], usage["elapsed_sec"],
                )
                if best_results is None or len(results) > len(best_results):
                    best_results, best_usage = results, usage
            except json.JSONDecodeError as e:
                last_error = f"JSON 파싱 실패: {e}"
                logger.warning("[%s] 배치%d JSON 파싱 실패 (시도 %d/2): %s",
                               provider.provider_name, batch_id, attempt + 1, e)
            except Exception as e:
                last_error = f"API 오류: {e}"
                logger.error("[%s] 배치%d API 오류 (시도 %d/2): %s",
                             provider.provider_name, batch_id, attempt + 1, e)
            # 완전하면 종료, 불완전하면 1회 재시도
            if best_results is not None and len(best_results) >= len(batch_items):
                break
            if attempt == 0:
                if best_results is not None:
                    logger.warning("│ 배치%d ⚠불완전응답 │ 최선 %d/%d → 재시도",
                                   batch_id, len(best_results), len(batch_items))
                await asyncio.sleep(2)

        if best_results is None:
            return (batch_id, last_error, {"input_tokens": 0, "output_tokens": 0, "elapsed_sec": time.time() - t_start})

        if len(best_results) < len(batch_items):
            logger.warning("│ 배치%d ⚠최종 불완전 │ %d/%d (나머지 누락 처리)",
                           batch_id, len(best_results), len(batch_items))
        # 최선 결과를 원래 시트에 매핑 (부족분은 누락 처리)
        mapped: list[tuple[str, dict]] = []
        for i, item in enumerate(batch_items):
            if i < len(best_results):
                mapped.append((item.sheet_name, best_results[i]))
            else:
                mapped.append((item.sheet_name, {
                    "indicator_number": item.indicator["indicator_number"],
                    "e_content": "(응답 누락)",
                    "f_score": 0,
                    "g_review": "※검토필요 — LLM 응답에서 누락됨",
                }))
        return (batch_id, mapped, best_usage)


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
    schema = detect_schema(template_path)
    if schema == "v4":
        template_data = load_v4(template_path)
        system_prompt = SYSTEM_PROMPT_V4
        build_fn = build_user_prompt_v4
        write_fn = write_results_v4
        BATCH_SIZE = 10
        max_out = 32000
    else:
        template_data = load_template(template_path)
        system_prompt = SYSTEM_PROMPT
        build_fn = build_user_prompt
        write_fn = write_sheet_results
        BATCH_SIZE = 10
        max_out = 32000
    logger.info("══════ 분석 시작 │ schema=%s │ 시트=%s │ 모델=%s ══════",
                schema, selected_sheets, provider.model)
    logger.info("│ ①로드 │ 배치크기 %d │ 출력캡 %d", BATCH_SIZE, max_out)

    output_path = prepare_output_file(template_path, OUTPUT_DIR, company_name)
    semaphore = asyncio.Semaphore(MAX_CONCURRENT)
    tracker = CostTracker(model=provider.model)

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
    logger.info("│ ①로드 완료 │ 분석대상 %d지표 → %d배치", len(all_items), total)

    # 병렬 실행
    tasks = [
        _analyze_batch(semaphore, provider, pdf_bytes, bid, batch,
                       system_prompt, build_fn, max_out)
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
                write_fn(output_path, company_name, sheet_name, sheet_results)
                logger.info("│ 배치%d ④기록 │ %s │ %d건 → 엑셀", batch_id, sheet_name, len(sheet_results))
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
                write_fn(output_path, company_name, sheet_name, sheet_results)
            status = f"배치{batch_id} 오류: {result}"
            if progress_callback:
                progress_callback(label, status, progress, batch_cost)

    logger.info("══════ 분석 완료 │ %s │ %d배치 │ 비용 $%.4f │ %s ══════",
                company_name, total, tracker.total_cost, output_path.name)
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
