"""KASE ESG 자동 분석 시스템 — Streamlit 앱."""
from __future__ import annotations

import sys
from pathlib import Path

# 프로젝트 루트를 sys.path에 추가
sys.path.insert(0, str(Path(__file__).resolve().parent))

import logging
import streamlit as st
import pandas as pd

from config import (
    MODELS,
    OUTPUT_DIR,
    TEMPLATES_DIR,
    get_all_industries_with_templates,
    get_templates_for_industry,
    get_template_version,
    get_api_key,
    get_available_models,
)
from core.pdf_handler import extract_company_name, get_page_count
from core.excel_handler import load_template, get_sheet_names
from core.schemas import detect_schema, load_v4
from core.providers import create_provider
from core.analyzer import run_analysis_sync
from core.history_logger import log_analysis, is_gsheet_configured
from core.drive_sync import sync_templates_from_drive

# ── 로깅 설정 ──
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
)

# ── 앱 시작 시 Google Drive에서 템플릿 동기화 (세션당 1회) ──
if "drive_synced" not in st.session_state:
    synced = sync_templates_from_drive()
    st.session_state.drive_synced = True
    if synced:
        logging.info("Drive 템플릿 동기화 완료: %s", synced)

# ── 페이지 설정 ──
st.set_page_config(
    page_title="KASE ESG 자동 분석",
    page_icon="📊",
    layout="wide",
    menu_items={
        "Get Help": None,
        "Report a bug": None,
        "About": None,
    },
)

# ── 우측 상단 메뉴/Deploy 버튼 숨기기 ──
st.markdown(
    """
    <style>
    [data-testid="stMainMenu"] {display: none;}
    header [data-testid="stStatusWidget"] {display: none;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ── 접근 게이트 (선택) ──
# secrets에 app_password가 설정돼 있으면 비번을 요구하고, 없으면 누구나 접근(게이트 비활성).
# 공개 URL + 공용 LLM 키 사용 시 secrets에 app_password를 넣어 학회원만 쓰게 한다.
def _check_access() -> bool:
    try:
        pw = st.secrets.get("app_password")
    except Exception:
        pw = None
    if not pw:
        return True
    if st.session_state.get("_authed"):
        return True
    st.title("🔒 KASE ESG 분석")
    entered = st.text_input("접근 비밀번호를 입력하세요", type="password")
    if entered:
        if entered == str(pw):
            st.session_state["_authed"] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    return False


if not _check_access():
    st.stop()

# ══════════════════════════════════════════════
#  사이드바: 모델 선택 + API 상태
# ══════════════════════════════════════════════
with st.sidebar:
    st.title("⚙️ 설정")

    # 산업 분류 + 템플릿 선택
    available_industries = get_all_industries_with_templates()
    if not available_industries:
        st.error("❌ templates/ 폴더에 템플릿이 없습니다.")
        st.caption("파일명 규칙: `{산업분류}_{버전}.xlsx`\n\n예: `식품_v5.2.xlsx`")
        st.stop()

    selected_industry = st.selectbox(
        "산업 분류",
        available_industries,
        index=0,
    )

    industry_templates = get_templates_for_industry(selected_industry)
    template_labels = [f"{t.stem} ({get_template_version(t)})" for t in industry_templates]
    selected_template_idx = st.selectbox(
        "평가지표 템플릿",
        range(len(industry_templates)),
        format_func=lambda i: template_labels[i],
        index=0,
        help="templates/ 폴더에 파일을 추가하면 자동으로 표시됩니다.",
    )
    selected_template_path = industry_templates[selected_template_idx]
    selected_template_version = get_template_version(selected_template_path)

    # 모델 선택 (GPT 기본값)
    model_names = list(MODELS.keys())
    default_idx = model_names.index("GPT") if "GPT" in model_names else 0
    selected_model = st.selectbox(
        "LLM 모델 선택",
        model_names,
        index=default_idx,
        help="환경변수에 해당 모델의 API 키가 설정되어 있어야 합니다.",
    )

    # API 키 상태 표시
    st.markdown("---")
    st.subheader("API 키 상태")
    for name, cfg in MODELS.items():
        key = get_api_key(name)
        if key:
            st.success(f"✅ {name} — 설정됨 (`{cfg['env_key']}`)")
        else:
            st.error(f"❌ {name} — 미설정 (`{cfg['env_key']}`)")

    # 선택된 모델의 키 확인
    selected_api_key = get_api_key(selected_model)
    if not selected_api_key:
        st.warning(
            f"⚠️ **{selected_model}**의 API 키가 환경변수에 없습니다.\n\n"
            f"`{MODELS[selected_model]['env_key']}` 환경변수를 설정하세요."
        )

    # Google Sheets 이력 상태
    st.markdown("---")
    st.subheader("이력 저장")
    if is_gsheet_configured():
        st.success("✅ Google Sheets 연동됨")
    else:
        st.warning("⚠️ Google Sheets 미설정\n\n`credentials/gsheet_service_account.json` 필요")

    # 템플릿 정보
    st.markdown("---")
    st.subheader("템플릿 정보")
    st.info(f"📄 {selected_template_path.name}")
    st.caption(f"산업: {selected_industry} | 버전: {selected_template_version}")

    # 새 템플릿 업로드
    with st.expander("📤 새 템플릿 업로드"):
        uploaded_template = st.file_uploader(
            "템플릿 파일 (.xlsx)",
            type=["xlsx"],
            key="template_upload",
            help="파일명 규칙: {산업분류}_{버전}.xlsx\n예: 건설_v5.3.xlsx",
        )
        if uploaded_template:
            dest = TEMPLATES_DIR / uploaded_template.name
            dest.write_bytes(uploaded_template.read())
            st.success(f"✅ 저장 완료: {uploaded_template.name}")
            st.caption("페이지를 새로고침하면 목록에 표시됩니다.")
            st.rerun()

# ══════════════════════════════════════════════
#  메인 영역
# ══════════════════════════════════════════════
st.title("📊 KASE ESG 자동 분석 시스템")
st.caption("지속가능경영보고서(SR)를 KASE 평가지표 기준으로 자동 분석합니다.")

# ── 1. PDF 업로드 ──
st.header("1. SR PDF 업로드")
uploaded_files = st.file_uploader(
    "분석할 지속가능경영보고서(SR) PDF를 업로드하세요.",
    type=["pdf"],
    accept_multiple_files=True,
    help="복수 파일 업로드 가능. 순차적으로 하나씩 처리됩니다.",
)

if not uploaded_files:
    st.info("👆 PDF 파일을 업로드하면 분석을 시작할 수 있습니다.")
    st.stop()

# ── 2. 기업명 확인 ──
st.header("2. 기업명 확인")
st.caption("PDF 파일명에서 자동 추출됩니다. 필요시 수정하세요.")

company_names: dict[str, str] = {}
cols = st.columns(min(len(uploaded_files), 3))

for i, file in enumerate(uploaded_files):
    col = cols[i % len(cols)]
    with col:
        extracted = extract_company_name(file.name)
        company_names[file.name] = st.text_input(
            f"📄 {file.name}",
            value=extracted,
            key=f"company_{i}",
        )

# ── 3. 시트 선택 ──
st.header("3. 분석 시트 선택")

# 템플릿에서 시트 목록을 동적으로 읽기
SHEET_NAMES = get_sheet_names(selected_template_path)

col_btn1, col_btn2, _ = st.columns([1, 1, 4])
with col_btn1:
    select_all = st.button("전체 선택")
with col_btn2:
    deselect_all = st.button("전체 해제")

# 세션 상태로 체크박스 관리 (템플릿 변경 시 초기화)
if "sheet_checks" not in st.session_state or set(st.session_state.sheet_checks.keys()) != set(SHEET_NAMES):
    st.session_state.sheet_checks = {s: True for s in SHEET_NAMES}

if select_all:
    st.session_state.sheet_checks = {s: True for s in SHEET_NAMES}
if deselect_all:
    st.session_state.sheet_checks = {s: False for s in SHEET_NAMES}

sheet_cols = st.columns(min(len(SHEET_NAMES), 5))
for i, sname in enumerate(SHEET_NAMES):
    with sheet_cols[i % 5]:
        st.session_state.sheet_checks[sname] = st.checkbox(
            sname,
            value=st.session_state.sheet_checks.get(sname, True),
            key=f"sheet_{sname}",
        )

selected_sheets = [s for s, v in st.session_state.sheet_checks.items() if v]

if not selected_sheets:
    st.warning("⚠️ 최소 1개 시트를 선택하세요.")
    st.stop()

# 선택 요약 — 지표 수 + 배치 수 표시
_schema = detect_schema(selected_template_path)
_template_data = load_v4(selected_template_path) if _schema == "v4" else load_template(selected_template_path)
_BATCH_SIZE = 10  # core/analyzer.py 의 BATCH_SIZE 와 동일하게 유지

# 빠른 QA 지표 옵션을 캡션보다 먼저 구성 (캡션·expander 공용)
_ind_options: list[str] = []
_ind_label_to_key: dict[str, str] = {}
for _s in selected_sheets:
    for _ind in _template_data.get(_s, []):
        if _ind["has_existing_content"]:
            continue
        _key = f"{_s}|{_ind['indicator_number']}"
        _name = _ind.get("name") or _ind.get("indicator", "")
        _label = f"[{_s}] {_ind['indicator_number']} {str(_name)[:50]}"
        _ind_label_to_key[_label] = _key
        _ind_options.append(_label)

_total_indicators = len(_ind_options)  # 시트 전체 분석 대상 지표 수
# 직전 실행에서 빠른 QA로 고른 지표(현재 옵션에 유효한 것만) → 캡션이 실제 실행량을 반영
_qa_prev = [l for l in st.session_state.get("qa_indicators", []) if l in _ind_label_to_key]
_effective_indicators = len(_qa_prev) if _qa_prev else _total_indicators
_total_batches = -(-_effective_indicators // _BATCH_SIZE)  # ceil

if _qa_prev:
    st.info(
        f"📋 **{len(uploaded_files)}개 PDF** | **{len(selected_sheets)}개 시트** | "
        f"🔬 빠른 QA **{_effective_indicators}개 지표** (시트 전체 {_total_indicators}개 중) "
        f"→ **{_total_batches}개 배치** API 호출 예정"
    )
else:
    st.info(
        f"📋 **{len(uploaded_files)}개 PDF** | **{len(selected_sheets)}개 시트** | "
        f"**{_total_indicators}개 지표** → **{_total_batches}개 배치** API 호출 예정"
    )

# 🔬 빠른 QA — 특정 지표만 골라 실행 (비우면 시트 전체)
selected_indicators = None
with st.expander("🔬 빠른 QA — 특정 지표만 골라 실행 (비우면 위 시트 전체 실행)", expanded=bool(_qa_prev)):
    _picked = st.multiselect(
        "지표 선택 (검색 가능 — 일부만 골라 빠르게 QA)",
        _ind_options,
        key="qa_indicators",
        help="여기서 지표를 고르면 시트 전체 대신 선택한 지표만 분석합니다.",
    )
    if _picked:
        selected_indicators = [_ind_label_to_key[lbl] for lbl in _picked]
        st.success(f"⚡ 빠른 QA: 선택한 **{len(selected_indicators)}개 지표**만 실행됩니다.")

# ── 4. 분석 실행 ──
st.header("4. 분석 실행")

if not selected_api_key:
    st.error(f"❌ {selected_model}의 API 키가 없어 분석을 실행할 수 없습니다.")
    st.stop()

# 분석 상태 관리
if "stop_requested" not in st.session_state:
    st.session_state.stop_requested = False

start_clicked = st.button("▶️ 분석 시작", type="primary", width="stretch")

if start_clicked:
    st.session_state.stop_requested = False

    # 분석 중 UI: 중단 버튼 표시
    if st.button("⏹️ 분석 중단", type="secondary", width="stretch"):
        st.session_state.stop_requested = True

    st.markdown("---")

    model_cfg = MODELS[selected_model]
    provider = create_provider(
        provider_key=model_cfg["provider"],
        api_key=selected_api_key,
        model=model_cfg["default_model"],
    )

    result_paths: list[Path] = []
    all_cost_summaries: list[dict] = []
    stopped_early = False

    # PDF 순차 처리
    for pdf_idx, file in enumerate(uploaded_files):
        if st.session_state.stop_requested:
            stopped_early = True
            break

        company = company_names[file.name]
        pdf_bytes = file.read()
        file.seek(0)

        st.markdown(f"### 📄 PDF {pdf_idx + 1}/{len(uploaded_files)}: **{company}**")
        page_count = get_page_count(pdf_bytes)
        st.caption(f"페이지 수: {page_count}p | 모델: {selected_model} (`{model_cfg['default_model']}`)")

        progress_bar = st.progress(0.0)
        status_text = st.empty()
        log_container = st.empty()

        sheet_statuses: dict[str, str] = {}

        def progress_callback(sheet_name: str, status: str, progress: float, cost_info: dict | None = None):
            cost_str = ""
            if cost_info and cost_info.get("total_cost", 0) > 0:
                elapsed = cost_info.get("elapsed_sec", 0)
                m, s = divmod(int(elapsed), 60)
                time_str = f"{m}분 {s}초" if m else f"{s}초"
                cost_str = (
                    f" | {time_str}"
                    f" | {cost_info['input_tokens']:,} in + {cost_info['output_tokens']:,} out"
                    f" = ${cost_info['total_cost']:.4f}"
                )
            sheet_statuses[sheet_name] = f"{status}{cost_str}"
            progress_bar.progress(min(progress, 1.0))
            status_text.text(f"진행: {int(progress * 100)}% — 마지막 완료: {sheet_name}")
            lines = []
            for sn, st_val in sheet_statuses.items():
                icon = "❌" if "오류" in st_val else "✅"
                lines.append(f"{icon} **{sn}**: {st_val}")
            log_container.markdown("\n\n".join(lines))

        try:
            output_path, cost_summary = run_analysis_sync(
                pdf_bytes=pdf_bytes,
                company_name=company,
                selected_sheets=selected_sheets,
                provider=provider,
                template_path=selected_template_path,
                progress_callback=progress_callback,
                selected_indicators=selected_indicators,
            )
            result_paths.append(output_path)
            all_cost_summaries.append({"company": company, **cost_summary})
            progress_bar.progress(1.0)
            status_text.text(f"✅ {company} 분석 완료!")

            # Google Sheets 이력 저장
            saved = log_analysis(
                pdf_filename=file.name,
                company_name=company,
                industry=selected_industry,
                model_name=f"{selected_model} ({model_cfg['default_model']})",
                selected_sheets=selected_sheets,
                cost_summary=cost_summary,
                template_version=selected_template_version,
                result_filename=output_path.name,
            )
            if saved:
                st.caption("📋 Google Sheets 이력 저장 완료")

        except Exception as e:
            st.error(f"❌ {company} 분석 중 오류 발생: {e}")
            logging.exception("분석 오류")

    st.session_state.stop_requested = False

    if stopped_early:
        st.warning("⏹️ 사용자 요청에 의해 분석이 중단되었습니다. 완료된 결과까지만 저장됩니다.")

    # ── 5. 결과 ──
    if result_paths:
        st.header("5. 분석 결과")

        # ── 비용 요약 ──
        st.subheader("💰 비용 요약")
        grand_input = sum(c.get("total_input_tokens", 0) for c in all_cost_summaries)
        grand_output = sum(c.get("total_output_tokens", 0) for c in all_cost_summaries)
        grand_cost = sum(c.get("total_cost", 0) for c in all_cost_summaries)
        grand_elapsed = sum(c.get("total_elapsed_sec", 0) for c in all_cost_summaries)
        gm, gs = divmod(int(grand_elapsed), 60)

        cost_col1, cost_col2, cost_col3, cost_col4, cost_col5 = st.columns(5)
        with cost_col1:
            st.metric("총 소요시간", f"{gm}분 {gs}초")
        with cost_col2:
            st.metric("Input 토큰", f"{grand_input:,}")
        with cost_col3:
            st.metric("Output 토큰", f"{grand_output:,}")
        with cost_col4:
            st.metric("총 토큰", f"{grand_input + grand_output:,}")
        with cost_col5:
            st.metric("총 비용", f"${grand_cost:.4f}")

        # PDF별 비용 상세
        for cs in all_cost_summaries:
            el = cs.get("total_elapsed_sec", 0)
            em, es = divmod(int(el), 60)
            with st.expander(f"📊 {cs['company']} — ${cs.get('total_cost', 0):.4f} ({em}분 {es}초)"):
                sheet_costs = cs.get("sheets", {})
                if sheet_costs:
                    rows = []
                    for sn, sc in sheet_costs.items():
                        sec = sc.get("elapsed_sec", 0)
                        sm, ss = divmod(int(sec), 60)
                        rows.append({
                            "시트": sn,
                            "소요시간": f"{sm}분 {ss}초" if sm else f"{ss}초",
                            "Input 토큰": f"{sc['input_tokens']:,}",
                            "Output 토큰": f"{sc['output_tokens']:,}",
                            "Input 비용": f"${sc['input_cost']:.4f}",
                            "Output 비용": f"${sc['output_cost']:.4f}",
                            "합계": f"${sc['total_cost']:.4f}",
                        })
                    st.table(rows)

        st.markdown("---")

        # ── 다운로드 ──
        for path in result_paths:
            st.markdown(f"**📁 {path.name}**")
            col_dl, col_path = st.columns([1, 3])
            with col_dl:
                with open(path, "rb") as f:
                    st.download_button(
                        label="📥 다운로드",
                        data=f.read(),
                        file_name=path.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{path.name}",
                    )
            with col_path:
                st.code(str(path), language=None)

        # ── 결과 미리보기 ──
        st.subheader("결과 미리보기")
        for path in result_paths:
            with st.expander(f"📊 {path.stem}", expanded=True):
                tabs = st.tabs(selected_sheets)
                for tab, sheet_name in zip(tabs, selected_sheets):
                    with tab:
                        try:
                            _sc = detect_schema(path)
                            df = pd.read_excel(
                                str(path),
                                sheet_name=sheet_name,
                                header=0 if _sc == "v4" else 2,
                            )
                            _wanted = (
                                ["지표명", "AI AS-IS 내용", "AI 점수", "검토의견"]
                                if _sc == "v4"
                                else ["카테고리", "평가지표", "AI AS-IS 내용", "AI\n점수", "검토의견"]
                            )
                            display_cols = [c for c in df.columns if c in _wanted]
                            if display_cols:
                                st.dataframe(df[display_cols], width="stretch")
                            else:
                                st.dataframe(df, width="stretch")
                        except Exception as e:
                            st.warning(f"미리보기 로드 실패: {e}")
