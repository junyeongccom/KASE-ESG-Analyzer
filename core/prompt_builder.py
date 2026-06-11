"""시스템 프롬프트 및 유저 프롬프트 생성."""
from __future__ import annotations


SYSTEM_PROMPT = """\
당신은 ESG 지속가능경영보고서(SR) 분석 전문가입니다.
첨부된 SR PDF에서 아래 평가지표 각각에 대해 관련 내용을 찾아 분석하세요.

## 분석 규칙

### E열 — AS-IS 내용 (SR 원문 발췌)
- SR 원문을 **그대로 발췌**하세요. 요약하거나 해석하지 마세요.
- 반드시 **출처 페이지를 명시**하세요: [SR p.XX] 형식
- 관련 내용이 여러 페이지에 걸쳐 있으면 **모두 기입**하세요.
- 관련 내용이 없으면: "(공시확인불가)" + 왜 없다고 판단했는지 한 줄 메모
- 표(테이블)에 있는 데이터도 찾아서 포함하세요.

### F열 — 점수 산정
- **배점기준**을 정확히 읽고 아래 원칙에 따라 점수를 산정하세요:
  - 필수 조건 충족 → 기본 점수 부여
  - 가산 조건 각각 충족 여부 확인 → 추가 점수 합산
  - **만점을 초과할 수 없습니다.**
  - 판단이 불확실하면 **보수적으로(낮은 점수)** 산정하세요.
- 숫자만 기입하세요 (예: 1.5)

### G열 — 검토의견
아래 4가지 중 하나를 기입하세요:
- 정상 처리: "분석 완료"
- 공시 없음: "(공시확인불가)"
- 판단 불확실: "※검토필요 — [구체적 이유]"
- 부분 공시: "부분 공시 — [충족된 부분]과 [미충족 부분] 설명"

## 출력 형식
**반드시** 아래 JSON 형식으로만 응답하세요. 다른 텍스트를 추가하지 마세요.

```json
{
  "results": [
    {
      "indicator_number": "(1)",
      "e_content": "SR에서 찾은 원문 내용 [SR p.XX]\\n추가 발견 내용 [SR p.YY]",
      "f_score": 1.5,
      "g_review": "분석 완료"
    }
  ]
}
```

중요:
- indicator_number는 지표의 번호입니다 (예: "(1)", "(2)", "(13)")
- f_score는 반드시 숫자(정수 또는 소수)여야 합니다
- 모든 지표에 대해 빠짐없이 응답하세요
"""


def build_user_prompt(sheet_name: str, indicators: list[dict]) -> str:
    """유저 프롬프트를 생성한다."""
    lines = [
        f'아래는 "{sheet_name}" 영역의 평가지표 목록입니다.',
        "첨부된 SR PDF에서 각 지표를 분석해주세요.",
        "",
        "=" * 50,
        f"  평가 영역: {sheet_name}",
        f"  총 지표 수: {len(indicators)}개",
        "=" * 50,
        "",
    ]

    for item in indicators:
        lines.append(f"[지표 {item['indicator_number']}]")
        if item.get("category"):
            lines.append(f"카테고리: {item['category']}")
        lines.append(f"평가지표: {item['indicator']}")
        lines.append(f"배점기준: {item['criteria']}")
        if item.get("guideline"):
            lines.append(f"평가지침: {item['guideline']}")
        lines.append(f"만점: {item['max_score']}")
        lines.append("-" * 40)
        lines.append("")

    return "\n".join(lines)


# ── v4 (YN 트리 + 메타데이터) 프롬프트 ──

# 평가 시스템 버전 (시스템 프롬프트 + 로직). 변경 시 ~/KASE/SYSTEM_VERSION_LOG.xlsx 와 항상 동기화.
# 1.3.0: 결정론 11지표 구조적 채점기 도입(LLM 점수 → 코드 PDF교차검증 오버라이드). core/analyzer._apply_structural_override.
SYSTEM_VERSION = "1.3.0"

SYSTEM_PROMPT_V4 = """당신은 ESG 지속가능경영보고서(SR) 평가 전문가입니다. 첨부된 SR PDF만을 근거로 각 지표를 평가하세요.

## 절대 규칙 (환각 방지)
- SR에 **그대로(verbatim) 인용 가능한 문장/수치**가 있을 때만 '공시됨'으로 인정한다.
- ESG 일반지식·업계평균·추측으로 빈칸을 메우지 마라. 근거 없으면 미충족으로 본다.
- **표 하단 주석과 작은 글씨 각주까지** 반드시 확인하라 (정의·수치가 각주에 있는 경우가 많다).
- 모든 인정 근거에는 [SR p.XX] 페이지 + 원문 인용을 붙여라. 근거 못 찾으면 "(공시확인불가)".

## 점수 = YN 의사결정 트리 (정확히 이 순서)
 1) 필수조건 미충족 → '필수=N 점수'
 2) 필수조건 충족, 가산조건1 미충족 → 'Y→N 점수'
 3) 필수+가산1 충족, 가산조건2 미충족 → 'Y→Y→N 점수'
 4) 필수+가산1+가산2 모두 충족 → '만점'
- 가산조건 칸이 '(없음)'/'(없음 → 바로 만점)'이면 그 분기는 없으며 직전 단계 충족이 최고점.
- ⚠️ "데이터를 찾음" ≠ "필수조건 충족". 조건 문구가 요구하는 바를 정확히 충족해야 한다.

## 출력 (JSON만)
- **언어 규칙(중요)**: e_content(AS-IS 근거)는 보고서에 적힌 **원문 언어 그대로 축자 인용**하라 — 영문 보고서면 영어 원문 그대로, 절대 한국어로 번역·의역하지 마라(번역하면 원문 대조 시 출처검증에 실패한다). 반면 g_review(검토의견)는 **반드시 한국어**로, 왜 그 점수인지 한눈에 이해되게 설명하라.
- e_content는 **핵심 근거 원문·수치만 간결히** (최대 3줄). g_review는 1~2문장.
- indicator_number는 각 지표 머리의 번호를 그대로 기입 (예: "(1)", "(13)").
{"results":[{"indicator_number":"(1)","e_content":"근거 원문 [SR p.XX]","f_score":<숫자>,"g_review":"어느 분기로 몇 점인지 근거"}]}
"""


def _fmt_indicator_v4(ind: dict) -> str:
    def _sc(v):
        return v if (v not in (None, "") and str(v).strip()) else 0
    L = [f"[지표 {ind['indicator_number']}] {ind['name']}"]
    L.append(f"  · 필수조건: {ind.get('must')}  (미충족 시 {_sc(ind.get('s_mustN'))}점)")
    add1 = str(ind.get("add1")).strip() if ind.get("add1") is not None else ""
    if add1 and "없음" not in add1:
        L.append(f"  · 가산조건1: {add1}  (필수만 충족 시 {_sc(ind.get('s_YN'))}점)")
        add2 = str(ind.get("add2")).strip() if ind.get("add2") is not None else ""
        if add2 and "없음" not in add2:
            L.append(f"  · 가산조건2: {add2}  (필수+가산1 충족 시 {_sc(ind.get('s_YYN'))}점)")
        L.append(f"  · 위 조건 모두 충족 → {_sc(ind.get('s_full'))}점 (만점)")
    else:
        L.append(f"  · (가산조건 없음) 필수조건 충족 시 바로 만점 {_sc(ind.get('s_full'))}점")
    if ind.get("formula"):
        L.append(f"  · 참고산식: {ind['formula']}")
    if ind.get("guide"):
        L.append(f"  · 평가지침: {ind['guide']}")
    m = ind.get("meta") or {}
    if any(m.values()):
        L += [
            "  [메타데이터]",
            f"   - 지표 의도(Intent): {m.get('intent','')}",
            f"   - 용어 사전/동의어: {m.get('terms','')}",
            f"   - 인정 형태(Accept): {m.get('accept','')}",
            f"   - 불인정 반례(Reject): {m.get('reject','')}",
            f"   - 검색 키워드: {m.get('keywords','')}",
            f"   - 출처 요구(Citation): {m.get('citation','')}",
        ]
    return "\n".join(L)


def build_user_prompt_v4(sheet_name: str, indicators: list[dict]) -> str:
    """v4 유저 프롬프트(YN 트리 + 메타데이터)를 생성한다."""
    head = [
        f'아래는 "{sheet_name}" 영역의 평가지표입니다. 첨부 SR에서 각 지표를 평가하고 id를 그대로 echo 하세요.',
        f"총 지표 수: {len(indicators)}개",
        "",
    ]
    return "\n".join(head) + "\n" + "\n\n".join(_fmt_indicator_v4(i) for i in indicators)
