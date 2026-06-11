"""
@file structural_scorer.py
@description 14개 stably-wrong 지표용 구조적 채점기 — 유형별 (LLM 추출 + 코드 PDF검증 + 코드 채점). 결정론 지표의 날조·자기모순 제거.
유형: year_count / threshold / trend / computed_trend / presence / completeness / rubric
"""
import json, re, unicodedata, threading
from pathlib import Path
from dotenv import dotenv_values
from google import genai
from google.genai import types
import fitz

BASE = Path("/Users/junyeongc/KASE/AI도입/KASE_ESG_Analyzer")
PDFDIR = Path("/Users/junyeongc/KASE/AI도입/03_원본보고서_PDF")
ENV = dotenv_values(str(BASE / ".env"))
client = genai.Client(api_key=ENV["GOOGLE_API_KEY"], http_options=types.HttpOptions(timeout=120000))  # 120s 타임아웃 — 네트워크 끊김 시 무한 hang 방지

# 14개 지표 채점 설정 (정답=pro/실측)
INDS = [
    {"id": "#1 먼지개선", "pdf": "CJ제일", "type": "trend", "metric": "먼지(PM) 대기오염물질 배출량(톤)", "kw": ["먼지"], "truth": 1},
    {"id": "#2 SOx개선", "pdf": "오뚜기", "type": "trend", "metric": "황산화물(SOx) 배출량(톤), 오뚜기 단독 기준", "kw": ["SOx", "황산화물"], "truth": None},
    {"id": "#3 물스트레스조달", "pdf": "nestle", "type": "presence_code",
     "metric": "물 스트레스 지역에서 조달된 식품 원료 비율 (SASB FB-PF-440a.1) 정량 공시",
     "core": ["물 스트레스", "물스트레스", "water stress", "water-stress", "water-stressed", "수자원 스트레스"],
     "src": ["조달", "sourc", "공급되는", "procure", "구매 비율"],  # 소싱(취수·사용 reject)
     "window": 200, "truth": 0},  # 전방 200자 내 소싱어 + 실제 소수% 값 → 2 (취수량≠조달, 인덱스 라벨 '-'≠공시)
    {"id": "#4 임시직", "pdf": "CJ제일", "type": "completeness", "metric": "임시직 수 또는 비율",
     "disclose_kw": ["기간의 정함이 있는", "기간제", "계약직", "임시직", "비정규직", "단시간 근로",
                     "temporary", "fixed-term", "contingent", "non-permanent"],  # 6개사 SR 실측 임시직 공시 표현
     "exclusion": ["일용직", "노무파견", "파견", "도급", "하도급", "subcontract", "outsourc"],  # 6개사 SR 실측 카브아웃(별도 고용형태)
     "truth": 1.5},
    {"id": "#5 정규직80", "pdf": "CJ제일", "type": "threshold", "num": "정규직(기간의 정함이 없는 근로자) 총 수", "den": "전체 임직원 총 수", "op": ">=", "val": 80, "truth": 0},
    {"id": "#6 정보사고", "pdf": "오뚜기", "type": "rubric", "criterion": "정보보안 '사고 발생 후' 이를 처리하는 대응 체계/절차(Incident Response Plan, 사고대응팀, 신고·복구 절차). 사전 진단·모니터링·예방만 있으면 불충족.", "truth": 0},
    {"id": "#7 제3자인증", "pdf": "unilever", "type": "rubric", "criterion": "제품·서비스의 '안전성(safety)'을 제3자 인증기관에서 검증받은 내역(ISO 등 품질·안전 인증). cruelty-free/동물실험 안 함 같은 윤리 인증은 '안전성' 검증으로 불충족(보수적).", "truth": 0},
    {"id": "#8 LTIFR", "pdf": "풀무원", "type": "year_count", "n": 3, "metric": "근로손실재해율(LTIR/LTIFR), 자사 임직원", "kw": ["LTIR", "근로손실재해율"], "truth": 0},
    {"id": "#9 산재율", "pdf": "풀무원", "type": "year_count", "n": 3, "metric": "산업재해율(%), 자사 임직원", "kw": ["산업재해율"], "truth": 0},
    {"id": "#10 Scope3", "pdf": "오뚜기", "type": "year_count", "n": 3, "metric": "Scope 3 온실가스 배출량", "kw": ["Scope3", "Scope 3"], "truth": 0},
    {"id": "#11 집약도", "pdf": "unilever", "type": "year_count", "n": 3, "metric": "온실가스 배출 집약도(GHG emissions intensity)", "kw": ["intensity"], "truth": 0},
    {"id": "#12 지정폐기물", "pdf": "CJ제일", "type": "computed_trend", "formula": "(지정폐기물 매립 + 지정폐기물 소각) / 총 폐기물 배출량", "inputs": ["지정폐기물 매립량(톤)", "지정폐기물 소각량(톤)", "총 폐기물 배출량(톤)"], "direction": "감소", "truth": 2},
    {"id": "#13 협력사산재율", "pdf": "풀무원", "type": "year_count", "n": 3, "metric": "협력기업(협력업체) 산업재해율(%)", "kw": ["산업재해율", "협력기업"], "truth": 0},
    {"id": "#14 불공정거래", "pdf": "풀무원", "type": "year_count", "n": 3, "metric": "불공정거래/부정경쟁(공정경쟁) 위반·제재 건수", "kw": ["공정경쟁", "경쟁저해"], "truth": 0},
]

_PDFCACHE = {}
_BYTES = {}
_FITZ_LOCK = threading.Lock()  # PyMuPDF는 스레드 비안전 — fitz 호출 직렬화 (동시 파싱 시 segfault 방지)
def pdf_bytes(sub):
    if sub not in _BYTES:
        for f in PDFDIR.glob("*.pdf"):
            if unicodedata.normalize("NFC", sub) in unicodedata.normalize("NFC", f.name):
                _BYTES[sub] = (f.read_bytes(), f)
                break
        else:
            raise FileNotFoundError(sub)
    return _BYTES[sub]
def pdf_text(sub):
    if sub not in _PDFCACHE:
        with _FITZ_LOCK:                              # fitz.open + get_text 전체를 직렬화
            if sub not in _PDFCACHE:                  # double-check (락 대기 중 타 스레드가 채웠을 수)
                data, f = pdf_bytes(sub)
                doc = fitz.open(f) if f is not None else fitz.open(stream=data, filetype="pdf")  # f=None → 앱에서 bytes 직접 주입
                _PDFCACHE[sub] = " ".join(p.get_text() for p in doc)
    return _PDFCACHE[sub]


def warm(pdf_data):
    """앱 통합용 — 동시 실행 전 단일스레드로 PDF 텍스트 캐시 예열 (fitz 동시파싱 회피)."""
    key = "app_" + hashlib.md5(pdf_data).hexdigest()[:12]
    if key not in _BYTES:
        _BYTES[key] = (pdf_data, None)
    pdf_text(key)
    return key


def ask(sub, prompt):
    data, _ = pdf_bytes(sub)
    r = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[types.Content(parts=[types.Part.from_bytes(data=data, mime_type="application/pdf"),
                                        types.Part.from_text(text=prompt)])],
        config=types.GenerateContentConfig(temperature=0, response_mime_type="application/json",
                                           thinking_config=types.ThinkingConfig(thinking_budget=0)))
    out = json.loads(r.text)
    if isinstance(out, list):  # 모델이 [{...}]로 감싸는 경우 → dict로 정규화
        out = next((x for x in out if isinstance(x, dict)), {})
    return out


def verify_years(sub, kws, claimed):
    txt = pdf_text(sub); real = set()
    for kw in kws:
        for m in re.finditer(re.escape(kw), txt):
            real |= {int(y) for y in re.findall(r"\b(20[12]\d)\b", txt[max(0, m.start() - 500):m.start() + 200])}
    return sorted(int(y) for y in claimed if int(y) in real)


def score(ind):
    """유형별 채점 → (점수, 근거dict)."""
    t = ind["type"]
    if t == "year_count":
        ext = ask(ind["pdf"], f"데이터 추출도구. 점수X. 지표 '{ind['metric']}'의 숫자가 실제 공시된 연도만 나열.\n없는 연도 생성 금지. JSON: {{\"disclosed_years\":[정수], \"values\":{{}}}}")
        ver = verify_years(ind["pdf"], ind["kw"], ext.get("disclosed_years", []))
        ok = len(ver) >= ind["n"] and (ver[-1] - ver[-ind["n"]] == ind["n"] - 1)
        return (2 if ok else 0), {"공시연도": ver, "n>=%d" % ind["n"]: ok}
    if t == "threshold":
        ext = ask(ind["pdf"], f"데이터 추출도구. 점수X. '{ind['num']}'와 '{ind['den']}'의 당해연도 실제 수치만. JSON: {{\"num\":숫자,\"den\":숫자}}")
        num, den = ext.get("num"), ext.get("den")
        try:
            ratio = float(num) / float(den) * 100
            ok = ratio >= ind["val"] if ind["op"] == ">=" else ratio <= ind["val"]
            return (2 if ok else 0), {"비율%": round(ratio, 1), "기준": f"{ind['op']}{ind['val']}", "충족": ok}
        except (TypeError, ValueError, ZeroDivisionError):
            return 0, {"추출실패": [num, den]}
    if t in ("trend", "computed_trend"):
        if t == "trend":
            ext = ask(ind["pdf"], f"데이터 추출도구. 점수X. '{ind['metric']}'의 연도별 실제 값만. 없는 연도 생성 금지. JSON: {{\"by_year\":{{\"연도\":값}}}}")
            by = {int(k): float(re.sub(r'[,\s]', '', str(v))) for k, v in ext.get("by_year", {}).items() if str(v).replace(',', '').replace('.', '').strip().isdigit()}
        else:
            ext = ask(ind["pdf"], f"데이터 추출도구. 점수X. 다음 항목들의 연도별 실제 값만: {ind['inputs']}. 산식={ind['formula']}. JSON: {{\"매립\":{{\"연도\":값}},\"소각\":{{\"연도\":값}},\"총\":{{\"연도\":값}}}}")
            mae, so, tot = ext.get("매립", {}), ext.get("소각", {}), ext.get("총", {})
            ys = sorted(set(mae) & set(so) & set(tot), key=lambda x: int(x))
            by = {}
            for y in ys:
                try:
                    by[int(y)] = (float(re.sub(r'[,\s]', '', str(mae[y]))) + float(re.sub(r'[,\s]', '', str(so[y])))) / float(re.sub(r'[,\s]', '', str(tot[y]))) * 100
                except (ValueError, ZeroDivisionError):
                    pass
        yrs = sorted(by)
        if len(yrs) < 2:
            return 0, {"by_year": by, "사유": "연도 부족"}
        latest_down = by[yrs[-1]] < by[yrs[-2]]                  # 필수: 당해 < 전년
        consec3 = len(yrs) >= 3 and all(by[yrs[i]] < by[yrs[i - 1]] for i in range(len(yrs) - 2, len(yrs)))  # 3개년 연속감소
        s = 2 if (latest_down and consec3) else (1 if latest_down else 0)
        return s, {"by_year": {k: round(v, 3) for k, v in by.items()}, "당해감소": latest_down, "연속감소": consec3}
    if t == "presence_code":
        # 정량 공시 여부를 순수 코드로 — 핵심어(물스트레스) 전방 window 내 소싱어 + 실제 소수% 값 동시존재 (LLM 판단 폐기)
        # 취수량/사용량(reject)·SASB 인덱스 라벨('-')은 소수% 값이 없어 자동 제외
        txt = pdf_text(ind["pdf"]); w = ind.get("window", 200)
        for core in ind["core"]:
            for m in re.finditer(re.escape(core), txt):
                fwd = txt[m.start():m.start() + w]
                mv = re.search(r"\d+\.\d+\s*%", fwd)
                if mv and any(s in fwd for s in ind["src"]):
                    return 2, {"공시": True, "값": mv.group(), "근거": fwd.replace("\n", " ")[:90]}
        return 0, {"공시": False}
    if t == "presence":
        ext = ask(ind["pdf"], f"데이터 추출도구. 지표 '{ind['metric']}'가 보고서에 **정량 공시**됐는지. 단, {ind['reject']} 같은 다른 지표를 이걸로 착각 금지. JSON: {{\"disclosed\":true/false, \"evidence\":\"원문\", \"is_exact_metric\":true/false}}")
        ok = bool(ext.get("disclosed")) and bool(ext.get("is_exact_metric"))
        return (2 if ok else 0), {"공시": ext.get("disclosed"), "정확지표": ext.get("is_exact_metric")}
    if t == "completeness":
        # 완전성: 순수 코드 결정론 (LLM 판단 폐기 — flash가 제외각주를 보고도 all_types=True로 만점주던 오류 제거)
        txt = pdf_text(ind["pdf"])
        hit = [kw for kw in ind.get("disclose_kw", []) if kw in txt]   # 임시직 공시 표현 존재
        excl = [kw for kw in ind.get("exclusion", []) if kw in txt]    # 도급·파견 등 별도 고용형태 카브아웃 → 완전성 미달
        disclosed = bool(hit)
        alltypes = disclosed and not excl
        s = 2 if alltypes else (1.5 if disclosed else 0)
        return s, {"공시": disclosed, "근거표현": hit, "모든유형": alltypes, "제외표현": excl}
    if t == "rubric":
        ext = ask(ind["pdf"], f"평가 기준: {ind['criterion']}\n이 기준을 **엄격히** 적용해, 보고서가 충족하는지 판정. JSON: {{\"meets\":true/false, \"evidence\":\"원문\", \"reason\":\"왜\"}}")
        return (2 if bool(ext.get("meets")) else 0), {"충족": ext.get("meets"), "사유": str(ext.get("reason", ""))[:80]}
    return None, {"미지원유형": t}


# ── 실제 앱(core/analyzer) 통합 어댑터 — 결정론 11지표만 구조적 채점으로 점수 오버라이드 (#2 정책·#6·#7 루브릭 제외) ──
import hashlib

# 라우팅 대상: 구조ID → 템플릿(식품_v4) 정확 지표명
_ROUTE_NAMES = {
    "#1 먼지개선": "(먼지) 최근 3개년 대기오염물질 배출 실적이 개선되었다.",
    "#3 물스트레스조달": "물 스트레스 지수가 높거나 매우 높은 지역에서 조달된 식품 원료의 비율을 정량 공개하고 있다.",
    "#4 임시직": "임시직 수 또는 비율을 공개하고 있다.",
    "#5 정규직80": "전체 임직원 대비 정규직 임직원 비율이 80% 이상이다.",
    "#8 LTIFR": "최근 3개년 근로손실재해율(LTIFR)을 공개하고 있다.",
    "#9 산재율": "최근 3개년 산업재해율(%)을 공개하고 있다.",
    "#10 Scope3": "최근 3개년 온실가스 Scope 3 배출량을 공개하고 있다.",
    "#11 집약도": "최근 3개년 온실가스 배출 집약도를 공개하고 있다.",
    "#12 지정폐기물": "최근 3개년 총 폐기물 배츌량 대비 지정폐기물의 (매립+소각) 비율이 개선되었다.",
    "#13 협력사산재율": "최근 3개년 협력업체 산업재해율(%)을 공개하고 있다.",
    "#14 불공정거래": "최근 3개년도 불공정거래/부정경쟁을 하여 제재받은 내역을 공개한다.",
}


def _norm_name(s):
    """지표명 매칭용 정규화 — 공백·NBSP·마침표·쉼표 제거."""
    return str(s).replace(" ", "").replace(" ", "").replace(".", "").replace(",", "").strip()


_CFG_BY_ID = {i["id"]: i for i in INDS}
ROUTED = {_norm_name(nm): rid for rid, nm in _ROUTE_NAMES.items()}  # 정규화 지표명 → 구조ID


def score_for_app(indicator_name, pdf_data):
    """실제 앱 평가경로용. 라우팅된 결정론 지표면 (점수, 근거문자열) 반환, 아니면 None.

    pdf_data(bytes)를 콘텐츠 해시 키로 캐시에 주입 → 같은 보고서 반복호출 시 PDF 파싱 재사용.
    """
    rid = ROUTED.get(_norm_name(indicator_name))
    if not rid:
        return None
    key = "app_" + hashlib.md5(pdf_data).hexdigest()[:12]
    if key not in _BYTES:
        _BYTES[key] = (pdf_data, None)
    s, ev = score({**_CFG_BY_ID[rid], "pdf": key})
    return s, f"{rid} | {ev}"


if __name__ == "__main__":
    import sys
    only = sys.argv[1] if len(sys.argv) > 1 else None
    print("=== 구조적 채점기 스모크 (1회) ===")
    okc = 0
    for ind in INDS:
        if only and only not in ind["id"]:
            continue
        try:
            s, ev = score(ind)
            t = ind["truth"]
            match = "—" if t is None else ("✅" if s == t else "❌")
            if t is not None and s == t:
                okc += 1
            print(f"{ind['id']:16}[{ind['type']:13}] 점수={s} (정답={t}) {match}  {ev}")
        except Exception as e:
            print(f"{ind['id']:16} ❌ {type(e).__name__}: {str(e)[:100]}")
    print(f"\n정답 일치: {okc}/{sum(1 for i in INDS if i['truth'] is not None)} (truth 있는 것 중)")
