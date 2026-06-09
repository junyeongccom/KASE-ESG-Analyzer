"""v4 138지표 → 색상 분류(green 정확 / red 출처틀림 / orange 애매) + v4_categories.json 저장.
미검증 13건은 cross13.json(Claude 교차검증) 반영."""
import json, re
from collections import Counter
from pathlib import Path

BASE = Path("/Users/junyeongc/KASE/AI도입/KASE_ESG_Analyzer")


def norm(s):
    return re.sub(r"[\s\W]", "", str(s or "")).lower()


V4X = next(BASE.glob("output/CJ_v4_v4_*분석결과.xlsx"))
V4J = json.loads((BASE / "output" / (V4X.stem + "_verify.json")).read_text())
dis_by = {norm(x["name"]): x for x in json.loads((BASE / "output/judge_result.json").read_text())}
agr_by = {norm(x["name"]): x for x in json.loads((BASE / "output/external_judge_agreements.json").read_text())}
gr5_by = {norm(x["name"]): x for x in json.loads((BASE / "output/gemini_recheck5.json").read_text())}
_c = BASE / "output/cross13.json"
cross_by = {norm(x["name"]): x for x in json.loads(_c.read_text())} if _c.exists() else {}


def classify(x):
    k, cat, nv = norm(x["name"]), x["cat"], x.get("needs_vision")
    if cat == "공시확인불가":
        return "공시확인불가(사람판단)", "orange"
    if cat != "정상":
        return "누락/기타(사람판단)", "orange"
    if nv:
        return "출처 의심", "red"
    if k in dis_by:
        v = dis_by[k]
        if v.get("conflict"):
            return "점수논란-도메인충돌", "orange"
        if v["verdict"] == "B":
            return "정확", "green"
        return f"점수논란-judge:{v['verdict']}", "orange"
    if k in agr_by:
        a = agr_by[k]
        if a["verdict"] == "correct":
            return "정확", "green"
        if a["verdict"] == "uncertain":
            return "점수논란-외부불확실", "orange"
        g = gr5_by.get(k)
        if g and "정당" in g.get("conclusion", ""):
            return "정확", "green"
        if g:
            return "점수논란-공통과대", "orange"
        return "정확", "green"  # Claude 단독 노이즈 → 합의 정당
    if k in cross_by:
        c = cross_by[k]
        if c["verdict"] == "correct":
            return "정확", "green"
        return f"점수논란-교차검증:{c['verdict']}", "orange"
    return "미검증(교차검증 필요)", "orange"


records = []
for sn, data in V4J["sheets"].items():
    for x in data["rows"]:
        label, color = classify(x)
        records.append({"name": x["name"], "label": label, "color": color})

ccount = Counter(r["color"] for r in records)
print(f"═══ v4 {len(records)}지표 색상 분류 ═══")
print(f"  🟩 green (정확)     : {ccount['green']}")
print(f"  🟥 red   (출처틀림) : {ccount['red']}")
print(f"  🟧 orange(애매)     : {ccount['orange']}")
print("\n[세부 라벨]")
for lab, n in Counter(r["label"] for r in records).most_common():
    print(f"  {n:3}  {lab}")

catmap = {norm(r["name"]): r["color"] for r in records}
(BASE / "output/v4_categories.json").write_text(
    json.dumps({"map": catmap, "records": records}, ensure_ascii=False, indent=2, default=str))
print("\n저장: v4_categories.json")
