"""구조적 채점기 — 전 기업(6) × 14지표 × 10회 일관성 검증. 증분저장 + 재개 + 타임아웃(스코어러). 끊겨도 안전."""
import sys, json, time, os, threading
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
sys.path.insert(0, "/Users/junyeongc/KASE/AI도입/KASE_ESG_Analyzer")
from structural_scorer import INDS, score, pdf_text, pdf_bytes

COMPANIES = [("CJ제일", "CJ제일제당"), ("풀무원", "풀무원"), ("오뚜기", "오뚜기"),
             ("daesang", "대상"), ("nestle", "네슬레"), ("unilever", "유니레버")]
RUNS = 10
DIR = "/Users/junyeongc/KASE/AI도입/4_결과/v1.3.0_structural"
PARTIAL = f"{DIR}/structural_repro_allco_partial.json"
OUT = f"{DIR}/structural_repro_allco.json"

# 재개: 기존 부분결과 로드 (성공분만 유지)
res = {}
if os.path.exists(PARTIAL):
    try:
        res = json.load(open(PARTIAL)).get("raw", {})
        for co in res:
            for iid in res[co]:
                res[co][iid] = [s for s in res[co][iid] if not (isinstance(s, str) and "ERR" in s)]
        print(f"[재개] 기존 성공분 로드: {sum(len(v) for co in res.values() for v in co.values())}콜", flush=True)
    except Exception:
        res = {}

for sub, _ in COMPANIES:
    pdf_bytes(sub); pdf_text(sub)

# 필요한 것만 작업 생성 (성공 RUNS개 미만인 셀)
tasks = []
for sub, coname in COMPANIES:
    for ind in INDS:
        have = len(res.get(coname, {}).get(ind["id"], []))
        for _ in range(max(0, RUNS - have)):
            tasks.append((sub, coname, ind))
print(f"=== 검증: 남은 작업 {len(tasks)}콜 (총 {len(COMPANIES)*len(INDS)*RUNS} 목표) ===", flush=True)

lock = threading.Lock()


def run_one(sub, coname, ind):
    for attempt in range(2):
        try:
            s, _ = score({**ind, "pdf": sub})
            return (coname, ind["id"], s)
        except Exception as e:
            if attempt == 1:
                return (coname, ind["id"], f"ERR:{type(e).__name__}")
            time.sleep(1)


def save():
    json.dump({"raw": res}, open(PARTIAL, "w"), ensure_ascii=False, default=str)


done = 0; t0 = time.time()
with ThreadPoolExecutor(max_workers=6) as ex:
    futs = [ex.submit(run_one, *t) for t in tasks]
    for f in as_completed(futs):
        coname, iid, s = f.result()
        with lock:
            res.setdefault(coname, {}).setdefault(iid, []).append(s)
            done += 1
            if done % 20 == 0:
                save()
                print(f"  진행 {done}/{len(tasks)} | {time.time()-t0:.0f}s (저장됨)", flush=True)
save()

# 집계
summary = {}
for coname in res:
    for iid, scores in res[coname].items():
        good = [x for x in scores if not (isinstance(x, str) and "ERR" in x)]
        cnt = Counter(good)
        mode = cnt.most_common(1)[0][0] if cnt else None
        summary.setdefault(iid, {})[coname] = {"scores": good, "mode": mode, "consistent": len(set(good)) <= 1 and len(good) == RUNS}
json.dump({"raw": res, "summary": summary}, open(OUT, "w"), ensure_ascii=False, indent=2, default=str)

NONDET = {"#2 SOx개선", "#6 정보사고", "#7 제3자인증"}
errs = sum(1 for co in res for iid in res[co] for s in res[co][iid] if isinstance(s, str) and "ERR" in s)
cells = [(iid, co) for iid in summary for co in summary[iid] if iid not in NONDET]
con = sum(1 for iid, co in cells if summary[iid][co]["consistent"])
print(f"\n=== 결과 === ERR {errs} | 결정론 {len(cells)}셀 중 일관 {con}", flush=True)
for ind in INDS:
    iid = ind["id"]
    if iid in NONDET: continue
    row = summary.get(iid, {})
    print(f"  {iid:16}: 일관 {sum(1 for co in row if row[co]['consistent'])}/{len(row)}사", flush=True)
print(f"→ 저장 완료", flush=True)
