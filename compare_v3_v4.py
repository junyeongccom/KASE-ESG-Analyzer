"""Phase 3: v3 vs v4 비교.
두 검증 JSON(verify_results 산출)을 읽어 지표명으로 매칭, 완전성·점수·grounding 차이 산출.
매칭 안 되면 행 순서(row order) 폴백.

사용: python compare_v3_v4.py <v3_verify.json> <v4_verify.json>
"""
import sys, re, json
from pathlib import Path

SHEETMAP = [("식품-E", "식품_환경"), ("식품-S", "식품_사회")]  # (v3, v4)


def norm(s):
    return re.sub(r"[\s\W]", "", str(s or "")).lower()


def to_float(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def match_rows(r3, r4):
    """이름 정규화 매칭, 실패분은 행순서 폴백."""
    by3 = {norm(x["name"]): x for x in r3}
    by4 = {norm(x["name"]): x for x in r4}
    pairs, used4 = [], set()
    for x in r3:
        k = norm(x["name"])
        if k in by4:
            pairs.append((x, by4[k]))
            used4.add(k)
    name_matched = len(pairs)
    # 폴백: 남은 것들을 행순서로
    rem3 = [x for x in r3 if norm(x["name"]) not in by4]
    rem4 = [x for x in r4 if norm(x["name"]) not in used4]
    for a, b in zip(rem3, rem4):
        pairs.append((a, b))
    return pairs, name_matched, len(rem3), len(rem4)


def main(v3j, v4j):
    v3 = json.loads(Path(v3j).read_text())
    v4 = json.loads(Path(v4j).read_text())
    print(f"\n===== v3 vs v4 비교 =====\n v3={v3['file']}\n v4={v4['file']}")
    grand = {"better": 0, "worse": 0, "same": 0}
    for s3, s4 in SHEETMAP:
        r3 = v3["sheets"].get(s3, {}).get("rows", [])
        r4 = v4["sheets"].get(s4, {}).get("rows", [])
        if not r3 or not r4:
            print(f"\n── {s3}↔{s4}: 데이터 없음 (v3 {len(r3)}, v4 {len(r4)})")
            continue
        pairs, nmatch, u3, u4 = match_rows(r3, r4)
        print(f"\n── {s3}({len(r3)}) ↔ {s4}({len(r4)}) │ 이름매칭 {nmatch} + 행순서폴백 {len(pairs)-nmatch} = {len(pairs)}쌍 │ 미매칭 v3:{u3} v4:{u4}")

        # 완전성
        miss3 = sum(1 for x in r3 if x["cat"] == "누락(응답/오류)")
        miss4 = sum(1 for x in r4 if x["cat"] == "누락(응답/오류)")
        print(f"   ①완전성: 누락 v3={miss3} vs v4={miss4}")

        # grounding (정상 + 수치보유)
        def gstat(rows):
            g = [x["grounding"] for x in rows if x["cat"] == "정상" and x["grounding"] is not None]
            hi = sum(1 for v in g if v >= 0.7)
            return f"{hi}/{len(g)}" if g else "0/0"
        print(f"   ③출처 高grounding: v3={gstat(r3)} vs v4={gstat(r4)}")

        # 점수 비교 (양쪽 정상인 쌍만)
        both, up, down, same = 0, 0, 0, 0
        notable = []
        for a, b in pairs:
            if a["cat"] == "정상" and b["cat"] == "정상":
                sa, sb = to_float(a["score"]), to_float(b["score"])
                if sa is None or sb is None:
                    continue
                both += 1
                if sb > sa:
                    up += 1
                elif sb < sa:
                    down += 1
                else:
                    same += 1
                if abs(sb - sa) >= 1:
                    notable.append((a, b, sa, sb))
        print(f"   ②점수(양쪽 정상 {both}쌍): v4↑ {up} │ v4↓ {down} │ 동일 {same}")
        grand["better"] += up
        grand["worse"] += down
        grand["same"] += same
        # 큰 점수차 상위
        for a, b, sa, sb in sorted(notable, key=lambda t: abs(t[3] - t[2]), reverse=True)[:6]:
            print(f"     Δ{sb-sa:+.1f}  v3={sa} v4={sb} g(v3={a['grounding']},v4={b['grounding']}) │ {a['name'][:38]}")
        # 한쪽만 누락/공시확인불가 (완전성·평가차이 후보)
        flip = [(a, b) for a, b in pairs if a["cat"] != b["cat"]]
        if flip:
            print(f"   △ 상태 불일치 {len(flip)}쌍 (예시):")
            for a, b in flip[:5]:
                print(f"     v3[{a['cat']}] vs v4[{b['cat']}] │ {a['name'][:38]}")
    print(f"\n═══ 종합 점수변화(양쪽 정상): v4↑ {grand['better']} │ v4↓ {grand['worse']} │ 동일 {grand['same']} ═══")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
