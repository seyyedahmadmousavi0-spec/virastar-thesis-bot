"""Rule-specific synthetic evaluation. No user text is logged.

Positive fixtures label exactly one rule; other detections are unlabelled.
Negative fixtures label absence of error/warn for each supported target rule.
Precision/recall are fixture metrics, not deployment accuracy or probabilities.
"""
import json
from pathlib import Path
import sys
import checker

def score(expected, found):
    return int(expected and found), int(not expected and found), int(expected and not found), int(not expected and not found)

def run_benchmark(output_path=None):
    cases = json.loads((Path(__file__).parent / "tests" / "benchmark_cases.json").read_text(encoding="utf8"))
    rules = checker.load_rules()
    targets = sorted({rule for _, expected, rule in cases if expected})
    counts = {rule: dict(tp=0, fp=0, fn=0, tn=0) for rule in targets}
    unlabelled = 0
    manual_review = 0
    for text, expected, target in cases:
        issues = checker.check_text(0, text, "synthetic fixture", rules)
        manual_review += sum(i.confidence in {"B", "C"} for i in issues)
        found = {i.rule_id for i in issues}
        selected = [target] if expected else targets
        if expected:
            unlabelled += len(found - {target})
        for rule in selected:
            detected = rule in found if expected else any(i.rule_id == rule and i.severity in {"error","warn"} for i in issues)
            values = score(expected, detected)
            for key, value in zip(("tp","fp","fn","tn"),values):
                counts[rule][key] += value
    for row in counts.values():
        row["precision"] = row["tp"]/(row["tp"]+row["fp"]) if row["tp"]+row["fp"] else None
        row["recall"] = row["tp"]/(row["tp"]+row["fn"]) if row["tp"]+row["fn"] else None
    result = {"dataset":"synthetic-smoke-v1", "human_validated":False,
              "samples":len(cases), "negative_policy":"error/warn only; stylistic info on negatives is not labelled",
              "scope":"Only designated positive rules and error/warn on negative fixtures are scored. No population estimate.",
              "unlabelled_other_detections":unlabelled, "manual_review_detections":manual_review, "per_rule":counts}
    print(json.dumps(result,ensure_ascii=False,indent=2))
    if output_path:
        Path(output_path).write_text(json.dumps(result,ensure_ascii=False,indent=2)+"\n",encoding="utf8")
    return int(any(row["fp"] or row["fn"] for row in counts.values()))

if __name__ == "__main__":
    import argparse
    parser=argparse.ArgumentParser()
    parser.add_argument("--output")
    args=parser.parse_args()
    sys.exit(run_benchmark(args.output))
