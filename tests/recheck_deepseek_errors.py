from __future__ import annotations

import json
import sys
from pathlib import Path


PROJECT = Path(__file__).resolve().parents[1]
WORKSPACE = next(
    (parent for parent in [PROJECT, *PROJECT.parents] if (parent / "work" / "public-data").exists()),
    None,
)
if WORKSPACE is None:
    raise SystemExit("没有找到公开训练数据。")

sys.path.insert(0, str(PROJECT))
import server


previous_path = PROJECT / "evaluation_deepseek.json"
if not previous_path.exists():
    raise SystemExit("没有找到 evaluation_deepseek.json，请先运行完整评估。")
previous = json.loads(previous_path.read_text(encoding="utf-8"))
errors = [item for item in previous["details"] if not item["correct"]]
if not errors:
    raise SystemExit("上次评估没有错误案例，无需复核。")

results = []
for index, old in enumerate(errors, 1):
    path = WORKSPACE / "work" / "public-data" / old["file"]
    print(f"[{index}/{len(errors)}] 重新审核：{path.name}", flush=True)
    text = server.extract_docx(path.read_bytes())
    dataset_type = old["dataset_type"]
    intent = (
        "判断该项目是否可以通过立项申请，重点检查内容完整性、问题方案匹配、创新实质、应用价值和审批合规"
        if dataset_type == "立项申请书"
        else "判断项目材料是否完整、一致且具备可验收性，最终判断是否通过"
    )
    result = server.judge_one(server.Record(path.stem, path.name, text), dataset_type, intent)
    results.append({
        "file": path.name,
        "expected": old["expected"],
        "predicted": result["label"],
        "correct": old["expected"] == result["label"],
        "confidence": result["confidence"],
        "mode": result["mode"],
        "matched_rules": result["matched_rules"],
        "reason": result["reason"],
        "warning": result.get("warning"),
    })

output = PROJECT / "evaluation_recheck.json"
output.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
correct = sum(item["correct"] for item in results)
print(f"\n复核完成：{correct}/{len(results)} 条改进后判断正确")
for item in results:
    print(f"- {item['file']}：预期{item['expected']}，现在判断{item['predicted']}，置信度{item['confidence']:.0%}")
print(f"详细结果：{output}")
