from __future__ import annotations

import json
import sys
from pathlib import Path


PROJECT = Path(__file__).resolve().parents[1]
WORKSPACE = PROJECT
for parent in [PROJECT, *PROJECT.parents]:
    if (parent / "work" / "public-data").exists():
        WORKSPACE = parent
        break

sys.path.insert(0, str(PROJECT))
import server


DATA = WORKSPACE / "work" / "public-data"
LABEL_FILE = WORKSPACE / "work" / "extracted" / "labels.json"
OUTPUT = PROJECT / "evaluation_deepseek.json"

if not DATA.exists() or not LABEL_FILE.exists():
    raise SystemExit("没有找到公开训练数据。请在原Codex作业工作区中运行本测试。")
if not server.os.getenv("LLM_API_KEY", "").strip():
    raise SystemExit("没有找到DeepSeek Key，请检查 .env。")

label_rows = json.loads(LABEL_FILE.read_text(encoding="utf-8"))["审核结果"]
label_map = {row[0]: row[1] for row in label_rows[1:]}
paths = [
    path for path in sorted(DATA.glob("*.docx"))
    if path.name != "14 一种SF6气体试验数据自动录入装置研制.docx"
]

details = []
for index, path in enumerate(paths, 1):
    print(f"[{index}/{len(paths)}] 正在审核：{path.name}", flush=True)
    text = server.extract_docx(path.read_bytes())
    dataset_type = "立项申请书" if "立项申请书" in text else "计划任务书"
    intent = (
        "判断该项目是否可以通过立项申请，重点检查内容完整性、问题方案匹配、创新实质、应用价值和审批合规"
        if dataset_type == "立项申请书"
        else "判断项目材料是否完整、一致且具备可验收性，最终判断是否通过"
    )
    record = server.Record(path.stem, path.name, text)
    result = server.judge_one(record, dataset_type, intent)
    expected = label_map[path.name]
    details.append({
        "file": path.name,
        "dataset_type": dataset_type,
        "expected": expected,
        "predicted": result["label"],
        "correct": expected == result["label"],
        "mode": result["mode"],
        "confidence": result["confidence"],
        "matched_rules": result["matched_rules"],
        "reason": result["reason"],
        "warning": result.get("warning"),
    })

total = len(details)
correct = sum(item["correct"] for item in details)
by_label = {}
for label in ["通过", "不通过"]:
    subset = [item for item in details if item["expected"] == label]
    by_label[label] = {
        "total": len(subset),
        "correct": sum(item["correct"] for item in subset),
        "recall": (sum(item["correct"] for item in subset) / len(subset)) if subset else None,
    }

report = {
    "summary": {
        "total": total,
        "correct": correct,
        "accuracy": correct / total if total else 0,
        "by_label": by_label,
        "fallback_count": sum(item["mode"] != "llm" for item in details),
    },
    "details": details,
}
OUTPUT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

print("\n评估完成")
print(f"总体：{correct}/{total}，正确率 {report['summary']['accuracy']:.1%}")
for label, values in by_label.items():
    print(f"{label}：{values['correct']}/{values['total']}，召回率 {values['recall']:.1%}")
print(f"降级调用：{report['summary']['fallback_count']} 条")
print(f"详细结果：{OUTPUT}")
