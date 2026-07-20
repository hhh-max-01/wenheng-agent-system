import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import server


DATA = ROOT / "work" / "public-data"
LABELS = json.loads((ROOT / "work" / "extracted" / "labels.json").read_text(encoding="utf-8"))["审核结果"]
label_map = {row[0]: row[1] for row in LABELS[1:]}

rows = []
for path in sorted(DATA.glob("*.docx")):
    if path.name == "14 一种SF6气体试验数据自动录入装置研制.docx":
        continue
    record = server.Record(path.stem, path.name, server.extract_docx(path.read_bytes()))
    dataset_type = "立项申请书" if "立项申请书" in record.text else "计划任务书"
    result = server.judge_one(record, dataset_type, "判断项目材料是否完整、一致且具备可验收性")
    expected = label_map[path.name]
    rows.append({"file": path.name, "expected": expected, "predicted": result["label"], "ok": expected == result["label"], "rules": [r["rule_id"] for r in result["matched_rules"]]})

correct = sum(row["ok"] for row in rows)
print(json.dumps({"correct": correct, "total": len(rows), "accuracy": correct / len(rows), "rows": rows}, ensure_ascii=False, indent=2))
