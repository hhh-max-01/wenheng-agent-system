import base64
import json
import urllib.request


samples = {
    "complete.txt": "项目名称完整，研究内容明确，考核指标可量化，进度覆盖执行期，审批意见同意立项。",
    "template-residue.txt": "【填写说明：请在此处填写项目相关内容，删除本提示】",
}
files = [
    {"name": name, "base64": base64.b64encode(text.encode("utf-8")).decode("ascii")}
    for name, text in samples.items()
]

payload = json.dumps({
    "dataset_type": "计划任务书",
    "intent": "判断项目材料是否完整、一致且具备可验收性",
    "files": files,
}, ensure_ascii=False).encode("utf-8")
request = urllib.request.Request(
    "http://127.0.0.1:8000/api/analyze",
    data=payload,
    headers={"Content-Type": "application/json"},
    method="POST",
)
with urllib.request.urlopen(request, timeout=30) as response:
    result = json.loads(response.read().decode("utf-8"))

assert result["count"] == 2
assert all(item["label"] in {"通过", "不通过"} for item in result["results"])
assert all(set(["id", "dataset_type", "intent", "label", "matched_rules", "reason"]).issubset(item) for item in result["results"])
print(json.dumps(result, ensure_ascii=False, indent=2))
