import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import server


if not server.os.getenv("LLM_API_KEY", "").strip():
    raise SystemExit("未发现 LLM_API_KEY。请先运行 setup_deepseek.bat，并在 .env 中填写 Key。")

record = server.Record(
    id="connection_test",
    source_name="connection_test.txt",
    text="项目材料完整，研究内容明确，指标可量化，审批意见同意立项。",
)
result = server.judge_one(record, "立项申请书", "判断模型连接是否正常，并按材料作出测试判断")
if result["mode"] != "llm":
    raise SystemExit(f"连接失败，系统进入了 {result['mode']} 模式：{result.get('warning', '未返回详细原因')}")
print(f"连接成功：DeepSeek 返回标签 {result['label']}，结构化结果字段完整。")
