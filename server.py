from __future__ import annotations

import base64
import csv
import io
import json
import os
import re
import tempfile
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

from docx import Document
from pypdf import PdfReader


ROOT = Path(__file__).resolve().parent
WEB_ROOT = ROOT / "web"


def load_local_env() -> None:
    env_path = ROOT / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


load_local_env()
RULES = json.loads((ROOT / "rules.json").read_text(encoding="utf-8"))
MAX_REQUEST_BYTES = 25 * 1024 * 1024
MAX_TEXT_CHARS = 60_000


@dataclass
class Record:
    id: str
    source_name: str
    text: str


def clean_text(text: str) -> str:
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()[:MAX_TEXT_CHARS]


def extract_docx(data: bytes) -> str:
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
        handle.write(data)
        temp_path = Path(handle.name)
    try:
        doc = Document(temp_path)
        blocks = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                blocks.append(" | ".join(cell.text.strip().replace("\n", " / ") for cell in row.cells))
        return clean_text("\n".join(blocks))
    finally:
        temp_path.unlink(missing_ok=True)


def extract_pdf(data: bytes) -> str:
    reader = PdfReader(io.BytesIO(data))
    return clean_text("\n".join(page.extract_text() or "" for page in reader.pages))


def records_from_file(item: dict[str, Any]) -> list[Record]:
    name = str(item.get("name") or "uploaded")
    raw = base64.b64decode(item.get("base64") or "", validate=True)
    suffix = Path(name).suffix.lower()
    stem = Path(name).stem

    if suffix == ".docx":
        return [Record(stem, name, extract_docx(raw))]
    if suffix == ".pdf":
        return [Record(stem, name, extract_pdf(raw))]
    if suffix == ".doc":
        raise ValueError(f"{name} 是旧版 DOC，请先用 Word 另存为 DOCX 后上传")
    if suffix in {".txt", ".md"}:
        return [Record(stem, name, clean_text(raw.decode("utf-8-sig", errors="replace")))]
    if suffix == ".json":
        value = json.loads(raw.decode("utf-8-sig"))
        rows = value if isinstance(value, list) else value.get("data", [value]) if isinstance(value, dict) else [value]
        return [
            Record(str(row.get("id", f"{stem}_{idx:03d}")) if isinstance(row, dict) else f"{stem}_{idx:03d}", name,
                   clean_text(json.dumps(row, ensure_ascii=False)))
            for idx, row in enumerate(rows, 1)
        ]
    if suffix == ".csv":
        decoded = raw.decode("utf-8-sig", errors="replace")
        rows = list(csv.DictReader(io.StringIO(decoded)))
        return [Record(str(row.get("id") or f"{stem}_{idx:03d}"), name, clean_text(json.dumps(row, ensure_ascii=False)))
                for idx, row in enumerate(rows, 1)]
    raise ValueError(f"暂不支持 {suffix or '无扩展名'} 文件：{name}")


def snippet(text: str, pattern: str, width: int = 95) -> str | None:
    match = re.search(pattern, text, re.I)
    if not match:
        return None
    start = max(0, match.start() - 20)
    end = min(len(text), match.end() + width)
    return text[start:end].replace("\n", " ").strip()


def deterministic_checks(text: str, dataset_type: str) -> list[dict[str, str]]:
    checks: list[dict[str, str]] = []
    patterns = [
        ("R-C01", r"填写说明|删除本提示|待补充|请在此处填写", "存在未清理的模板提示或占位内容"),
        ("R-C04", r"百分之|准确率\s*\|\s*\|", "考核指标可能缺失、格式异常或不可量化"),
        ("R-C07", r"\n1\s*\|?\s*[^\n]+\n1\s*\|", "人员或事项序号重复"),
        ("R-C08", r"研究报告\s*\|\s*初稿", "验收成果状态仍为初稿"),
        ("R-C09", r"不同意立项|不予立项|退回修改", "审批意见包含明确否决信息"),
    ]
    for rule_id, pattern, message in patterns:
        evidence = snippet(text, pattern)
        if evidence:
            checks.append({"rule_id": rule_id, "message": message, "evidence": evidence})

    if dataset_type == "计划任务书":
        if "项目执行期限" not in text and "项目起止时间" not in text:
            checks.append({"rule_id": "R-C05", "message": "缺少项目执行期限", "evidence": "未检索到项目执行期限或项目起止时间"})
        if "考核指标" not in text:
            checks.append({"rule_id": "R-C04", "message": "缺少可验收考核指标", "evidence": "未检索到考核指标章节"})

        unit_values = re.findall(r"项目承担单位\s*\|\s*([^\n|]+)", text)
        normalized_units = {value.strip() for value in unit_values}
        if len(normalized_units) > 1:
            checks.append({"rule_id": "R-C03", "message": "项目承担单位前后不一致", "evidence": "；".join(sorted(normalized_units))})
        elif len(normalized_units) == 1:
            stated_unit = next(iter(normalized_units))
            full_names = set(re.findall(r"([\u4e00-\u9fff]{4,30}(?:有限责任公司|供电局|研究院|管理所))", text))
            longer_matches = sorted(name for name in full_names if name.endswith(stated_unit) and name != stated_unit)
            if longer_matches:
                checks.append({"rule_id": "R-C03", "message": "承担单位简称与其他章节全称不一致", "evidence": f"项目承担单位填写“{stated_unit}”，其他章节出现“{longer_matches[0]}”"})

        profile_people = {
            name: age for name, age in re.findall(r"\n([\u4e00-\u9fff]{2,4})\s*\|\s*(?:男|女)\s*\|\s*(\d{1,2})\s*\|", text)
        }
        team_people = {
            name: age for name, age in re.findall(r"\n\d+\s*\|\s*([\u4e00-\u9fff]{2,4})\s*\|\s*(\d{1,2})\s*\|", text)
        }
        for name, age in profile_people.items():
            if name in team_people and team_people[name] != age:
                checks.append({"rule_id": "R-C03", "message": "负责人年龄前后不一致", "evidence": f"{name}在基本信息中为{age}岁，在团队表中为{team_people[name]}岁"})

        team_section = re.search(r"序号\s*\|\s*姓名.*?(?=序号\s*\|\s*考核指标名称)", text, re.S)
        if team_section:
            ids = re.findall(r"\n(\d+)\s*\|", team_section.group(0))
            duplicates = sorted({item for item in ids if ids.count(item) > 1})
            if duplicates:
                checks.append({"rule_id": "R-C07", "message": "团队人员序号重复", "evidence": f"重复序号：{', '.join(duplicates)}"})
            if len(ids) > 20:
                checks.append({"rule_id": "R-C07", "message": "团队人员数量异常", "evidence": f"团队表共列出{len(ids)}行人员"})

        if re.search(r"外部专家.*?(?:茂名|佛山|广州|珠海|东莞)供电局", text, re.S):
            evidence = snippet(text, r"外部专家.*?(?:茂名|佛山|广州|珠海|东莞)供电局")
            checks.append({"rule_id": "R-C07", "message": "存在未在承担单位中说明的外部团队成员", "evidence": evidence or "发现外部专家单位"})

        duration = re.search(r"项目执行期限\s*\|\s*(\d{4})-(\d{2})\s*至\s*(\d{4})-(\d{2})", text)
        schedule_section = re.search(r"序号\s*\|\s*时间段.*?(?=经费类型\s*\|)", text, re.S)
        if duration and schedule_section:
            months = (int(duration.group(3)) - int(duration.group(1))) * 12 + int(duration.group(4)) - int(duration.group(2)) + 1
            stages = re.findall(r"\n\d+\s*\|\s*\d{4}-\d{2}\s*至", schedule_section.group(0))
            if months > 12 and len(stages) < 3:
                checks.append({"rule_id": "R-C05", "message": "进度安排未覆盖完整执行期", "evidence": f"项目执行期约{months}个月，但仅列出{len(stages)}个进度阶段"})

        if duration:
            months = (int(duration.group(3)) - int(duration.group(1))) * 12 + int(duration.group(4)) - int(duration.group(2)) + 1
            if months > 30:
                checks.append({"rule_id": "R-C05", "message": "项目执行期超过常见30个月范围，需复核", "evidence": duration.group(0)})
    else:
        for keyword, label in [("研究内容", "研究内容"), ("技术关键点及创新点", "创新点"), ("应用前景", "应用前景")]:
            if keyword not in text:
                checks.append({"rule_id": "R-A01", "message": f"缺少{label}", "evidence": f"未检索到“{keyword}”章节"})
    return checks


def select_rules(dataset_type: str, intent: str) -> list[dict[str, Any]]:
    words = set(re.findall(r"[\u4e00-\u9fff]{2,}|[A-Za-z]{3,}", intent))
    ranked = []
    for rule in RULES:
        if dataset_type not in rule["dataset_types"] and "通用" not in rule["dataset_types"]:
            continue
        score = sum(2 for word in words if word in rule["name"] or word in rule["description"])
        score += 1 if rule["severity"] == "blocking" else 0
        ranked.append((score, rule))
    ranked.sort(key=lambda item: (-item[0], item[1]["rule_id"]))
    return [item[1] for item in ranked[:10]]


def parse_json_response(content: str) -> dict[str, Any]:
    content = content.strip()
    content = re.sub(r"^```(?:json)?\s*|\s*```$", "", content, flags=re.I)
    start, end = content.find("{"), content.rfind("}")
    if start < 0 or end <= start:
        raise ValueError("模型没有返回 JSON 对象")
    return json.loads(content[start:end + 1])


def call_llm(
    record: Record,
    dataset_type: str,
    intent: str,
    checks: list[dict[str, str]],
    review_note: str | None = None,
) -> dict[str, Any]:
    api_key = os.getenv("LLM_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("LLM_API_KEY 未配置")
    base_url = os.getenv("LLM_BASE_URL", "https://api.deepseek.com").rstrip("/")
    model = os.getenv("LLM_MODEL", "deepseek-chat")
    candidate_rules = select_rules(dataset_type, intent)
    system = """你是业务文档合规判断智能体。必须先分析intent，再提取事实，分别寻找否决证据和通过条件，匹配规则，最后输出硬标签。不能根据项目题目热门程度猜测。evidence必须是文档中的短原文或明确的缺失事实，不得编造。仅输出一个JSON对象，不输出Markdown。"""
    prompt = {
        "task": "根据intent审核文档并返回固定结构结果",
        "dataset_type": dataset_type,
        "intent": intent,
        "record_id": record.id,
        "candidate_rules": candidate_rules,
        "deterministic_checks": checks,
        "decision_policy": [
            "先检查blocking规则；命中且证据充分通常判不通过",
            "未发现否决证据不等于自动通过；通过必须有充分正面证据",
            "信息缺失影响核心判断时判不通过，不要猜测",
            "matched_rules只放真正用于裁决且有证据的规则",
        ],
        "calibration_guardrails": [
            "签字或承诺书日期留空本身不属于模板残留，除非intent明确审核签章完整性",
            "某标题附近没有正文不等于内容缺失；必须检查全文其他表格或章节是否已提供对应信息",
            "预算比较必须确认统计口径相同；费用性明细不能直接与包含资本性的总经费比较",
            "资本性经费未在费用性年度栏展示，不得自动推断预算不闭合",
            "非关键格式瑕疵不能替代与当前intent有关的实质否决证据",
        ],
        "required_output": {
            "label": "通过 或 不通过",
            "matched_rules": [{"rule_id": "规则编号", "rule_name": "规则名", "evidence": "文档原文证据"}],
            "reason": "不超过120字",
            "confidence": "0到1之间的小数",
        },
        "document": record.text,
    }
    if review_note:
        prompt["independent_review"] = review_note
    messages = [
        {"role": "system", "content": system},
        {"role": "user", "content": json.dumps(prompt, ensure_ascii=False)},
    ]
    last_error: Exception | None = None
    for attempt in range(2):
        if attempt:
            messages.append({"role": "user", "content": "上一次返回无法解析。请严格只返回一个有效JSON对象，不要输出解释、代码围栏或思考过程。"})
        body = json.dumps({
            "model": model,
            "temperature": 0,
            "max_tokens": 1400,
            "response_format": {"type": "json_object"},
            "messages": messages,
        }, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            f"{base_url}/chat/completions",
            data=body,
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=90) as response:
                payload = json.loads(response.read().decode("utf-8"))
            content = payload["choices"][0]["message"].get("content") or ""
            return parse_json_response(content)
        except urllib.error.HTTPError as exc:
            messages_by_status = {
                400: "DeepSeek拒绝了请求参数",
                401: "DeepSeek API Key无效或未生效",
                402: "DeepSeek账户余额不足",
                429: "DeepSeek请求过于频繁，请稍后重试",
            }
            raise RuntimeError(messages_by_status.get(exc.code, f"DeepSeek接口返回HTTP {exc.code}")) from exc
        except urllib.error.URLError as exc:
            raise RuntimeError("无法连接DeepSeek，请检查网络、代理或防火墙") from exc
        except (ValueError, KeyError, json.JSONDecodeError) as exc:
            last_error = exc
    raise RuntimeError(f"模型连续两次未返回有效JSON：{last_error}")


def heuristic_result(record: Record, dataset_type: str, intent: str, checks: list[dict[str, str]]) -> dict[str, Any]:
    blocking = {rule["rule_id"]: rule for rule in RULES if rule["severity"] == "blocking"}
    matched = []
    seen = set()
    for check in checks:
        rule = blocking.get(check["rule_id"])
        signature = (check["rule_id"], check["evidence"])
        if rule and signature not in seen:
            seen.add(signature)
            matched.append({"rule_id": rule["rule_id"], "rule_name": rule["name"], "evidence": check["evidence"]})
    label = "不通过" if matched else "通过"
    reason = "发现影响材料有效性或可验收性的明确问题。" if matched else "本地预检查未发现明确否决项；正式结果仍建议由模型复核。"
    return {"label": label, "matched_rules": matched, "reason": reason, "confidence": 0.82 if matched else 0.56}


def normalize_result(raw: dict[str, Any], record: Record, dataset_type: str, intent: str, mode: str) -> dict[str, Any]:
    label = raw.get("label") if raw.get("label") in {"通过", "不通过"} else "不通过"
    rules = raw.get("matched_rules") if isinstance(raw.get("matched_rules"), list) else []
    normalized_rules = []
    for item in rules[:8]:
        if isinstance(item, dict):
            normalized_rules.append({
                "rule_id": str(item.get("rule_id", "R-UNKNOWN")),
                "rule_name": str(item.get("rule_name", "未命名规则")),
                "evidence": str(item.get("evidence", "未提供证据"))[:300],
            })
    try:
        confidence = max(0.0, min(1.0, float(raw.get("confidence", 0.5))))
    except (TypeError, ValueError):
        confidence = 0.5
    return {
        "id": record.id,
        "source_name": record.source_name,
        "dataset_type": dataset_type,
        "intent": intent,
        "label": label,
        "matched_rules": normalized_rules,
        "reason": str(raw.get("reason", "未提供判断说明"))[:300],
        "confidence": confidence,
        "mode": mode,
    }


def judge_one(record: Record, dataset_type: str, intent: str) -> dict[str, Any]:
    checks = deterministic_checks(record.text, dataset_type)
    if os.getenv("LLM_API_KEY", "").strip():
        try:
            raw = call_llm(record, dataset_type, intent, checks)
            if raw.get("label") == "不通过" and not checks:
                first_result = json.dumps(raw, ensure_ascii=False)
                raw = call_llm(
                    record,
                    dataset_type,
                    intent,
                    checks,
                    review_note=(
                        "这是一次反误杀独立复核。程序未发现高置信度结构问题。"
                        "请质疑第一次结论是否把签字日期空白、跨章节内容位置或不同预算口径误当成否决证据。"
                        "如果仍有与intent直接相关、证据充分的实质问题，可以维持不通过；否则改为通过。"
                        f"第一次结果：{first_result}"
                    ),
                )
            return normalize_result(raw, record, dataset_type, intent, "llm")
        except Exception as exc:
            fallback = heuristic_result(record, dataset_type, intent, checks)
            result = normalize_result(fallback, record, dataset_type, intent, "fallback")
            result["warning"] = f"模型调用失败，已使用本地预检查：{str(exc)[:160]}"
            return result
    return normalize_result(heuristic_result(record, dataset_type, intent, checks), record, dataset_type, intent, "demo")


class AppHandler(SimpleHTTPRequestHandler):
    def translate_path(self, path: str) -> str:
        clean = path.split("?", 1)[0].split("#", 1)[0]
        if clean == "/":
            clean = "/index.html"
        return str(WEB_ROOT / clean.lstrip("/"))

    def send_json(self, value: Any, status: int = 200) -> None:
        data = json.dumps(value, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        if self.path == "/api/health":
            self.send_json({
                "ok": True,
                "mode": "llm" if os.getenv("LLM_API_KEY") else "demo",
                "provider": "DeepSeek" if "deepseek" in os.getenv("LLM_BASE_URL", "https://api.deepseek.com").lower() else "OpenAI-compatible",
                "model": os.getenv("LLM_MODEL", "deepseek-chat"),
            })
            return
        super().do_GET()

    def do_POST(self) -> None:
        if self.path != "/api/analyze":
            self.send_json({"error": "接口不存在"}, HTTPStatus.NOT_FOUND)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length <= 0 or length > MAX_REQUEST_BYTES:
                raise ValueError("请求为空或超过25MB限制")
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            dataset_type = str(payload.get("dataset_type", ""))
            intent = str(payload.get("intent", "")).strip()
            files = payload.get("files", [])
            if dataset_type not in {"计划任务书", "立项申请书"}:
                raise ValueError("请选择正确的数据集类型")
            if not intent:
                raise ValueError("请输入判断 intent")
            if not isinstance(files, list) or not files:
                raise ValueError("请至少上传一个文件")
            records: list[Record] = []
            for item in files:
                records.extend(records_from_file(item))
            if not records:
                raise ValueError("上传文件中没有可判断的数据")
            results = [None] * len(records)
            with ThreadPoolExecutor(max_workers=min(3, len(records))) as pool:
                futures = {pool.submit(judge_one, record, dataset_type, intent): idx for idx, record in enumerate(records)}
                for future in as_completed(futures):
                    results[futures[future]] = future.result()
            self.send_json({
                "results": results,
                "count": len(results),
                "mode": "llm" if os.getenv("LLM_API_KEY") else "demo",
            })
        except (ValueError, json.JSONDecodeError, base64.binascii.Error) as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self.send_json({"error": f"系统处理失败：{type(exc).__name__}"}, HTTPStatus.INTERNAL_SERVER_ERROR)


def main() -> None:
    host = os.getenv("APP_HOST", "127.0.0.1")
    port = int(os.getenv("PORT", "8000"))
    server = ThreadingHTTPServer((host, port), AppHandler)
    mode = "LLM" if os.getenv("LLM_API_KEY") else "演示"
    print(f"业务文档审核系统已启动：http://{host}:{port}（{mode}模式）", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
