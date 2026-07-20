# 文衡：智能业务文档审核系统

文衡是本次“智能体编排业务判断挑战”的可运行实现。评委可以选择`计划任务书`或`立项申请书`，上传业务文档并输入判断 Intent。系统返回每条数据的硬标签、命中规则、文档证据和简短理由。

## 1. 系统能做什么

- 选择两类数据集：计划任务书、立项申请书。
- 上传一个或多个 DOCX、PDF、JSON、CSV、TXT、Markdown 文件。
- 从 JSON 数组或 CSV 表格中拆分多条数据。
- 分析用户 Intent，只选择相关判断规则。
- 先执行确定性预检查，再让大模型完成语义分析和裁决。
- 输出固定 JSON 结构。
- API 不可用时自动降级为本地规则演示模式。
- 在页面展示每条记录的规则、证据、理由和置信度。

> 旧版 `.doc` 不适合服务器稳定解析。请先用 Word 另存为 `.docx` 后上传。

## 2. 新手快速运行（Windows）

### 第一步：安装 Python

安装 Python 3.11 或更高版本。安装时勾选“Add Python to PATH”。

### 第二步：安装两个文档解析依赖

直接双击 `install_dependencies.bat`。它会在项目内部创建独立的 `.venv`，然后安装Word和PDF解析组件，不影响其他Python项目。

如果希望手动安装，可以在 VS Code 中打开本项目，然后打开“终端”，执行：

```powershell
py -m pip install -r requirements.txt
```

如果电脑无法识别 `py`，改用：

```powershell
python -m pip install -r requirements.txt
```

### 第三步：配置模型

双击 `setup_deepseek.bat`。它会创建只保存在本机的 `.env` 并用记事本打开。将Key填写在第一行等号后面：

```text
LLM_API_KEY=你的API密钥
LLM_BASE_URL=https://api.deepseek.com
LLM_MODEL=deepseek-chat
```

不要把 `.env` 上传到 GitHub，也不要把真实密钥写进 README 或源代码。

如果暂时不填写 API Key，系统仍能启动，但只运行本地规则演示模式。

保存后，可以先运行连接测试：

```powershell
py tests/test_deepseek_connection.py
```

测试只显示连接是否成功，不会打印你的Key。

也可以直接双击 `test_deepseek.bat`，看到“连接成功”后再启动网站。

### 第四步：启动

双击 `start_windows.bat`，或在终端执行：

```powershell
py server.py
```

浏览器打开：

```text
http://127.0.0.1:8000
```

## 3. 输出示例

```json
{
  "id": "sample_001",
  "dataset_type": "计划任务书",
  "intent": "判断创新程度是否足够",
  "label": "通过",
  "matched_rules": [
    {
      "rule_id": "R-A03",
      "rule_name": "创新实质性",
      "evidence": "文档中的相关原文"
    }
  ],
  "reason": "项目提出可区分的技术方案并给出可验证依据。",
  "confidence": 0.84
}
```

## 4. 项目结构

```text
server.py               网页服务、文档解析、智能体流程和模型调用
rules.json              稳定的候选规则库
web/index.html           中文网页结构
web/styles.css           网页视觉样式
web/app.js               上传、运行和结果展示
tests/                   接口测试和公开子集自检
design.md                系统设计与泛化说明
.env.example             安全配置示例
requirements.txt         Python依赖
start_windows.bat        Windows一键启动
install_dependencies.bat 一键创建环境并安装依赖
setup_deepseek.bat       安全创建本地DeepSeek配置
test_deepseek.bat        双击测试DeepSeek连接
Dockerfile               公网容器部署入口
```

## 5. DeepSeek API 配置

系统默认使用 DeepSeek 的 OpenAI-compatible `chat/completions` 接口。部署时只需设置：

- `LLM_API_KEY`
- `LLM_BASE_URL`
- `LLM_MODEL`

密钥只保存在服务器环境变量中，不会发送到浏览器。`deepseek-chat`用于主审核；系统使用JSON输出模式，并在格式异常时自动重试一次。

## 6. 测试

服务启动后运行：

```powershell
py tests/test_api_smoke.py
```

公开可读子集的本地规则自检：

```powershell
py tests/evaluate_public_subset.py
```

公开数据得分不代表隐藏集表现。该测试只用于发现解析错误、规则误报和输出格式问题。

## 7. 公网部署

部署平台需要满足：

- 可运行 Python 3.11+
- 可安装 `requirements.txt`
- 能设置环境变量
- 对外开放由 `PORT` 指定的端口
- 评测期间保持服务在线

启动命令：

```text
python server.py
```

### 使用 Render 部署

项目已提供 `render.yaml`，推荐步骤：

1. 把不含`.env`的项目源代码上传到GitHub仓库。
2. 登录Render，选择新建Blueprint或Web Service并连接该仓库。
3. Render会读取`render.yaml`并创建Python服务。
4. 在Render的环境变量设置中填写`LLM_API_KEY`，标记为Secret。
5. 等待部署完成，访问`onrender.com`公网地址。
6. 上传一份通过和一份不通过样本进行验收。

免费服务闲置后可能休眠，首次访问需要等待唤醒。提交前应提前打开公网URL确认服务处于运行状态。

正式部署前应使用无登录浏览器完成一次全流程测试，并确认 API Key 未进入前端文件或代码仓库。

## 8. 已知限制

- 不直接解析传统二进制 `.doc`；需转换为 `.docx`。
- PDF 为扫描图片且没有文本层时，需要额外 OCR，目前不会自动识别图片文字。
- 文档超过60,000字符时会截断，以控制调用成本和时延。
- 没有 API Key 时的演示结果只适合检查流程，不能代替正式模型判断。
