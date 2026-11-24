# 会议纪要 Agent

Agent 监听现成的会议文字稿和项目记录，调用阿里云百炼整理内容，将决定、行动项和待确认事项写入 SQLite，再生成带原生图表的 Word。它按项目生成周报并投递到本地收件箱，也支持发布到已配置的 Slack 群聊。

## 本地演示

在源码目录用 Python 3.11 或以上版本执行以下命令。安装完成后，直接使用 `.venv/bin/python` 运行 Agent。

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m examples.demo
.venv/bin/python -m unittest discover -s tests -v
```

在 Windows PowerShell 中用 `py -3 -m venv .venv` 创建环境，后续命令使用 `.\.venv\Scripts\python.exe`。

离线演示读入两份模拟文字稿，用 `examples/` 下的 JSON 代替推理结果。演示在 `data/demo/documents/portal/` 下生成两份 Word，并在终端预览周报。本期结果为新增 4 项、完成 3 项、受阻 1 项。第二份记录更新 `A0002` 和 `A0003` 为已完成，保留此前的负责人和截止日。

## Agent 的工作流程

```text
文字稿进入收件箱或按文件导入历史记录
            ↓
分段编号 → 筛选相关历史事项 → Qwen 提取 → 校验字段与原文
            ↓
关联行动项与待确认问题 → 保存会议版本 → 渲染 Word 和图表
            ↓
每周五按项目统计变化 → 本地收件箱 → 可选发布到 Slack
```

| 模块 | 职责 |
|---|---|
| `extract.py` | 读取 `.txt`、`.md` 和 `.docx`，分段后优先提供原文引用的编号和名称相关的候选。要求模型输出摘要议题、决定、行动项、待确认问题和原文依据。 |
| `ledger.py` | 保存会议版本与行动项事件，按日期重放状态。明确新增任务生成新编号。更新优先使用原文已有编号，或匹配唯一同名且负责信息兼容的未结任务，归属不明确时进入待确认。 |
| `word.py` | 输出摘要、决定、行动项表、原文跳转和带嵌入数据表的 Word 原生条形图。图表从本次行动项的最终状态计数。 |
| `app.py` | 提供 `run`、`ingest`、`report` 命令，处理会议修订、待生成 Word 恢复、周报预览和群聊发送去重。 |
| `worker.py` | 监听文件收件箱，防止同一文件重复推理，给暂时性网络错误安排退避重试。运行中更新周报预览，按配置时区在周五 17:00 发布当期报告。 |

模型为每条记录输出 `kind`、`text`、`paragraph` 和 `quote`。行动项带负责人、截止日期、状态和操作类型。程序检查引文、负责人及日期格式，并把常见的未定提议、计划中的完成和与明确原文日期冲突的条目放入待确认。真实会议稿评测用于衡量语义、相对日期和任务识别的准确性。

SQLite 保存规范化原文、结构化条目、会议修订和正式推理回执。更正旧会议后，从当前有效版本重新计算受影响的行动项，并更新下游 Word。待确认问题以 `Q0001` 这类编号跨周保留。后续文字稿可以用「Q0001 已确认上线日期」关闭它。编号按项目分配。部署时把数据库和输出目录放在持久存储。

## 接入真实模型和工作流

### 配置阿里云百炼

百炼公共 API 已于 2026 年 10 月 10 日退役 `qwen3-32b`。项目默认使用百炼在新加坡和北京提供的 `qwen3.8-27b`，并从启动进程的 `DASHSCOPE_API_KEY` 读取标准按量付费 API Key。 [32B 退役公告](https://www.alibabacloud.com/en/notice/model_studio_notice_on_the_postponed_retirement_of_selected_legacy_models_7d0?_p_lc=1) · [百炼模型与地区](https://www.alibabacloud.com/help/en/model-studio/rate-limit) · [获取 API Key](https://www.alibabacloud.com/help/en/model-studio/get-api-key)

```bash
read -rsp '百炼 API Key: ' DASHSCOPE_API_KEY
printf '\n'
export DASHSCOPE_API_KEY
```

默认 Base URL 是新加坡的 `https://dashscope-intl.aliyuncs.com/compatible-mode/v1`。如果密钥在北京创建，同一个终端需设置北京地址。也可以将 `DASHSCOPE_BASE_URL` 设为控制台显示的同地域业务空间地址。 [百炼官方地域地址](https://www.alibabacloud.com/help/tc/model-studio/base-url)

```bash
export DASHSCOPE_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1
```

Agent 通过百炼的 Chat Completions API，在关闭思考模式的情况下请求 JSON Object。提示词描述所需字段，Pydantic 在本机校验结构，原文引用校验通过后才写台账和 Word。 [百炼 Chat Completions API](https://www.alibabacloud.com/help/en/model-studio/qwen-api-via-openai-chat-completions) · [百炼结构化输出](https://docs.modelstudio.console.alibabacloud.com/en/model-studio/qwen-structured-output)

如果已有百炼 Qwen3-32B 专属部署，将 `DASHSCOPE_MODEL` 设为控制台显示的调用 ID，`DASHSCOPE_BASE_URL` 设为对应业务空间地址，并使用该空间的 API Key。 [百炼模型部署文档](https://www.alibabacloud.com/help/en/model-studio/model-deployment-api-reference)

### 启动完整流程

已配置百炼 Key 后，从项目目录启动。可以用 `Ctrl+C` 停止。

```bash
.venv/bin/python -m meeting_agent run
```

监控目录默认是 `data/inbox/`。把写入完毕的文字稿放进去，项目和会议编号用英文、数字、下划线或连字符。文件名形如 `2026-11-23__m-2026-11-23__客户门户上线准备.txt`，项目用一级子目录区分。直接放在收件箱根目录的文件归到 `general` 项目。

```bash
mkdir -p data/inbox/portal
cp examples/2026-11-23.txt 'data/inbox/portal/2026-11-23__m-2026-11-23__客户门户上线准备.txt'
```

Word 生成到 `data/documents/<项目>/<会议编号>.docx`。打开 `data/reports/<项目>/latest.md` 可查看最新预览和文件链接。运行中的 Agent 每周五 17:00（默认时区 `Asia/Taipei`）把周报写入同一目录下按期截止日命名的 Markdown 文件。运行 `run --once` 可立即扫一次并退出。

同一个文件内容变更时，保持项目和会议编号，替换原文件即可开始修订。暂时性网络故障会自动退避重试。修复权限或输入问题后，可以改好源文件，或用 `run --retry-failed` 重试原文件。

### 从外部工作流导入

会议系统、文件同步任务或工作流工具取得文本后，用项目编号和上游记录的稳定编号调用同一条命令。历史工作记录可逐份导入，首次导入建议按日期排序。

```bash
.venv/bin/python -m meeting_agent ingest ./examples/2026-11-23.txt \
  --project portal --meeting m-2026-11-23 --date 2026-11-23 \
  --title '客户门户上线准备'
```

相同项目、编号和内容的重跑直接复用记录，Word 缺失时会重新生成。编号相同、内容变更会更新当前版本并归档旧版。后续会议引用 `A0001` 这类编号，可以精确更新同一任务。

### 生成和发送群聊周报

预览指定周期的汇总可以使用以下命令。

```bash
.venv/bin/python -m meeting_agent report --project portal --name '客户门户项目' \
  --start 2026-11-23 --end 2026-11-27
```

如需把通知发到真实 Slack 群，在运行环境中配置 `MEETING_SLACK_WEBHOOK` 和目标群的 `MEETING_CHANNEL_ID`，再运行 `.venv/bin/python -m meeting_agent run --publish`。也可用以下命令手动发布，默认取当前日期及前六天。 [Slack Incoming Webhook 文档](https://docs.slack.dev/messaging/sending-messages-using-incoming-webhooks/)

```bash
.venv/bin/python -m meeting_agent report --project portal --name '客户门户项目' \
  --timezone Asia/Taipei --publish
```

给 `MEETING_DOCS_BASE_URL` 配置团队文件库的 HTTPS 目录地址，周报会引用该目录下的 `<project>/<meeting>.docx`。部署时用全局 `--documents` 指向同步目录，文件访问权限由团队文件库管理。`--db` 和 `--documents` 放在 `ingest` 或 `report` 之前，同一项目的导入与发布使用同一个数据库。

SQLite 对同一项目、群聊、周期的发送去重。Slack 明确拒绝时可以重跑。遇到超时或异常回执时，等发送进程结束后核对群聊，再按原报告日期运行 `report --delivery-result sent` 或 `report --delivery-result rejected`。第二种情况会释放该次发送记录，之后可以重试。

## 试点验证

### 真实模型验收

配置百炼 API Key 和同地域地址后，从源码目录运行。验收使用当前配置的百炼模型，如需明确指定调用 ID，可以加 `--model`。

```bash
.venv/bin/python -m examples.verify_live
```

验收每次使用新建的数据库，以两份模拟会议文字稿为输入，通过 Agent 平时使用的 `/compatible-mode/v1/chat/completions` 调用百炼。通过需要两次都收到目标模型的响应 ID 和生成 token 数，还要生成两份带原生图表的 Word，并验证同一个行动项从「进行中」变为「已完成」，周报统计随之更新。

每次运行在 `data/live-acceptance/` 下创建独立目录。`acceptance.json` 记录请求与返回的模型 ID、响应 ID、耗时、token 数、每项检查结果，以及输入与产物的 SHA-256。两份 Word 和周报预览也保存在同一目录。失败时命令退出码为 1，记录会保存已获得的回执或具体报错。

报告用百炼返回的 `model` 字段核对模型调用 ID。验收专属 32B 时，把 `acceptance.json` 和百炼控制台里该调用 ID 对应的基础模型配置一并保留。

### 回放实际会议稿

离线演示和自动测试覆盖文件监听、会议修订和 Word 重建、同名任务、跨周待确认事项、原文校验、原生图表、周报和发送去重。进一步用经过人工核对的真实文字稿衡量决定和行动项召回率、负责人和日期准确率、提议误判率及模型调用耗时。
