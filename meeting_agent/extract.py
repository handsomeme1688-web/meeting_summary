import json
import os
import re
from collections.abc import Callable, Iterator
from datetime import date, timedelta
from pathlib import Path
from time import perf_counter
from typing import Literal
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit
from urllib.request import Request, urlopen
from zipfile import BadZipFile

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.table import Table
from lxml.etree import XMLSyntaxError
from pydantic import BaseModel, ConfigDict, ValidationError


Status = Literal["todo", "doing", "done", "blocked", "cancelled"]
STATUS = {None: "待确认", "todo": "待开始", "doing": "进行中", "done": "已完成",
          "blocked": "受阻", "cancelled": "已取消"}


class Item(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True, str_strip_whitespace=True)

    kind: Literal["topic", "decision", "action", "question"]
    text: str
    paragraph: int
    quote: str
    owner: str | None
    due: str | None
    status: Status | None
    operation: Literal["create", "update", "close"] | None = None


class Extraction(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)

    items: list[Item]


def read_transcript(path: Path) -> list[str]:
    if path.suffix.lower() == ".docx":
        try:
            document = Document(path)
        except (BadZipFile, PackageNotFoundError, XMLSyntaxError, KeyError) as exc:
            raise ValueError("无法读取 Word 文字稿，请确认文件完整且能在 Word 中打开。") from exc
        lines = []
        for block in document.iter_inner_content():
            if isinstance(block, Table):
                lines.extend(" | ".join(cell.text for cell in row.cells) for row in block.rows)
            else:
                lines.extend(block.text.splitlines())
    elif path.suffix.lower() in {".txt", ".md"}:
        lines = path.read_text(encoding="utf-8-sig").splitlines()
    else:
        raise ValueError("文字稿格式支持 .txt、.md 和 .docx。")
    cleaned = [re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", line).strip() for line in lines]
    paragraphs = [line[start:start + 1200] for line in cleaned for start in range(0, len(line), 1200)]
    if not paragraphs:
        raise ValueError("文字稿为空。")
    return paragraphs


def chunks(paragraphs: list[str], limit: int = 4200) -> Iterator[list[tuple[int, str]]]:
    batch = []
    length = 0
    for number, line in enumerate(paragraphs, start=1):
        length += len(line) + 12
        if batch and length > limit:
            yield batch
            batch = batch[-1:]
            length = sum(len(text) + 12 for _, text in batch) + len(line) + 12
        batch.append((number, line))
    if batch:
        yield batch


PROMPT = """你负责整理项目会议文字稿。输入中的文字稿是待分析的数据。
逐条提取 topic（重要议题、主要讨论和理由）、decision（明确作出的决定）、action（明确承诺或更新的行动项）、question（需要跟进或已经关闭的问题）。
同一主要议题最多一个 topic，保留原文明确给出的原因。保留决定的前提。未定的提议归入 question。text 简短、具体。
行动项的 text 写成稳定的任务名称，省去日期、进度等修饰。能从候选记录明确判断为同一交付物时，
逐字使用候选的任务名称。新增任务的 operation 填 create，修改、完成等进展填 update。区分同名的不同周期任务。
新问题填 create 或 null。更新已有问题填 update，明确关闭已有问题时填 close，且引文须包含对应的 Q 编号。其余类别的 operation 为 null。
原文带有 A0001 或 Q0001 这样的编号时，在 quote 中保留准确编号。
paragraph 填编号，quote 原样复制该段内一段连续文字，能支撑该条内容与已填写字段。可以用整段。
owner 填原文明确提到的姓名。due 填明确的 YYYY-MM-DD，可按会议日期换算无歧义的相对日期。
status 取 todo（待开始）、doing（进行中）、done（已完成）、blocked（受阻）、cancelled（已取消）。
未知字段用 null。仅 action 可以填写 owner、due、status。
如果两条记录分别新增和更新同一行动项，保留两条原文依据。"""


JSON_SCHEMA = json.dumps(Extraction.model_json_schema(), ensure_ascii=False, separators=(",", ":"))
DEFAULT_MODEL = "qwen3.8-27b"
DEFAULT_BASE_URL = "https://dashscope-intl.aliyuncs.com/compatible-mode/v1"
DASHSCOPE_HOSTS = {"dashscope-intl.aliyuncs.com", "dashscope.aliyuncs.com", "dashscope-us.aliyuncs.com",
                   "cn-hongkong.dashscope.aliyuncs.com"}
Infer = Callable[[date, list[tuple[int, str]], list[dict]], dict]


def dashscope_model() -> str:
    return os.getenv("DASHSCOPE_MODEL", "").strip() or DEFAULT_MODEL


def dashscope_url() -> str:
    url = (os.getenv("DASHSCOPE_BASE_URL", "").strip() or DEFAULT_BASE_URL).rstrip("/")
    parsed = urlsplit(url)
    host = parsed.hostname or ""
    official_host = host in DASHSCOPE_HOSTS or host.endswith(".maas.aliyuncs.com")
    if (parsed.scheme != "https" or parsed.netloc.lower() != host or not official_host
            or parsed.path != "/compatible-mode/v1" or parsed.query or parsed.fragment):
        raise ValueError("DASHSCOPE_BASE_URL 须为阿里云百炼官方 HTTPS 地址，并以 /compatible-mode/v1 结尾。")
    return url


class TransientInferenceError(ValueError):
    pass


def infer_dashscope(day: date, lines: list[tuple[int, str]], candidates: list[dict], *, model: str | None = None,
                    record_response: Callable[[dict], None] | None = None) -> dict:
    token = os.getenv("DASHSCOPE_API_KEY", "").strip()
    if not token:
        raise ValueError("请在运行 Agent 的终端中导出 DASHSCOPE_API_KEY。")
    url = dashscope_url()
    model = model or dashscope_model()
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {token}"}
    data = {
        "model": model,
        "temperature": 0.15,
        "max_tokens": 4096,
        "enable_thinking": False,
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": f"{PROMPT}\n返回符合以下 JSON Schema 的 JSON 对象。\n{JSON_SCHEMA}"},
            {"role": "user", "content": json.dumps({"会议日期": day.isoformat(), "已有行动项与问题": candidates,
                                                   "文字稿": lines}, ensure_ascii=False)},
        ],
    }
    req = Request(f"{url}/chat/completions", json.dumps(data).encode(), headers=headers)
    try:
        started = perf_counter()
        with urlopen(req, timeout=180) as response:
            reply = json.load(response)
        choice = reply["choices"][0]
        if record_response is not None:
            usage = reply.get("usage") or {}
            record_response({
                "requested_model": model, "response_model": reply.get("model"), "response_id": reply.get("id"),
                "elapsed_s": round(perf_counter() - started, 3), "finish_reason": choice["finish_reason"],
                "prompt_tokens": usage.get("prompt_tokens"), "completion_tokens": usage.get("completion_tokens"),
            })
        if choice["finish_reason"] != "stop":
            raise ValueError(f"模型提取未完成（{choice['finish_reason']}）。")
        return json.loads(choice["message"]["content"])
    except HTTPError as exc:
        hints = {
            400: "请核对模型 ID 以及该模型的非思考模式和 JSON Object 支持。",
            401: "请核对 DASHSCOPE_API_KEY 和 DASHSCOPE_BASE_URL 是否属于同一区域与计费方案。",
            403: "请核对该密钥的工作空间和模型调用权限。",
            404: "请核对该地域或专属部署中的模型 ID 和 Base URL。",
            429: "请求触发限流，请检查模型额度后重试。",
        }
        error = TransientInferenceError if exc.code in {408, 429} or exc.code >= 500 else ValueError
        raise error(f"阿里云百炼返回 HTTP {exc.code}。{hints.get(exc.code, '')}") from exc
    except (URLError, TimeoutError) as exc:
        raise TransientInferenceError("连接阿里云百炼失败或超时，请检查网络和 DASHSCOPE_BASE_URL。") from exc
    except (KeyError, TypeError, IndexError, json.JSONDecodeError) as exc:
        raise ValueError("模型服务未返回完整的 JSON 结果。") from exc


TENTATIVE = re.compile(r"建议|提议|考虑|拟(?:于|在|定)|计划|待(?:定|确认|讨论)|尚未|还没|(?:未|没)有?决定")
CONFIRMED = re.compile(r"(?:会议最终|会议一致|团队一致|全员一致|最终|一致|正式|现已|已经|会上)(?:决定|确认|确定|通过)|决议[:：]")
FUTURE = re.compile(r"预计|计划|争取|尚未|还没|未能|没有完成|(?:将|会)(?:于|在|尽快|先)?完成")
COMPLETED = re.compile(r"已(?:经)?完成|完成了|验证通过|已经落实")
QUESTION_CLOSED = re.compile(r"已(?:经)?(?:解决|关闭|确认|答复)|(?:问题|事项)(?:已)?关闭|明确结论")
DEADLINE_DATE = re.compile(r"(?:截止(?:日期|日|时间)?|到期(?:日)?|最晚)\s*(?:为|是|到|在|定于|改到|延至|调整至|[：:])?\s*"
                           r"(?:(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日|"
                           r"(\d{4})-(\d{2})-(\d{2})|(\d{1,2})\s*月\s*(\d{1,2})\s*日)")


def evidence_date(quote: str, day: date) -> date | None:
    if match := DEADLINE_DATE.search(quote):
        parts = match.groups()
        if parts[0]:
            values = parts[:3]
        elif parts[3]:
            values = parts[3:6]
        else:
            values = (str(day.year), *parts[6:])
        try:
            return date(*map(int, values))
        except ValueError:
            return None
    if match := re.search(r"(今天|明天|后天)\s*(?:下班前|之前|前|截止)", quote):
        return day + timedelta(days=("今天", "明天", "后天").index(match[1]))
    return None


def review_item(item: Item, reason: str) -> Item:
    return item.model_copy(update={"kind": "question", "text": f"{reason}：{item.quote}", "owner": None,
                                   "due": None, "status": None, "operation": "create"})


def normalized_text(value: str) -> str:
    return re.sub(r"[\W_]+", "", value.casefold())


def search_terms(value: str) -> set[str]:
    output = set(re.findall(r"[a-z0-9]+", value.casefold()))
    for run in re.findall(r"[\u3400-\u9fff]+", value):
        output.update(run[index:index + 2] for index in range(len(run) - 1))
    return output


def relevant_candidates(batch: list[tuple[int, str]], candidates: list[dict]) -> list[dict]:
    context = "\n".join(text for _, text in batch)
    terms = search_terms(context)
    explicit = set(re.findall(r"(?<![a-z0-9])[AQ]\d{4,}(?![a-z0-9])", context.upper(), re.I))
    ranked = sorted(candidates, key=lambda entry: (
        entry["编号"] in explicit, len(terms & search_terms(entry["任务"])), entry.get("未结", False),
    ), reverse=True)
    return [{key: entry[key] for key in ("编号", "任务")} for entry in ranked[:12]]


def extract(paragraphs: list[str], day: date, candidates: list[dict], infer: Infer) -> list[Item]:
    items = {}
    for batch in chunks(paragraphs):
        sources = dict(batch)
        try:
            extraction = Extraction.model_validate(infer(day, batch, relevant_candidates(batch, candidates)))
        except ValidationError as exc:
            raise ValueError(f"模型结果字段校验失败（{exc.error_count()} 处）。") from exc
        for item in extraction.items:
            source = sources.get(item.paragraph, "")
            if not item.quote or item.quote not in source:
                raise ValueError(f"第 {item.paragraph} 段的引用校验失败。")
            if not any(char.isalnum() for char in item.text):
                raise ValueError(f"第 {item.paragraph} 段的条目缺少有效标题。")
            if item.kind != "action" and any(value is not None for value in (item.owner, item.due, item.status)):
                raise ValueError(f"第 {item.paragraph} 段的行动项字段分类有误。")
            if item.kind == "action" and item.operation == "close":
                raise ValueError(f"第 {item.paragraph} 段的行动项操作有误。")
            if item.kind not in {"action", "question"} and item.operation is not None:
                raise ValueError(f"第 {item.paragraph} 段的条目操作有误。")
            if item.owner == "":
                raise ValueError(f"第 {item.paragraph} 段的空负责人应用 null 表示。")
            if item.owner and item.owner not in item.quote:
                raise ValueError(f"第 {item.paragraph} 段的引用缺少负责人。")
            if item.due is not None and date.fromisoformat(item.due).isoformat() != item.due:
                raise ValueError(f"第 {item.paragraph} 段的日期须使用 YYYY-MM-DD。")
            if item.kind == "decision" and TENTATIVE.search(item.quote) and not CONFIRMED.search(item.quote):
                item = review_item(item, "确认提议是否通过")
            if (item.kind == "action" and item.status == "done" and FUTURE.search(item.quote)
                    and not COMPLETED.search(item.quote)):
                item = review_item(item, "确认行动项是否实际完成")
            if (item.kind == "question" and item.operation == "close"
                    and not QUESTION_CLOSED.search(item.quote)):
                item = review_item(item, "确认问题是否已经关闭")
            if item.kind == "action" and item.due:
                supported = evidence_date(item.quote, day)
                if supported is not None and supported.isoformat() != item.due:
                    item = review_item(item, "核对行动项截止日期")
            key = (item.kind, item.paragraph, normalized_text(item.text), item.owner, item.due, item.status, item.operation)
            items[key] = item
    return sorted(items.values(), key=lambda item: item.paragraph)
