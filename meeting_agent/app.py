import argparse
import hashlib
import json
import os
import re
import sqlite3
from contextlib import closing
from datetime import date, datetime, timedelta
from functools import partial
from html import escape
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlsplit
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from .extract import Infer, STATUS, extract, infer_dashscope, read_transcript
from .ledger import TERMINAL, actions_as_of, connect, identifiers, meetings, open_questions, report_data, resolve, snapshot_items
from .word import write_word


def slug(value: str) -> str:
    if not re.fullmatch(r"[a-zA-Z0-9][a-zA-Z0-9_-]{0,79}", value):
        raise argparse.ArgumentTypeError("编号使用英文、数字、下划线或连字符，最多 80 个字符。")
    return value


def sync_documents(db: sqlite3.Connection, project: str, docs: Path) -> None:
    while True:
        with db:
            db.execute("BEGIN IMMEDIATE")
            pending = db.execute("SELECT m.meeting_id, m.day, m.title, m.items, m.revision FROM document_jobs j "
                                 "JOIN meetings m USING (project, meeting_id) WHERE j.project = ? "
                                 "ORDER BY m.day, m.seq LIMIT 1", (project,)).fetchone()
            if pending is None:
                return
            key = pending["meeting_id"]
            write_word(docs / project / f"{key}.docx", pending["title"], pending["day"], json.loads(pending["items"]),
                       revision=pending["revision"])
            db.execute("DELETE FROM document_jobs WHERE project = ? AND meeting_id = ?", (project, key))


def preceding(records: list[dict], day: str, current: sqlite3.Row | None) -> list[dict]:
    sequence = current["seq"] if current else None
    return [record for record in records if record["seq"] != sequence
            and (record["day"] < day or (record["day"] == day and (sequence is None or record["seq"] < sequence)))]


def ingest(db: sqlite3.Connection, docs: Path, args, infer: Infer) -> Path:
    lines = read_transcript(args.transcript)
    day = args.date.isoformat()
    normalized = json.dumps([day, args.title, lines], ensure_ascii=False).encode()
    fingerprint = hashlib.sha256(normalized).hexdigest()
    path = docs / args.project / f"{args.meeting}.docx"
    find = "SELECT * FROM meetings WHERE project = ? AND meeting_id = ?"
    key = (args.project, args.meeting)
    original = db.execute(find, key).fetchone()
    expected = original["fingerprint"] if original else None
    items = None
    trace = []
    if expected != fingerprint:
        history = preceding(meetings(db, args.project, day), day, original)
        actions = actions_as_of(history)
        questions = open_questions(history)
        candidates = [{"编号": code, "任务": item["text"], "未结": item["status"] not in TERMINAL}
                      for code, item in actions.items()]
        candidates.extend({"编号": code, "任务": item["text"], "未结": True}
                          for code, item in questions.items() if "question_id" in item)
        call = partial(infer, record_response=trace.append) if infer is infer_dashscope else infer
        items = extract(lines, args.date, candidates, call)

    with db:
        db.execute("BEGIN IMMEDIATE")
        prior = db.execute(find, key).fetchone()
        if prior and prior["fingerprint"] == fingerprint:
            if not path.is_file():
                db.execute("INSERT OR IGNORE INTO document_jobs VALUES (?, ?)", key)
        else:
            if (prior["fingerprint"] if prior else None) != expected:
                raise ValueError("同一份会议记录正在修订，请重试。")
            assert items is not None
            history = meetings(db, args.project)
            tickets = db.execute("SELECT record FROM meeting_revisions WHERE project = ?", (args.project,))
            archived = [{"items": json.loads(json.loads(row["record"])["items"])} for row in tickets]
            past = preceding(history, day, prior)
            previous = json.loads(prior["items"]) if prior else []
            resolved = resolve(items, actions_as_of(past), open_questions(past), identifiers(history + archived), previous)
            source = json.dumps({"filename": args.transcript.name, "paragraphs": lines}, ensure_ascii=False)
            trace_json = json.dumps(trace, ensure_ascii=False)
            items_json = json.dumps(resolved, ensure_ascii=False)
            if prior:
                db.execute("INSERT INTO meeting_revisions VALUES (?, ?, ?, ?)",
                           (*key, prior["revision"], json.dumps(dict(prior), ensure_ascii=False)))
                db.execute("UPDATE meetings SET day = ?, title = ?, fingerprint = ?, items = ?, source = ?, trace = ?, "
                           "revision = revision + 1 WHERE project = ? AND meeting_id = ?",
                           (day, args.title, fingerprint, items_json, source, trace_json, *key))
            else:
                db.execute("INSERT INTO meetings(project, meeting_id, day, title, fingerprint, items, source, trace) "
                           "VALUES (?, ?, ?, ?, ?, ?, ?, ?)", (*key, day, args.title, fingerprint, items_json, source, trace_json))
            db.execute("INSERT OR IGNORE INTO document_jobs VALUES (?, ?)", key)
            current = meetings(db, args.project)
            originals = {record["meeting_id"]: json.dumps(record["items"], ensure_ascii=False) for record in current}
            snapshot_items(current)
            for record in current:
                value = json.dumps(record["items"], ensure_ascii=False)
                if value != originals[record["meeting_id"]]:
                    changed_key = (args.project, record["meeting_id"])
                    db.execute("UPDATE meetings SET items = ? WHERE project = ? AND meeting_id = ?", (value, *changed_key))
                    db.execute("INSERT OR IGNORE INTO document_jobs VALUES (?, ?)", changed_key)
    sync_documents(db, args.project, docs)
    return path


def format_report(project: str, name: str, start: date, end: date, data: dict, base_url: str | None) -> str:
    def action(item):
        who = escape(item["owner"] or "负责人待确认")
        due = item["due"] or "截止日待确认"
        return f"{item['action_id']} {escape(item['text'])} · {STATUS[item['status']]} · {who} · {due}"

    totals = data["totals"]
    updated_ids = {item["action_id"] for item in data["changed"]}
    lines = [f"*{escape(name)} 工作进展 · {start} 至 {end}*",
             f"新增 {totals['new']} 项 · 完成 {totals['done']} 项 · 逾期 {totals['overdue']} 项 · 受阻 {totals['blocked']} 项"]
    sections = (
        ("行动项更新", [action(item) for item in data["changed"]]),
        ("往期仍需跟进", [action(item) for item in data["active"] if item["action_id"] not in updated_ids]),
        ("会议决定", [escape(text) for _, text in data["decisions"]]),
        ("待确认问题", [escape(text) for _, text in data["questions"]]),
    )
    for heading, items in sections:
        if items:
            lines.extend(("", f"*{heading}*", *(f"• {item}" for item in items[:6])))
            if len(items) > 6:
                lines.append(f"另有 {len(items) - 6} 项。")
    if data["covering"]:
        latest = data["covering"][-5:]
        lines.extend(("", f"*来源纪要（本期 {len(data['covering'])} 份）*"))
        for record in latest:
            title = escape(record["title"]).replace("|", "｜")
            if base_url:
                url = f"{base_url.rstrip('/')}/{quote(project)}/{quote(record['meeting_id'])}.docx"
                title = f"<{url}|{title}>"
            lines.append(f"• {record['day']} {title}")
    return "\n".join(lines)


class SlackRejected(Exception):
    pass


def post_slack(webhook: str, body: str) -> None:
    payload = json.dumps({"text": body, "unfurl_links": False, "unfurl_media": False}).encode()
    req = Request(webhook, payload, headers={"Content-Type": "application/json"})
    try:
        with urlopen(req, timeout=20) as response:
            if response.status != 200 or response.read(32).strip() != b"ok":
                raise ValueError("Slack 回执异常，发送状态待确认。")
    except HTTPError as exc:
        if exc.code in {400, 401, 403, 404, 410, 429}:
            raise SlackRejected(f"Slack 拒绝了请求（HTTP {exc.code}），可以在修正后重试。") from exc
        raise ValueError(f"Slack 发送状态待确认（HTTP {exc.code}）。") from exc
    except (URLError, TimeoutError) as exc:
        raise ValueError("Slack 发送超时或连接中断，发送状态待确认。") from exc


DELIVERY = "project = ? AND channel = ? AND start = ? AND end = ?"


def publish(db: sqlite3.Connection, key: tuple[str, str, str, str], webhook: str, body: str) -> str:
    target = urlsplit(webhook)
    if target.scheme != "https" or target.netloc not in {"hooks.slack.com", "hooks.slack-gov.com"}:
        raise ValueError("请配置 Slack Incoming Webhook 的 HTTPS 地址。")
    with db:
        db.execute("BEGIN IMMEDIATE")
        existing = db.execute(f"SELECT state FROM deliveries WHERE {DELIVERY}", key).fetchone()
        if existing:
            if existing["state"] == "sent":
                return "这期报告已发送。"
            raise ValueError("这期报告的发送状态待确认。核对群聊后使用 --delivery-result。")
        db.execute("INSERT INTO deliveries VALUES (?, ?, ?, ?, 'sending')", key)

    try:
        post_slack(webhook, body)
    except SlackRejected as exc:
        with db:
            db.execute(f"DELETE FROM deliveries WHERE {DELIVERY}", key)
        raise ValueError(str(exc)) from exc
    with db:
        db.execute(f"UPDATE deliveries SET state = 'sent' WHERE {DELIVERY}", key)
    return "周报已发送到 Slack。"


def settle(db: sqlite3.Connection, key: tuple[str, str, str, str], outcome: str) -> str:
    statement = "UPDATE deliveries SET state = 'sent'" if outcome == "sent" else "DELETE FROM deliveries"
    with db:
        cursor = db.execute(f"{statement} WHERE {DELIVERY} AND state = 'sending'", key)
    if cursor.rowcount == 0:
        raise ValueError("该报告没有待确认的发送记录。")
    return "已确认发送。" if outcome == "sent" else "已清除待确认状态，可以重新发送。"


def validate_docs_base_url(url: str | None) -> None:
    if url:
        base = urlsplit(url)
        if base.scheme != "https" or not base.hostname or base.query or base.fragment or base.username:
            raise ValueError("文档基址使用团队文件库的 HTTPS 目录地址。")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="从会议文字稿生成 Word 和项目周报。")
    parser.add_argument("--db", type=Path, default=Path("data/meetings.sqlite3"))
    parser.add_argument("--documents", type=Path, default=Path("data/documents"))
    commands = parser.add_subparsers(dest="command", required=True)
    meeting = commands.add_parser("ingest", help="导入一份有项目归属的文字稿或工作记录。")
    meeting.add_argument("transcript", type=Path)
    meeting.add_argument("--project", type=slug, required=True)
    meeting.add_argument("--meeting", type=slug, required=True)
    meeting.add_argument("--title", required=True)
    meeting.add_argument("--date", type=date.fromisoformat, required=True)
    meeting.add_argument("--fixture", type=Path, help="用示例 JSON 代替模型响应，仅供离线演示。")
    report = commands.add_parser("report", help="预览或发送项目报告。")
    report.add_argument("--project", type=slug, required=True)
    report.add_argument("--name")
    report.add_argument("--start", type=date.fromisoformat)
    report.add_argument("--end", type=date.fromisoformat)
    report.add_argument("--timezone", default="Asia/Taipei")
    report.add_argument("--docs-base-url", default=os.getenv("MEETING_DOCS_BASE_URL"))
    report.add_argument("--channel-id", default=os.getenv("MEETING_CHANNEL_ID"))
    mode = report.add_mutually_exclusive_group()
    mode.add_argument("--publish", action="store_true", help="实际发送到配置的 Slack 群聊。")
    mode.add_argument("--delivery-result", choices=("sent", "rejected"))
    worker = commands.add_parser("run", help="监听文字稿目录，生成 Word 并按周发布。")
    worker.add_argument("--inbox", type=Path, default=Path("data/inbox"))
    worker.add_argument("--reports", type=Path, default=Path("data/reports"))
    worker.add_argument("--timezone", default="Asia/Taipei")
    worker.add_argument("--interval", type=int, default=10, help="扫描间隔秒数。")
    worker.add_argument("--once", action="store_true", help="扫描一次并退出。")
    worker.add_argument("--retry-failed", action="store_true", help="启动时重试未成功的文字稿。")
    worker.add_argument("--publish", action="store_true", help="每周发布到已配置的 Slack 群聊。")
    args = parser.parse_args(argv)

    try:
        if args.command == "run":
            from .worker import run

            return run(args)
        with closing(connect(args.db)) as db:
            if args.command == "ingest":
                infer = infer_dashscope
                if args.fixture:
                    fixture = json.loads(args.fixture.read_text(encoding="utf-8"))

                    def infer(_day, batch, _candidates):
                        allowed = {number for number, _ in batch}
                        return {"items": [item for item in fixture["items"] if item["paragraph"] in allowed]}

                print(ingest(db, args.documents, args, infer))
                return 0
            end = args.end or datetime.now(ZoneInfo(args.timezone)).date()
            start = args.start or end - timedelta(days=6)
            if start > end:
                raise ValueError("开始日期须早于或等于结束日期。")
            if (args.publish or args.delivery_result) and not args.channel_id:
                raise ValueError("请配置稳定的 Slack 群聊编号 MEETING_CHANNEL_ID。")
            key = (args.project, args.channel_id, start.isoformat(), end.isoformat())
            if args.delivery_result:
                if not args.start or not args.end:
                    raise ValueError("确认发送状态时，请提供原报告的 --start 和 --end。")
                print(settle(db, key, args.delivery_result))
                return 0
            validate_docs_base_url(args.docs_base_url)
            sync_documents(db, args.project, args.documents)
            records = meetings(db, args.project, end.isoformat())
            data = report_data(records, start, end)
            if not data["covering"] and not data["active"] and not data["questions"]:
                print("这段期间暂无项目进展或未结事项。")
                return 0
            body = format_report(args.project, args.name or args.project, start, end, data, args.docs_base_url)
            if args.publish:
                print(publish(db, key, os.getenv("MEETING_SLACK_WEBHOOK", ""), body))
            else:
                print(body)
            return 0
    except (OSError, ValueError, KeyError, sqlite3.Error, ZoneInfoNotFoundError) as exc:
        parser.exit(1, f"{exc}\n")
