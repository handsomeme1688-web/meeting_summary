import os
import re
from contextlib import closing
from datetime import date, datetime, timedelta
from hashlib import sha256
from pathlib import Path
from tempfile import NamedTemporaryFile
from time import sleep
from types import SimpleNamespace
from urllib.parse import quote
from zoneinfo import ZoneInfo

from .app import format_report, ingest, publish, sync_documents, validate_docs_base_url
from .extract import TransientInferenceError, infer_dashscope
from .ledger import connect, meetings, report_data


NAME = re.compile(r"(\d{4}-\d{2}-\d{2})__([a-zA-Z0-9][a-zA-Z0-9_-]{0,79}?)(?:__(.+))?")
PROJECT = re.compile(r"[a-zA-Z0-9][a-zA-Z0-9_-]{0,79}")


def source_args(inbox: Path, path: Path) -> SimpleNamespace:
    relative = path.relative_to(inbox)
    match = NAME.fullmatch(path.stem)
    project = relative.parts[0] if len(relative.parts) == 2 else "general"
    if not match or not PROJECT.fullmatch(project):
        raise ValueError("请使用 项目编号/YYYY-MM-DD__会议编号__标题.txt 命名。")
    day, meeting, title = match.groups()
    return SimpleNamespace(project=project, meeting=meeting, date=date.fromisoformat(day),
                           title=title or meeting, transcript=path)


def atomic_text(path: Path, content: str) -> None:
    if path.is_file() and path.read_text(encoding="utf-8") == content:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=path.parent, mode="w", encoding="utf-8", delete=False) as handle:
        temporary = Path(handle.name)
        handle.write(content)
    try:
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def local_report(path: Path, project: str, docs: Path, start: date, end: date, data: dict) -> None:
    text = format_report(project, project, start, end, data, None)
    links = []
    for record in data["covering"]:
        filename = f"{record['meeting_id']}.docx"
        relative = os.path.relpath(docs / project / filename, path.parent)
        links.append(f"- [{filename}]({quote(Path(relative).as_posix())})")
    if links:
        text += "\n\n## 纪要文件\n\n" + "\n".join(links)
    atomic_text(path, text + "\n")


def last_week_end(now: datetime) -> date:
    friday = now.date() - timedelta(days=(now.weekday() - 4) % 7)
    if friday == now.date() and now.hour < 17:
        friday -= timedelta(days=7)
    return friday


def has_updates(data: dict) -> bool:
    return bool(data["covering"] or data["active"] or data["questions"])


def run_cycle(db, args, now: datetime, infer=infer_dashscope, attempted_deliveries=None, *, retry_failed=False) -> int:
    attempted_deliveries = set() if attempted_deliveries is None else attempted_deliveries
    failures = 0
    sources = sorted((*args.inbox.glob("*"), *args.inbox.glob("*/*")))
    for path in sources:
        if (path.suffix.lower() not in {".txt", ".md", ".docx"} or path.is_symlink()
                or path.parent.is_symlink() or not path.is_file()):
            continue
        try:
            if not args.once and now.timestamp() - path.stat().st_mtime < 2:
                continue
            fingerprint = sha256(path.read_bytes()).hexdigest()
        except FileNotFoundError:
            continue
        except OSError as exc:
            print(f"读取失败 {path.name} · {exc}", flush=True)
            failures += 1
            continue

        key = str(path.absolute())
        previous = db.execute("SELECT * FROM inbox_runs WHERE path = ?", (key,)).fetchone()
        same = previous is not None and previous["fingerprint"] == fingerprint
        attempts = previous["attempts"] if same and not retry_failed else 0
        if same and previous["status"] == "ok":
            continue
        if same and not retry_failed and (previous["status"] == "failed" or now.timestamp() < previous["retry_at"]):
            failures += 1
            continue

        status, message, retry_at = "ok", "", 0.0
        try:
            document = ingest(db, args.documents, source_args(args.inbox, path), infer)
            print(f"纪要已生成 · {document}", flush=True)
        except (OSError, ValueError) as exc:
            failures += 1
            attempts += 1
            status, message = "failed", str(exc)
            if isinstance(exc, (OSError, TransientInferenceError)) and attempts < 4:
                status = "retry"
                retry_at = now.timestamp() + 30 * 2 ** (attempts - 1)
            print(f"{path.name} · {message}", flush=True)
        with db:
            db.execute("INSERT INTO inbox_runs VALUES (?, ?, ?, ?, ?, ?) ON CONFLICT(path) DO UPDATE SET "
                       "fingerprint=excluded.fingerprint, status=excluded.status, message=excluded.message, "
                       "attempts=excluded.attempts, retry_at=excluded.retry_at",
                       (key, fingerprint, status, message, attempts, retry_at))

    end = last_week_end(now)
    start = end - timedelta(days=6)
    projects = [row["project"] for row in db.execute("SELECT DISTINCT project FROM meetings ORDER BY project")]
    for project in projects:
        try:
            sync_documents(db, project, args.documents)
            records = meetings(db, project, now.date().isoformat())
            if not records:
                continue
            preview_end = date.fromisoformat(records[-1]["day"])
            preview_start = preview_end - timedelta(days=6)
            preview = report_data(records, preview_start, preview_end)
            local_report(args.reports / project / "latest.md", project, args.documents, preview_start, preview_end, preview)

            weekly = report_data(records, start, end)
            if not has_updates(weekly):
                continue
            local_report(args.reports / project / f"{end}.md", project, args.documents, start, end, weekly)
            if args.publish:
                channel = os.environ["MEETING_CHANNEL_ID"]
                key = (project, channel, start.isoformat(), end.isoformat())
                if key not in attempted_deliveries:
                    attempted_deliveries.add(key)
                    body = format_report(project, project, start, end, weekly, os.getenv("MEETING_DOCS_BASE_URL"))
                    print(publish(db, key, os.environ["MEETING_SLACK_WEBHOOK"], body), flush=True)
        except (OSError, ValueError) as exc:
            failures += 1
            print(f"{project} · {exc}", flush=True)
    return failures


def run(args) -> int:
    if not os.getenv("DASHSCOPE_API_KEY", "").strip():
        raise ValueError("请在启动 Agent 的终端中导出 DASHSCOPE_API_KEY。")
    if args.interval < 1:
        raise ValueError("监听间隔至少为 1 秒。")
    if args.publish and not all(os.getenv(name) for name in ("MEETING_SLACK_WEBHOOK", "MEETING_CHANNEL_ID")):
        raise ValueError("发送到真实 Slack 群需要 MEETING_SLACK_WEBHOOK 和 MEETING_CHANNEL_ID。")
    validate_docs_base_url(os.getenv("MEETING_DOCS_BASE_URL"))
    zone = ZoneInfo(args.timezone)
    args.inbox.mkdir(parents=True, exist_ok=True)
    args.reports.mkdir(parents=True, exist_ok=True)
    args.documents.mkdir(parents=True, exist_ok=True)
    print(f"文字稿收件箱 · {args.inbox.resolve()}\n本地周报 · {args.reports.resolve()}", flush=True)
    attempted_deliveries = set()
    retry_failed = args.retry_failed
    with closing(connect(args.db)) as db:
        try:
            while True:
                failures = run_cycle(db, args, datetime.now(zone), attempted_deliveries=attempted_deliveries,
                                     retry_failed=retry_failed)
                if args.once:
                    return int(failures > 0)
                retry_failed = False
                sleep(args.interval)
        except KeyboardInterrupt:
            return 0
