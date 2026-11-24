import json
import re
import sqlite3
import unicodedata
from collections import Counter, defaultdict, deque
from datetime import date
from pathlib import Path

from .extract import Item, STATUS


SCHEMA = """
CREATE TABLE IF NOT EXISTS meetings (
    seq INTEGER PRIMARY KEY, project TEXT NOT NULL, meeting_id TEXT NOT NULL,
    day TEXT NOT NULL, title TEXT NOT NULL, fingerprint TEXT NOT NULL,
    items TEXT NOT NULL, revision INTEGER NOT NULL DEFAULT 1, source TEXT, trace TEXT,
    UNIQUE(project, meeting_id)
);
CREATE INDEX IF NOT EXISTS project_day ON meetings(project, day, seq);
CREATE TABLE IF NOT EXISTS meeting_revisions (
    project TEXT NOT NULL, meeting_id TEXT NOT NULL, revision INTEGER NOT NULL,
    record TEXT NOT NULL, PRIMARY KEY(project, meeting_id, revision)
);
CREATE TABLE IF NOT EXISTS document_jobs (
    project TEXT NOT NULL, meeting_id TEXT NOT NULL, PRIMARY KEY(project, meeting_id)
);
CREATE TABLE IF NOT EXISTS inbox_runs (
    path TEXT PRIMARY KEY, fingerprint TEXT NOT NULL, status TEXT NOT NULL, message TEXT NOT NULL,
    attempts INTEGER NOT NULL, retry_at REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS deliveries (
    project TEXT NOT NULL, channel TEXT NOT NULL, start TEXT NOT NULL, end TEXT NOT NULL,
    state TEXT NOT NULL, PRIMARY KEY(project, channel, start, end)
);
"""

FIELDS = ("owner", "due", "status")
TERMINAL = {"done", "cancelled"}


def connect(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(path, timeout=30)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA journal_mode = WAL")
    db.executescript(SCHEMA)
    with db:
        db.execute("BEGIN IMMEDIATE")
        columns = {row["name"] for row in db.execute("PRAGMA table_info(meetings)")}
        for column, definition in (("revision", "INTEGER NOT NULL DEFAULT 1"), ("source", "TEXT"), ("trace", "TEXT")):
            if column not in columns:
                db.execute(f"ALTER TABLE meetings ADD COLUMN {column} {definition}")
        if "revision" not in columns:
            db.execute("INSERT OR IGNORE INTO document_jobs SELECT project, meeting_id FROM meetings")
    return db


def meetings(db: sqlite3.Connection, project: str, through: str = "9999-12-31") -> list[dict]:
    rows = db.execute("SELECT seq, meeting_id, day, title, items, revision FROM meetings "
                      "WHERE project = ? AND day <= ? ORDER BY day, seq", (project, through))
    return [{**row, "items": json.loads(row["items"])} for row in rows]


def update_action(state: dict, event: dict) -> dict:
    previous = state.get(event["action_id"], {})
    action = {**event, "text": previous.get("text", event["text"])}
    for field in FIELDS:
        if event[field] is None:
            action[field] = previous.get(field)
    state[event["action_id"]] = action
    return action


def actions_as_of(records: list[dict]) -> dict:
    actions = {}
    for record in records:
        for event in record["items"]:
            if event["kind"] == "action":
                update_action(actions, event)
    return actions


def normalize(text: str) -> str:
    return re.sub(r"[\W_]+", "", unicodedata.normalize("NFKC", text).casefold())


def question_key(record: dict, event: dict, index: int) -> str:
    return event.get("question_id") or f"{record['meeting_id']}-{index}"


def open_questions(records: list[dict]) -> dict:
    questions = {}
    for record in records:
        for index, event in enumerate(record["items"]):
            if event["kind"] != "question":
                continue
            key = question_key(record, event, index)
            if event.get("operation") == "close":
                questions.pop(key, None)
            else:
                questions[key] = {**event, "day": record["day"], "meeting_id": record["meeting_id"]}
    return questions


def identifiers(records: list[dict]) -> set[str]:
    return {item[field] for record in records for item in record["items"]
            for field in ("action_id", "question_id") if field in item}


def snapshot_items(records: list[dict]) -> None:
    actions = {}
    for record in records:
        for event in record["items"]:
            if event["kind"] == "action":
                action = update_action(actions, event)
                event["snapshot"] = {field: action[field] for field in FIELDS}


def resolve(items: list[Item], as_of: dict, questions: dict, reserved: set[str], previous: list[dict]) -> list[dict]:
    used = {prefix: max((int(code[1:]) for code in reserved if code.startswith(prefix)), default=0) for prefix in "AQ"}
    reusable = defaultdict(deque)
    reused = set()
    previous_actions = {event["action_id"] for event in previous if "action_id" in event}
    for event in previous:
        for field in ("action_id", "question_id"):
            if field in event and event[field] not in as_of and event.get("operation") not in {"close", "update"}:
                key = (event["kind"], normalize(event["text"]))
                reusable[key].append(event[field])

    def allocate(prefix: str, event: dict) -> str:
        available = reusable[(event["kind"], normalize(event["text"]))]
        while available:
            code = available.popleft()
            if code not in reused:
                return code
        used[prefix] += 1
        return f"{prefix}{used[prefix]:04d}"

    resolved = []
    for item in items:
        event = item.model_dump()
        if item.kind == "action":
            explicit = set(re.findall(r"(?<![a-z0-9])A[0-9]{4,}(?![a-z0-9])", item.quote.upper(), re.I))
            matches = [code for code, active in as_of.items()
                       if active["status"] not in TERMINAL and normalize(active["text"]) == normalize(item.text)
                       and (item.owner is None or active["owner"] in {None, item.owner})]
            action_id = None
            if explicit:
                if len(explicit) == 1 and explicit <= (as_of.keys() | previous_actions):
                    action_id = next(iter(explicit))
            elif item.operation == "update" or (item.operation is None and item.status in {"done", "blocked"} and matches):
                if len(matches) == 1:
                    action_id = matches[0]
            else:
                action_id = allocate("A", event)
            if action_id is None:
                event.update(kind="question", text=f"确认行动项归属：{item.text}", owner=None, due=None,
                             status=None, operation="create")
            else:
                reused.add(action_id)
                event["operation"] = "update" if action_id in as_of else "create"
                event.update(action_id=action_id, text=as_of.get(action_id, {}).get("text", item.text))
                action = update_action(as_of, event)
                event["snapshot"] = {field: action[field] for field in FIELDS}
        if event["kind"] == "question":
            explicit = set(re.findall(r"(?<![a-z0-9])Q[0-9]{4,}(?![a-z0-9])", item.quote.upper(), re.I))
            if len(explicit) == 1 and explicit <= questions.keys():
                key = next(iter(explicit))
                if event["operation"] == "close":
                    questions.pop(key)
                else:
                    event["operation"] = "update"
                    questions[key] = event
            else:
                if event["operation"] in {"close", "update"}:
                    event.update(operation="create", text=f"确认待办问题的归属：{event['text']}")
                key = allocate("Q", event)
                questions[key] = event
            event["question_id"] = key
            reused.add(key)
        resolved.append(event)
    return resolved


def report_data(records: list[dict], start: date, end: date) -> dict:
    prior, current = {}, {}
    changed, decisions, covering = {}, [], []
    first, last = start.isoformat(), end.isoformat()
    for record in records:
        if record["day"] > last:
            break
        in_period = record["day"] >= first
        if in_period:
            covering.append(record)
        for event in record["items"]:
            kind = event["kind"]
            if kind == "action":
                key = event["action_id"]
                old = current.get(key)
                new = update_action(current, event)
                if in_period and (old is None or any(old[field] != new[field] for field in FIELDS)):
                    changed[key] = new
                if not in_period:
                    prior[key] = new
            elif in_period and kind == "decision":
                decisions.append((record["day"], event["text"]))

    active = [item for item in current.values() if item["status"] not in TERMINAL]
    overdue = [item for item in active if item["due"] and item["due"] < last]
    blocked = [item for item in active if item["status"] == "blocked"]
    done = [item for key, item in current.items() if item["status"] == "done"
            and prior.get(key, {}).get("status") != "done"]
    active.sort(key=lambda item: (item["status"] != "blocked", item["due"] or "9999-12-31", item["action_id"]))
    outstanding = open_questions([record for record in records if record["day"] <= last])
    questions = [(item["day"], f"{item['question_id']} {item['text']}" if "question_id" in item else item["text"])
                 for item in outstanding.values()]
    return dict(
        covering=covering, decisions=decisions, questions=questions, changed=list(changed.values()), active=active,
        totals=dict(new=len(current.keys() - prior.keys()), done=len(done), overdue=len(overdue), blocked=len(blocked)),
    )


def status_counts(items: list[dict]) -> list[tuple[str, int]]:
    counts = Counter(item["snapshot"]["status"] for item in items)
    return [(label, counts[status]) for status, label in STATUS.items() if counts[status]]
