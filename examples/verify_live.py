import argparse
import json
from contextlib import closing
from datetime import UTC, date, datetime
from functools import partial
from hashlib import sha256
from pathlib import Path
from types import SimpleNamespace
from uuid import uuid4
from zipfile import BadZipFile, ZipFile

from meeting_agent.app import format_report, ingest
from meeting_agent.extract import dashscope_model, infer_dashscope
from meeting_agent.ledger import actions_as_of, connect, meetings, report_data


SAMPLES = (
    (date(2026, 11, 23), "客户门户上线准备",
     "会议最终决定：11 月 26 日向 10 家白名单客户开放新门户。\n"
     "陈遥：我负责配置错误率告警，截止日为 2026-11-25。目前正在配置。\n"),
    (date(2026, 11, 26), "客户门户进度确认",
     "A0001 已完成。错误率告警已配置并验证通过。\n"),
)
PROJECT = "live-check"


def has_native_chart(path: Path) -> bool:
    with ZipFile(path) as archive:
        parts = set(archive.namelist())
        return {"word/charts/chart1.xml", "word/embeddings/action-status.xlsx"} <= parts


def verify(model: str, root: Path) -> tuple[Path, bool]:
    started = datetime.now(UTC)
    run = root / f"{started:%Y%m%dT%H%M%SZ}-{uuid4().hex[:8]}"
    run.mkdir(parents=True)
    receipts, words, inputs, checks = [], [], [], {}
    infer = partial(infer_dashscope, model=model, record_response=receipts.append)
    error = None
    try:
        with closing(connect(run / "ledger.sqlite3")) as db:
            for day, title, transcript in SAMPLES:
                source = run / f"{day}.txt"
                source.write_text(transcript, encoding="utf-8")
                inputs.append(source)
                args = SimpleNamespace(project=PROJECT, meeting=f"m-{day}", date=day, title=title, transcript=source)
                words.append(ingest(db, run / "documents", args, infer))
            records = meetings(db, PROJECT)

        original = actions_as_of(records[:1])
        current = actions_as_of(records)
        expected = {"owner": "陈遥", "due": "2026-11-25"}
        start, end = SAMPLES[0][0], SAMPLES[-1][0]
        data = report_data(records, start, end)
        checks = {
            "model_and_response_ids": len(receipts) == len(SAMPLES) and all(
                receipt["response_model"] == model and receipt["response_id"] for receipt in receipts
            ),
            "generated_tokens_reported": bool(receipts) and all(
                isinstance(receipt["completion_tokens"], int) and receipt["completion_tokens"] > 0
                for receipt in receipts
            ),
            "decision_extracted": any(item["kind"] == "decision" for item in records[0]["items"]),
            "action_extracted": set(original) == {"A0001"} and original["A0001"]["status"] == "doing"
            and all(original["A0001"][field] == value for field, value in expected.items()),
            "same_action_completed": set(current) == {"A0001"} and current["A0001"]["status"] == "done"
            and all(current["A0001"][field] == value for field, value in expected.items()),
            "two_word_documents_with_native_charts": len(words) == len(SAMPLES) and all(map(has_native_chart, words)),
            "weekly_counts": data["totals"] == {"new": 1, "done": 1, "overdue": 0, "blocked": 0},
        }
        preview = run / "report-preview.txt"
        preview.write_text(format_report(PROJECT, "模型验收", start, end, data, None) + "\n", encoding="utf-8")
    except (BadZipFile, OSError, ValueError) as exc:
        error = str(exc)

    artifacts = [*inputs, *words]
    if (run / "report-preview.txt").is_file():
        artifacts.append(run / "report-preview.txt")
    passed = error is None and bool(checks) and all(checks.values())
    result = {
        "provider": "Alibaba Cloud Model Studio", "started_at_utc": started.isoformat(),
        "requested_model": model, "passed": passed, "error": error,
        "requests": receipts, "checks": checks,
        "artifacts_sha256": {str(path.relative_to(run)): sha256(path.read_bytes()).hexdigest() for path in artifacts},
    }
    proof = run / "acceptance.json"
    proof.write_text(json.dumps(result, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return proof, passed


def main() -> int:
    parser = argparse.ArgumentParser(description="调用已部署的模型，验证会议 Agent 并保存验收记录。")
    parser.add_argument("--model", default=dashscope_model())
    parser.add_argument("--output", type=Path, default=Path("data/live-acceptance"))
    args = parser.parse_args()
    proof, passed = verify(args.model, args.output)
    print(f"{'通过' if passed else '未通过'} · {proof}")
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
