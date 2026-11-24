import io
import json
import unittest
from contextlib import closing, redirect_stdout
from datetime import date, datetime, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import Mock, patch
from zoneinfo import ZoneInfo

from docx import Document

from meeting_agent.app import ingest, main
from meeting_agent.extract import TransientInferenceError, extract, relevant_candidates
from meeting_agent.ledger import actions_as_of, connect, meetings, report_data
from meeting_agent.worker import last_week_end, run_cycle


DAY = date(2026, 11, 23)
WEEK_END = date(2026, 11, 27)
NOW = datetime(2026, 11, 27, 18, tzinfo=ZoneInfo("Asia/Taipei"))


def item(kind, text, quote, *, owner=None, due=None, status=None, operation=None):
    return {"kind": kind, "text": text, "paragraph": 1, "quote": quote,
            "owner": owner, "due": due, "status": status, "operation": operation}


def response(*items):
    return {"items": list(items)}


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(TemporaryDirectory()))
        self.db = self.enterContext(closing(connect(self.root / "db.sqlite")))
        self.docs = self.root / "documents"

    def save(self, key, day, *items):
        path = self.root / f"{key}.txt"
        path.write_text("\n".join(dict.fromkeys(entry["quote"] for entry in items)), encoding="utf-8")
        args = SimpleNamespace(project="portal", meeting=key, title=key, date=day, transcript=path)
        for entry in items:
            entry["paragraph"] = path.read_text(encoding="utf-8").splitlines().index(entry["quote"]) + 1
        return ingest(self.db, self.docs, args, lambda *_: response(*items))

    def test_revision_replays_later_documents_and_archives_original(self):
        action = item("action", "核对账单", "小陈负责核对账单，截止日为 2026-11-25。", owner="小陈",
                      due="2026-11-25", status="doing", operation="create")
        self.save("m1", DAY, action)
        complete = item("action", "核对账单", "A0001 已完成。", status="done", operation="update")
        later_doc = self.save("m2", WEEK_END, complete)
        revised = {**action, "quote": "小陈负责核对账单，截止日为 2026-11-26。", "due": "2026-11-26"}
        self.save("m1", DAY, revised)
        records = meetings(self.db, "portal")
        self.assertEqual([record["revision"] for record in records], [2, 1])
        tasks = actions_as_of(records)
        self.assertEqual(list(tasks), ["A0001"])
        self.assertEqual(tasks["A0001"]["due"], "2026-11-26")
        self.assertEqual(tasks["A0001"]["status"], "done")
        self.assertEqual(Document(later_doc).tables[0].cell(1, 2).text, "2026-11-26")
        archived = self.db.execute("SELECT record FROM meeting_revisions").fetchone()[0]
        self.assertIn("2026-11-25", json.loads(archived)["source"])

    def test_same_name_creates_separate_tasks_and_unclear_update_stays_pending(self):
        for index, month in enumerate(("本月", "下月"), start=1):
            self.save(f"m{index}", DAY, item("action", "审核月报", f"小陈负责审核{month}月报。",
                                             owner="小陈", status="doing", operation="create"))
        self.save("m3", WEEK_END, item("action", "审核月报", "月报已审核完成。", status="done", operation="update"))
        records = meetings(self.db, "portal")
        tasks = actions_as_of(records)
        self.assertEqual(set(tasks), {"A0001", "A0002"})
        self.assertTrue(all(task["status"] == "doing" for task in tasks.values()))
        self.assertEqual(records[-1]["items"][0]["kind"], "question")
        self.save("m4", WEEK_END, item("action", "审核月报", "A0002 已完成。", status="done", operation="update"))
        tasks = actions_as_of(meetings(self.db, "portal"))
        self.assertEqual([tasks[key]["status"] for key in sorted(tasks)], ["doing", "done"])

    def test_questions_carry_over_until_closed_by_id(self):
        self.save("old", DAY, item("question", "最终上线日期？", "上线日期还没定。"))
        next_week = DAY + timedelta(days=7)
        report = report_data(meetings(self.db, "portal"), next_week, next_week)
        self.assertEqual(report["questions"], [(str(DAY), "Q0001 最终上线日期？")])
        self.save("close", next_week, item("question", "已约定日期", "Q0001 已确认上线日期。", operation="close"))
        report = report_data(meetings(self.db, "portal"), next_week, next_week)
        self.assertEqual(report["questions"], [])

    def test_tentative_decision_future_completion_and_date_mismatch_are_reviewed(self):
        cases = (
            item("decision", "正式上线", "建议周五上线，具体时间待定。"),
            item("action", "提交方案", "预计明天完成方案，尚未提交。", status="done"),
            item("action", "确认日期", "截止日为 2026-11-26。", due="2026-11-25"),
        )
        for entry in cases:
            with self.subTest(quote=entry["quote"]):
                extracted = extract([entry["quote"]], DAY, [], lambda *_: response(entry))
                self.assertEqual(extracted[0].kind, "question")
                self.assertIn(entry["quote"], extracted[0].text)

    def test_retrieval_prefers_matching_task_and_explicit_id(self):
        actions = [{"编号": f"A{n:04}", "任务": f"背景任务 {n}", "未结": True} for n in range(3, 60)]
        actions.extend([{"编号": "A0001", "任务": "配置错误率告警", "未结": True},
                        {"编号": "A0002", "任务": "整理审批权限", "未结": False}])
        selected = relevant_candidates([(1, "配置错误率告警已完成，另请回顾 A0002。")], actions)
        self.assertEqual([entry["编号"] for entry in selected[:2]], ["A0002", "A0001"])
        self.assertEqual(len(selected), 12)

    def test_pending_document_is_recovered_without_new_cloud_call(self):
        entry = item("topic", "说明发布节奏", "讨论了两阶段发布节奏。")
        with patch("meeting_agent.app.write_word", side_effect=OSError("disk busy")):
            with self.assertRaises(OSError):
                self.save("m1", DAY, entry)
        self.assertEqual(self.db.execute("SELECT count(*) FROM document_jobs").fetchone()[0], 1)
        with patch("meeting_agent.app.extract", side_effect=AssertionError("unexpected cloud call")):
            path = self.save("m1", DAY, entry)
        self.assertTrue(path.is_file())
        self.assertEqual(self.db.execute("SELECT count(*) FROM document_jobs").fetchone()[0], 0)

    def worker_args(self):
        args = SimpleNamespace(inbox=self.root / "inbox", reports=self.root / "reports", documents=self.docs,
                               publish=False, once=True)
        (args.inbox / "portal").mkdir(parents=True)
        return args

    def test_workflow_ingests_files_and_schedules_local_group_report(self):
        args = self.worker_args()
        file = args.inbox / "portal" / "2026-11-23__m1__上线评审.txt"
        quote = "团队决定周五向首批客户开放。"
        file.write_text(quote, encoding="utf-8")
        infer = Mock(return_value=response(item("decision", "周五开放", quote)))
        with redirect_stdout(io.StringIO()):
            self.assertEqual(run_cycle(self.db, args, NOW, infer), 0)
            self.assertEqual(run_cycle(self.db, args, NOW, infer), 0)
        infer.assert_called_once()
        self.assertTrue((self.docs / "portal/m1.docx").is_file())
        weekly = (args.reports / "portal/2026-11-27.md").read_text(encoding="utf-8")
        self.assertIn("周五开放", weekly)
        self.assertIn("../../documents/portal/m1.docx", weekly)
        self.assertTrue((args.reports / "portal/latest.md").is_file())
        self.assertEqual(last_week_end(NOW.replace(hour=16)), date(2026, 11, 20))

    def test_workflow_retries_transient_failure_with_backoff(self):
        args = self.worker_args()
        file = args.inbox / "portal" / "2026-11-23__m1.txt"
        source = "确认排期。"
        file.write_text(source, encoding="utf-8")
        infer = Mock(side_effect=[TransientInferenceError("timeout"), response(item("decision", "排期已确认", source))])
        with redirect_stdout(io.StringIO()):
            self.assertEqual(run_cycle(self.db, args, NOW, infer), 1)
            self.assertEqual(run_cycle(self.db, args, NOW + timedelta(seconds=5), infer), 1)
            self.assertEqual(run_cycle(self.db, args, NOW + timedelta(seconds=31), infer), 0)
        self.assertEqual(infer.call_count, 2)
        self.assertTrue((self.docs / "portal/m1.docx").is_file())

    def test_cli_run_uses_only_environment_key_for_local_workflow(self):
        args = self.worker_args()
        source = "团队决定下周上线。"
        (args.inbox / "portal/2026-11-23__m1.txt").write_text(source, encoding="utf-8")
        output = json.dumps(response(item("decision", "下周上线", source)), ensure_ascii=False)
        reply = json.dumps({"id": "test-response", "model": "qwen3.8-27b", "usage": {"completion_tokens": 31},
                            "choices": [{"finish_reason": "stop", "message": {"content": output}}]}).encode()
        cli = ["--db", str(self.root / "run.db"), "--documents", str(self.docs),
               "run", "--once", "--inbox", str(args.inbox), "--reports", str(args.reports)]
        with (redirect_stdout(io.StringIO()), patch.dict("os.environ", {"DASHSCOPE_API_KEY": "test-key"}, clear=True),
              patch("meeting_agent.extract.urlopen", return_value=io.BytesIO(reply)) as cloud,
              patch("meeting_agent.worker.datetime") as clock):
            clock.now.return_value = NOW
            self.assertEqual(main(cli), 0)
            cloud.assert_called_once()
        with closing(connect(self.root / "run.db")) as db:
            trace = json.loads(db.execute("SELECT trace FROM meetings").fetchone()[0])
            self.assertEqual(trace[0]["response_id"], "test-response")
        self.assertIn("下周上线", (args.reports / "portal/latest.md").read_text(encoding="utf-8"))
