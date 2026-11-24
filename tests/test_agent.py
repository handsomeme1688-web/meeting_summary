import json
import unittest
from contextlib import closing
from datetime import date
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch
from zipfile import ZipFile

from lxml import etree

from examples.verify_live import verify
from meeting_agent.app import SlackRejected, format_report, ingest, publish, settle
from meeting_agent.extract import extract, infer_dashscope
from meeting_agent.ledger import actions_as_of, connect, meetings, report_data
from meeting_agent.word import CHART_NS


EXAMPLES = Path(__file__).resolve().parents[1] / "examples"
SLACK = "https://hooks.slack.com/services/test/test/test"
DAY_1 = date(2026, 11, 23)
DAY_2 = date(2026, 11, 26)
END = date(2026, 11, 27)


def fixture(day):
    return json.loads((EXAMPLES / f"{day}.json").read_text(encoding="utf-8"))


class AgentTests(unittest.TestCase):
    def setUp(self):
        temp = self.enterContext(TemporaryDirectory())
        self.root = Path(temp)
        self.db = self.enterContext(closing(connect(self.root / "records.db")))
        self.key = ("portal", "channel", str(DAY_1), str(END))

    def import_day(self, day, infer=None):
        args = SimpleNamespace(project="portal", meeting=f"m-{day}", date=day, title="客户门户",
                               transcript=EXAMPLES / f"{day}.txt")
        return ingest(self.db, self.root / "docs", args, infer or (lambda *_: fixture(day)))

    def test_two_meetings_keep_action_id_and_due_date(self):
        self.import_day(DAY_1)
        self.import_day(DAY_2)
        records = meetings(self.db, "portal")
        actions = actions_as_of(records)
        self.assertEqual(len(actions), 4)
        self.assertEqual(actions["A0002"]["status"], "done")
        self.assertEqual(actions["A0002"]["due"], "2026-11-25")
        self.assertEqual(actions["A0002"]["owner"], "陈遥")
        self.assertEqual(actions["A0003"]["status"], "done")
        data = report_data(records, DAY_1, END)
        self.assertEqual(data["totals"], dict(new=4, done=3, overdue=0, blocked=1))
        self.assertEqual(report_data(records, DAY_1, DAY_1)["totals"]["done"], 1)
        preview = format_report("portal", "<@everyone>", DAY_1, END, data, "https://docs.example.test/minutes")
        self.assertIn("&lt;@everyone&gt;", preview)
        self.assertIn("/minutes/portal/m-2026-11-26.docx", preview)

    def test_repeat_ingest_rebuilds_missing_word_without_a_new_model_call(self):
        path = self.import_day(DAY_1)
        path.unlink()
        with patch("meeting_agent.app.extract", side_effect=AssertionError("model called again")):
            self.assertEqual(self.import_day(DAY_1), path)
        self.assertTrue(path.is_file())
        self.assertEqual(len(meetings(self.db, "portal")), 1)

    def test_invalid_evidence_stops_before_saving(self):
        bad = fixture(DAY_1)
        bad["items"][0]["quote"] = "发送到另一个群聊"
        with self.assertRaisesRegex(ValueError, "引用校验失败"):
            self.import_day(DAY_1, lambda *_: bad)
        self.assertEqual(meetings(self.db, "portal"), [])
        self.assertFalse((self.root / "docs").exists())

    def test_overlap_preserves_one_copy_of_a_quote(self):
        lines = ["背景" * 590, "背景" * 590, "确定上线。" + "背景" * 590, "背景" * 590]
        item = fixture(DAY_1)["items"][0]
        item.update(paragraph=3, text="确定上线。", quote="确定上线。")
        seen = []

        def infer(_day, batch, _candidates):
            included = 3 in dict(batch)
            seen.append(included)
            return {"items": [item] if included else []}

        self.assertEqual(len(extract(lines, DAY_1, [], infer)), 1)
        self.assertEqual(seen, [True, True])

    def test_word_contains_native_chart_and_matching_bookmarks(self):
        path = self.import_day(DAY_1)
        with ZipFile(path) as package:
            self.assertIsNone(package.testzip())
            xml = etree.fromstring(package.read("word/document.xml"))
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            self.assertEqual(set(xml.xpath("//w:hyperlink/@w:anchor", namespaces=ns)),
                             set(xml.xpath("//w:bookmarkStart/@w:name", namespaces=ns)))
            chart = etree.fromstring(package.read("word/charts/chart1.xml"))
            counts = chart.xpath("//c:numCache/c:pt/c:v/text()", namespaces={"c": CHART_NS})
            self.assertEqual(sorted(map(int, counts)), [1, 1, 2])
            with ZipFile(BytesIO(package.read("word/embeddings/action-status.xlsx"))) as workbook:
                self.assertIn("xl/worksheets/sheet1.xml", workbook.namelist())

    def test_dashscope_uses_environment_key_and_default_model(self):
        reply = {"choices": [{"finish_reason": "stop", "message": {"content": '{"items":[]}'}}]}
        with (patch.dict("os.environ", {"DASHSCOPE_API_KEY": "test-key", "MEETING_MODEL": "old-local-id"}, clear=True),
              patch("meeting_agent.extract.urlopen", return_value=BytesIO(json.dumps(reply).encode())) as request):
            self.assertEqual(infer_dashscope(DAY_1, [(1, "周报")], []), {"items": []})
        sent = request.call_args.args[0]
        data = json.loads(sent.data)
        self.assertEqual(sent.full_url, "https://dashscope-intl.aliyuncs.com/compatible-mode/v1/chat/completions")
        self.assertEqual(sent.get_header("Authorization"), "Bearer test-key")
        self.assertEqual(data["model"], "qwen3.8-27b")
        self.assertFalse(data["enable_thinking"])
        self.assertEqual(data["response_format"], {"type": "json_object"})
        self.assertIn("JSON Schema", data["messages"][0]["content"])

    def test_dashscope_checks_key_and_destination_before_connecting(self):
        for env, message in (({}, "DASHSCOPE_API_KEY"), ({"DASHSCOPE_API_KEY": "test-key",
                              "DASHSCOPE_BASE_URL": "https://api.example.org/compatible-mode/v1"}, "DASHSCOPE_BASE_URL")):
            with (self.subTest(env=env), patch.dict("os.environ", env, clear=True),
                  patch("meeting_agent.extract.urlopen") as request):
                with self.assertRaisesRegex(ValueError, message):
                    infer_dashscope(DAY_1, [(1, "周报")], [])
                request.assert_not_called()

    def test_dashscope_accepts_workspace_endpoint_and_model_id(self):
        url = "https://ws-demo.cn-beijing.maas.aliyuncs.com/compatible-mode/v1"
        env = {"DASHSCOPE_API_KEY": "test-key", "DASHSCOPE_BASE_URL": url, "DASHSCOPE_MODEL": "private-model-id"}
        reply = {"choices": [{"finish_reason": "stop", "message": {"content": '{"items":[]}'}}]}
        with (patch.dict("os.environ", env, clear=True),
              patch("meeting_agent.extract.urlopen", return_value=BytesIO(json.dumps(reply).encode())) as request):
            infer_dashscope(DAY_1, [(1, "周报")], [])
        sent = request.call_args.args[0]
        self.assertEqual(sent.full_url, f"{url}/chat/completions")
        self.assertEqual(json.loads(sent.data)["model"], "private-model-id")

    def test_live_acceptance_requires_the_requested_model_in_both_responses(self):
        model = "qwen3.8-27b"
        decision = {"kind": "decision", "text": "向 10 家白名单客户开放新门户。", "paragraph": 1,
                    "quote": "11 月 26 日向 10 家白名单客户开放新门户。",
                    "owner": None, "due": None, "status": None}
        action = {"kind": "action", "text": "配置错误率告警", "paragraph": 2,
                  "quote": "陈遥：我负责配置错误率告警，截止日为 2026-11-25。目前正在配置。",
                  "owner": "陈遥", "due": "2026-11-25", "status": "doing"}
        update = {**action, "paragraph": 1, "quote": "A0001 已完成。", "owner": None, "due": None, "status": "done"}
        for reported in (model, "different-model"):
            with self.subTest(reported=reported):
                responses = []
                for n, items in enumerate(([decision, action], [update])):
                    payload = {"model": reported, "id": f"response-{n}",
                               "usage": {"prompt_tokens": 80, "completion_tokens": 60},
                               "choices": [{"finish_reason": "stop", "message": {"content": json.dumps({"items": items})}}]}
                    responses.append(BytesIO(json.dumps(payload).encode()))

                with (patch.dict("os.environ", {"DASHSCOPE_API_KEY": "test-key"}, clear=True),
                      patch("meeting_agent.extract.urlopen", side_effect=responses) as request):
                    proof, passed = verify(model, self.root / "checks")
                result = json.loads(proof.read_text(encoding="utf-8"))
                self.assertIs(passed, reported == model)
                self.assertIs(result["checks"]["model_and_response_ids"], reported == model)
                self.assertTrue(result["checks"]["same_action_completed"])
                requested = [json.loads(call.args[0].data)["model"] for call in request.call_args_list]
                self.assertEqual(requested, [model, model])

    def test_slack_prevents_duplicates_and_recovers_confirmed_failure(self):
        with patch("meeting_agent.app.post_slack", side_effect=SlackRejected("rejected")):
            with self.assertRaisesRegex(ValueError, "rejected"):
                publish(self.db, self.key, SLACK, "body")
        with patch("meeting_agent.app.post_slack") as send:
            publish(self.db, self.key, SLACK, "body")
            self.assertEqual(publish(self.db, self.key, SLACK, "body"), "这期报告已发送。")
            send.assert_called_once_with(SLACK, "body")

    def test_unknown_slack_delivery_waits_for_reconciliation(self):
        with patch("meeting_agent.app.post_slack", side_effect=ValueError("timeout")) as send:
            with self.assertRaisesRegex(ValueError, "timeout"):
                publish(self.db, self.key, SLACK, "body")
            with self.assertRaisesRegex(ValueError, "发送状态待确认"):
                publish(self.db, self.key, SLACK, "body")
            send.assert_called_once()
        settle(self.db, self.key, "sent")
        self.assertEqual(publish(self.db, self.key, SLACK, "body"), "这期报告已发送。")
