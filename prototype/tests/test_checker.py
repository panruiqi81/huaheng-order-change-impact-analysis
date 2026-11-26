"""单元测试：python -m unittest discover -s tests -v（在 prototype/ 目录下）"""

from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from eval.cases import BASE, CASES, TABLES, materialize  # noqa: E402
from eval.verify import check  # noqa: E402
from oc_checker.__main__ import run  # noqa: E402
from oc_checker.audit import bind_task, write_manifest  # noqa: E402
from oc_checker.models import load_snapshot  # noqa: E402
from oc_checker.render import to_markdown  # noqa: E402

NOW = datetime(2026, 11, 26, 4, 0, tzinfo=timezone.utc)


class TestCases(unittest.TestCase):
    def test_all_cases(self):
        with tempfile.TemporaryDirectory() as tmp:
            for case in CASES:
                with self.subTest(case=case.id):
                    result = run(materialize(case, Path(tmp) / case.id), case.order, now=NOW)
                    self.assertEqual(check(case, result), [])

    def test_expected_statuses_are_not_all_the_same(self):
        statuses = {c.expect for c in CASES}
        self.assertGreaterEqual(len(statuses), 4, "评估集应覆盖多种结论")


class TestReadOnly(unittest.TestCase):
    def test_snapshot_untouched_and_no_side_files(self):
        before = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in BASE.iterdir()}
        run(BASE, "O-202610-001", now=NOW)
        after = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in BASE.iterdir()}
        self.assertEqual(before, after)
        self.assertEqual(set(after), set(TABLES.values()), "不得在快照目录内产生任何文件")


class TestAudit(unittest.TestCase):
    def test_task_binds_rules_inputs_and_outputs(self):
        with tempfile.TemporaryDirectory() as tmp:
            snap_dir = materialize(CASES[0], Path(tmp) / "snap")
            out = Path(tmp) / "out"
            result = run(snap_dir, "O-202610-001", out, now=NOW)
            task = result["task"]
            self.assertTrue(task["task_id"].startswith("HH-20261126-"))
            self.assertEqual(task["rules_version"], "rules-huaheng-1.0.0")
            for name, digest in task["input_files"].items():
                self.assertEqual(digest, hashlib.sha256((snap_dir / name).read_bytes()).hexdigest())
            manifest = json.loads((out / "manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(manifest["task_id"], task["task_id"])
            self.assertEqual(manifest["outputs"]["report.json"],
                             hashlib.sha256((out / "report.json").read_bytes()).hexdigest())
            for name in ("report.md", "report.json"):
                self.assertIn(task["task_id"], (out / name).read_text(encoding="utf-8"), "输出须带任务编号")

    def test_input_change_changes_task_id(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = materialize(CASES[0], Path(tmp) / "a")
            b = materialize(CASES[7], Path(tmp) / "b")
            ta = bind_task(load_snapshot(a), "O-202610-001", NOW)["task_id"]
            tb = bind_task(load_snapshot(b), "O-202610-001", NOW)["task_id"]
            self.assertNotEqual(ta, tb)
            self.assertEqual(ta, bind_task(load_snapshot(a), "O-202610-001", NOW)["task_id"], "同一输入应得到同一任务编号")


class TestManualReview(unittest.TestCase):
    def test_failed_input_yields_no_partial_conclusion(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = run(materialize(CASES[13], Path(tmp) / "snap"), "O-202610-001", now=NOW)
        self.assertEqual(result["overall_status"], "MANUAL_REVIEW")
        self.assertIsNone(result["version_check"])
        self.assertIsNone(result["capacity_check"])
        self.assertEqual(result["confirmed"], [])
        self.assertTrue(any("缓存" in r for r in result["manual_review_reasons"]))

    def test_report_states_no_external_commitment(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = run(materialize(CASES[0], Path(tmp) / "snap"), "O-202610-001", now=NOW)
        md = to_markdown(result)
        self.assertIn("不构成对外交期承诺", md)
        owners = {p["owner"] for p in result["pending"]}
        self.assertIn("生产计划 + 销售（会签）", owners)
        self.assertIn("工厂经理 + 人事", owners)
        self.assertIn("销售（向客户答复后回填）", owners)


if __name__ == "__main__":
    unittest.main()
