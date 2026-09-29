from datetime import datetime
from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from taobao_uploader.browser import DuplicateVideoWarning, SubmissionUncertain
from taobao_uploader.publish_stats import PublicationStats
from taobao_uploader.runner import BatchRunner
from taobao_uploader.state import StateStore
from taobao_uploader.tasks import HEADERS, TZ, TaskValidationError, read_tasks


class FakeBrowser:
    def __init__(self, uncertain=False):
        self.prepared = []
        self.submitted = []
        self.uncertain = uncertain
        self.closed = False

    def start(self):
        pass

    def prepare(self, task):
        self.prepared.append(task.task_id)

    def submit_and_verify(self, task):
        self.submitted.append(task.task_id)
        if self.uncertain:
            raise SubmissionUncertain("结果未知")
        return "平台显示提交成功"

    def close(self):
        self.closed = True


class TaskFlowTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        (self.root / "a.mp4").write_bytes(b"video")
        self.xlsx = self.root / "tasks.xlsx"

    def workbook(self, rows):
        wb = Workbook()
        ws = wb.active
        ws.append(HEADERS)
        for row in rows:
            ws.append(row)
        wb.save(self.xlsx)

    def row(self, task_id="a", **changes):
        data = dict(zip(HEADERS, (task_id, "a.mp4", "测试视频", "内容", "", "", "", "", "", "", "内容无需标注")))
        data.update(changes)
        return [data[key] for key in HEADERS]

    def test_import_and_row_specific_validation(self):
        self.workbook([self.row(), self.row("b", **{"商品ID": "1;2;3;4;5;6;7"})])
        with self.assertRaises(TaskValidationError) as context:
            read_tasks(self.xlsx)
        self.assertIn("第 3 行", str(context.exception))
        self.assertIn("最多 6 个", str(context.exception))

    def test_successful_task_is_rechecked_after_restart(self):
        self.workbook([self.row()])
        tasks = read_tasks(self.xlsx)
        store = StateStore(self.root / "state.sqlite3")
        self.addCleanup(store.close)
        store.register(tasks)
        first = FakeBrowser()
        BatchRunner(tasks, store, lambda *_: None, lambda: first).run()
        self.assertEqual(store.get("a")[0], "成功")
        second = FakeBrowser()
        BatchRunner(tasks, store, lambda *_: None, lambda: second).run()
        self.assertEqual(second.submitted, ["a"])
        self.assertTrue(second.closed)

    def test_interrupted_submission_never_resubmits(self):
        self.workbook([self.row()])
        tasks = read_tasks(self.xlsx)
        store = StateStore(self.root / "state.sqlite3")
        self.addCleanup(store.close)
        store.register(tasks)
        store.set("a", "提交中")
        store.recover_interrupted(tasks)
        self.assertEqual(store.get("a")[0], "待核查")
        browser = FakeBrowser()
        BatchRunner(tasks, store, lambda *_: None, lambda: browser).run()
        self.assertEqual(browser.submitted, [])

    def test_dry_run_prepares_without_submitting(self):
        self.workbook([self.row()])
        tasks = read_tasks(self.xlsx)
        store = StateStore(self.root / "state.sqlite3")
        self.addCleanup(store.close)
        store.register(tasks)
        browser = FakeBrowser()
        BatchRunner(tasks, store, lambda *_: None, lambda: browser, dry_run=True).run()
        self.assertEqual(browser.prepared, ["a"])
        self.assertEqual(browser.submitted, [])
        self.assertEqual(store.get("a")[0], "验证通过")

    def test_duplicate_warning_skips_without_submission_and_stays_skipped(self):
        self.workbook([self.row()])
        tasks = read_tasks(self.xlsx)
        store = StateStore(self.root / "state.sqlite3")
        self.addCleanup(store.close)

        class DuplicateBrowser(FakeBrowser):
            def prepare(self, task):
                self.prepared.append(task.task_id)
                raise DuplicateVideoWarning("平台提示：视频可能重复发布")

        first = DuplicateBrowser()
        stats_dir = self.root / "stats"
        BatchRunner(tasks, store, lambda *_: None, lambda: first,
                    stats_dir=stats_dir).run()
        self.assertEqual(store.get("a")[:2], ("跳过", "平台提示：视频可能重复发布"))
        self.assertEqual(first.submitted, [])
        stats = PublicationStats(stats_dir)
        self.addCleanup(stats.close)
        self.assertEqual(len(stats.rows()), 1)
        self.assertEqual(stats.rows()[0][-1], "视频重复上传")
        second = DuplicateBrowser()
        BatchRunner(tasks, store, lambda *_: None, lambda: second,
                    stats_dir=stats_dir).run()
        self.assertEqual(second.prepared, ["a"])
        self.assertEqual(second.submitted, [])

    def test_uncertain_submission_is_not_retried(self):
        self.workbook([self.row()])
        tasks = read_tasks(self.xlsx)
        store = StateStore(self.root / "state.sqlite3")
        self.addCleanup(store.close)
        store.register(tasks)
        first = FakeBrowser(uncertain=True)
        BatchRunner(tasks, store, lambda *_: None, lambda: first).run()
        self.assertEqual(store.get("a")[0], "待核查")
        second = FakeBrowser()
        BatchRunner(tasks, store, lambda *_: None, lambda: second).run()
        self.assertEqual(second.submitted, [])

    def test_uncertain_submission_can_be_skipped_for_rest_of_batch(self):
        self.workbook([self.row("a"), self.row("b")])
        tasks = read_tasks(self.xlsx)
        store = StateStore(self.root / "state.sqlite3")
        self.addCleanup(store.close)
        store.register(tasks)

        class FirstUncertainBrowser(FakeBrowser):
            def submit_and_verify(self, task):
                self.submitted.append(task.task_id)
                if task.task_id == "a":
                    raise SubmissionUncertain("结果未知")
                return "平台显示提交成功"

        browser = FirstUncertainBrowser()
        BatchRunner(tasks, store, lambda *_: None, lambda: browser,
                    pause_on_uncertain=False).run()
        self.assertEqual(browser.submitted, ["a", "b"])
        self.assertEqual(store.get("a")[0], "待核查")
        self.assertEqual(store.get("b")[0], "成功")

    def test_changed_successful_task_reenters_upload_queue(self):
        self.workbook([self.row()])
        tasks = read_tasks(self.xlsx)
        store = StateStore(self.root / "state.sqlite3")
        self.addCleanup(store.close)
        store.register(tasks)
        store.set("a", "成功")
        self.workbook([self.row(**{"标题": "修改后标题"})])
        store.register(read_tasks(self.xlsx))
        self.assertEqual(store.get("a")[0], "待执行")

    def test_unsubmitted_task_can_be_corrected(self):
        self.workbook([self.row()])
        store = StateStore(self.root / "state.sqlite3")
        self.addCleanup(store.close)
        store.register(read_tasks(self.xlsx))
        store.set("a", "失败", "封面缺失")
        self.workbook([self.row(**{"标题": "改好后标题"})])
        store.register(read_tasks(self.xlsx))
        self.assertEqual(store.get("a")[0], "待执行")

    def test_date_must_be_future_beijing_time(self):
        self.workbook([self.row(**{"发布时间": "2026-09-23 10:00"})])
        now = datetime(2026, 9, 23, 11, 0, tzinfo=TZ)
        with self.assertRaisesRegex(TaskValidationError, "晚于当前北京时间"):
            read_tasks(self.xlsx, now=now)

    def test_scheduled_time_uses_platform_window(self):
        now = datetime(2026, 9, 23, 11, 0, tzinfo=TZ)
        self.workbook([self.row(**{"发布时间": "2026-09-23 12:59"})])
        with self.assertRaisesRegex(TaskValidationError, "提前 2 小时"):
            read_tasks(self.xlsx, now=now)
        self.workbook([self.row(**{"发布时间": "2026-10-24 11:01"})])
        with self.assertRaisesRegex(TaskValidationError, "未来 30 天"):
            read_tasks(self.xlsx, now=now)
        self.workbook([self.row(**{"发布时间": "2026-09-23 13:00"})])
        self.assertEqual(read_tasks(self.xlsx, now=now)[0].publish_at.hour, 13)


if __name__ == "__main__":
    unittest.main()
