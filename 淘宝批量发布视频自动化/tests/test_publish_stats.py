from datetime import datetime, timedelta
from dataclasses import replace
from pathlib import Path
import tempfile
import unittest

from openpyxl import load_workbook

from taobao_uploader.browser import VerifiedPublication
from taobao_uploader.publish_stats import HEADERS, PublicationStats
from taobao_uploader.runner import BatchRunner
from taobao_uploader.state import StateStore
from taobao_uploader.tasks import TZ, Task


class ConfirmedBrowser:
    def __init__(self):
        self.submissions = 0

    def start(self):
        pass

    def prepare(self, task):
        pass

    def submit_and_verify(self, task):
        self.submissions += 1
        return VerifiedPublication("1234567890123456", "后台已找到作品 ID 1234567890123456",
                                   datetime(2026, 9, 27, 10, 5, tzinfo=TZ))

    def close(self):
        pass


class PublicationStatsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        video = self.root / "GT-05测试.mp4"
        video.write_bytes(b"video")
        self.task = Task(2, "task-1", video, "GT-05骑行", "山路测试", (), None,
                         (), (), "山地骑行的快乐", None, "内容无需标注")
        self.source = self.root / "tasks.xlsx"

    def test_six_columns_and_idempotent_export(self):
        stats = PublicationStats(self.root / "stats")
        self.addCleanup(stats.close)
        submitted = datetime(2026, 9, 27, 10, 0, tzinfo=TZ)
        verified = submitted + timedelta(minutes=1)
        for _ in range(2):
            stats.record(self.task, "1234567890123456", source_excel=self.source,
                         submitted_at=submitted, verified_at=verified)
            path = stats.export()
        wb = load_workbook(path, read_only=True, data_only=True)
        self.addCleanup(wb.close)
        rows = list(wb.active.values)
        self.assertEqual(rows[0], HEADERS)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1][:4], ("GT-05测试.mp4", "1234567890123456",
                                        "GT-05骑行\n山路测试", "山地骑行的快乐"))
        self.assertEqual(rows[1][4], datetime(2026, 9, 27, 10, 0))
        self.assertEqual(rows[0], HEADERS)

    def test_repeat_publication_of_same_task_is_recorded_as_an_attempt(self):
        stats = PublicationStats(self.root / "stats")
        self.addCleanup(stats.close)
        submitted = datetime(2026, 9, 27, 10, 0, tzinfo=TZ)
        verified = submitted + timedelta(minutes=1)
        stats.record(self.task, "1234567890123456", source_excel=self.source,
                     submitted_at=submitted, verified_at=verified)
        stats.record(self.task, "9999999999999999", source_excel=self.source,
                     submitted_at=submitted + timedelta(hours=1),
                     verified_at=verified + timedelta(hours=1))
        self.assertEqual(len(stats.rows()), 2)
        task_ids = [row[0] for row in stats.db.execute(
            "SELECT task_id FROM publications ORDER BY submitted_at"
        )]
        self.assertEqual(task_ids, ["task-1", "task-1#attempt-2"])

    def test_scheduled_time_and_conflicting_work_id(self):
        task = Task(2, "task-1", self.task.video, self.task.title, self.task.body,
                    (), None, (), (), self.task.topic,
                    datetime(2026, 9, 28, 10, 0, tzinfo=TZ), self.task.declaration)
        stats = PublicationStats(self.root / "stats")
        self.addCleanup(stats.close)
        stats.record(task, "1234567890123456", source_excel=self.source,
                     submitted_at=datetime(2026, 9, 27, 10, tzinfo=TZ),
                     verified_at=datetime(2026, 9, 27, 11, tzinfo=TZ))
        with self.assertRaisesRegex(ValueError, "另一作品 ID"):
            stats.record(task, "9999999999999999", source_excel=self.source,
                         submitted_at=None, verified_at=datetime.now(TZ))
        wb = load_workbook(stats.export(), read_only=True, data_only=True)
        self.addCleanup(wb.close)
        self.assertEqual(wb.active["E2"].value, datetime(2026, 9, 28, 10, 0))

    def test_runner_records_once_and_backfill_skips_existing(self):
        store = StateStore(self.root / "state.sqlite3")
        self.addCleanup(store.close)
        store.register([self.task])
        browser = ConfirmedBrowser()
        stats_dir = self.root / "stats"
        runner = BatchRunner([self.task], store, lambda *_: None, lambda: browser,
                             source_workbook=self.source, stats_dir=stats_dir)
        runner.run()
        runner.run()
        self.assertEqual(browser.submissions, 2)
        self.assertEqual(store.get(self.task.task_id)[0], "成功")
        stats = PublicationStats(stats_dir)
        self.addCleanup(stats.close)
        self.assertEqual(len(stats.rows()), 1)
        self.assertEqual(stats.backfill_successes([self.task], self.root / "state.sqlite3",
                                                  self.source), 0)

    def test_confirmed_content_tags_appear_in_copy_without_extra_column(self):
        tagged = replace(self.task, content_tags=("儿童自行车", "骑行"))
        stats = PublicationStats(self.root / "stats")
        self.addCleanup(stats.close)
        stats.record(tagged, "1234567890123456", source_excel=self.source,
                     submitted_at=datetime(2026, 9, 27, 10, tzinfo=TZ),
                     verified_at=datetime(2026, 9, 27, 11, tzinfo=TZ))
        wb = load_workbook(stats.export(), read_only=True, data_only=True)
        self.addCleanup(wb.close)
        self.assertEqual(wb.active.max_column, 5)
        self.assertEqual(wb.active["C2"].value,
                         "GT-05骑行\n山路测试\n#儿童自行车 #骑行")

    def test_backfill_only_confirmed_success(self):
        store = StateStore(self.root / "state.sqlite3")
        self.addCleanup(store.close)
        store.register([self.task])
        store.set(self.task.task_id, "成功", "后台已确认作品 ID 1234567890123456")
        stats = PublicationStats(self.root / "stats")
        self.addCleanup(stats.close)
        self.assertEqual(stats.backfill_successes([self.task], self.root / "state.sqlite3",
                                                  self.source), 1)
        self.assertEqual(stats.backfill_successes([self.task], self.root / "state.sqlite3",
                                                  self.source), 0)
        self.assertEqual(len(stats.rows()), 1)

    def test_duplicate_skip_is_exported_with_reason_and_no_fake_id_or_publish_time(self):
        stats = PublicationStats(self.root / "stats")
        self.addCleanup(stats.close)
        when = datetime(2026, 9, 28, 9, 0, tzinfo=TZ)
        for _ in range(2):
            stats.record_skipped(self.task, source_excel=self.source, skipped_at=when)
        wb = load_workbook(stats.export(), read_only=True, data_only=True)
        self.addCleanup(wb.close)
        rows = list(wb.active.values)
        self.assertEqual(rows[0], HEADERS)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1][0], self.task.video.name)
        self.assertEqual(rows[1][1], "")
        self.assertIsNone(rows[1][4])
        self.assertEqual(rows[1][5], "视频重复上传")


if __name__ == "__main__":
    unittest.main()
