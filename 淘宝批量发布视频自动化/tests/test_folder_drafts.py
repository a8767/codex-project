from pathlib import Path
from contextlib import closing
import sqlite3
import tempfile
import unittest

from openpyxl import load_workbook

from taobao_uploader.copy_library import seed_copy_templates
from taobao_uploader.folder_drafts import generate_folder_draft
from taobao_uploader.reference_match import _category, recommend
from taobao_uploader.tasks import HEADERS, TaskValidationError, read_tasks


class FolderDraftTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.folder = self.root / "淘宝视频批量上传"
        (self.folder / "GT-05").mkdir(parents=True)
        (self.folder / "氛围感视频").mkdir(parents=True)
        for relative in ("GT-05/片段1.mp4", "GT-05/片段2.MP4", "氛围感视频/片段.mp4"):
            (self.folder / relative).write_bytes(b"video")
        self.db_path = self.root / "references.sqlite3"
        db = sqlite3.connect(self.db_path)
        db.executescript("""
            CREATE TABLE videos(video_key TEXT PRIMARY KEY,title TEXT NOT NULL);
            CREATE TABLE video_topics(video_key TEXT,topic TEXT);
            INSERT INTO videos VALUES('v1','GT05山地车骑行 #山地车');
            INSERT INTO videos VALUES('v2','GT-05山地自行车日常 #山地车');
            INSERT INTO videos VALUES('v3','GT-05A公路车 #公路车');
            INSERT INTO video_topics VALUES('v1','山地车');
            INSERT INTO video_topics VALUES('v2','山地车');
            INSERT INTO video_topics VALUES('v3','公路车');
        """)
        db.close()

    def test_model_folders_get_matching_topics_and_generic_folder_stays_generic(self):
        output = self.root / "草稿.xlsx"
        result = generate_folder_draft(self.folder, self.db_path, output)
        self.assertEqual((result.total, result.model_matched, result.generic), (3, 2, 1))
        wb = load_workbook(output)
        self.assertEqual(tuple(c.value for c in wb.active[1][:len(HEADERS)]), HEADERS)
        rows = list(wb.active.iter_rows(min_row=2, values_only=True))
        model_rows = [row for row in rows if "GT-05" in row[1]]
        generic = next(row for row in rows if "氛围感视频" in row[1])
        self.assertEqual(len(model_rows), 2)
        self.assertTrue(all("GT-05" not in row[2] and "GT-05" not in row[3]
                            and row[8] == "山地骑行的快乐"
                            for row in model_rows))
        self.assertTrue(all(row[7] == "山地自行车;骑行;户外骑行;#GT-05" for row in model_rows))
        self.assertNotEqual(model_rows[0][2], model_rows[1][2])
        self.assertTrue(all(row[3] and "GT-05" in row[3] for row in model_rows))
        self.assertTrue(all("参考视频=2" in row[12] for row in model_rows))
        self.assertNotIn("GT-05", generic[2])
        self.assertEqual(generic[8], "进来一起享受骑行")
        self.assertEqual(generic[7], "骑行;自行车")
        self.assertEqual(generic[10], None)  # Declaration is a required human review step.
        self.assertEqual(len({row[0] for row in rows}), 3)
        wb.close()

    def test_thirty_model_videos_get_distinct_titles_and_body_copy(self):
        for number in range(3, 31):
            (self.folder / "GT-05" / f"片段{number}.mp4").write_bytes(b"video")
        db = sqlite3.connect(self.db_path)
        try:
            before = db.execute("SELECT count(*) FROM videos").fetchone()[0]
            count = seed_copy_templates(db)
            self.assertEqual(seed_copy_templates(db), count)
            self.assertEqual(db.execute("SELECT count(*) FROM videos").fetchone()[0], before)
            self.assertGreaterEqual(count, 80)
        finally:
            db.close()

        output = self.root / "三十条草稿.xlsx"
        generate_folder_draft(self.folder, self.db_path, output)
        workbook = load_workbook(output, read_only=True)
        try:
            rows = [row for row in workbook.active.iter_rows(min_row=2, values_only=True)
                    if "GT-05" in row[1]]
            self.assertEqual(len(rows), 30)
            self.assertEqual(len({row[2] for row in rows}), 30)
            self.assertEqual(len({row[3] for row in rows}), 30)
            self.assertTrue(all(len(row[2]) <= 30 and len(row[3]) <= 1000
                                and row[10] is None for row in rows))
        finally:
            workbook.close()

    def test_generated_draft_requires_declaration_before_import(self):
        output = self.root / "草稿.xlsx"
        generate_folder_draft(self.folder, self.db_path, output)
        with self.assertRaisesRegex(TaskValidationError, "创作者声明"):
            read_tasks(output)
        wb = load_workbook(output)
        for row in range(2, wb.active.max_row + 1):
            wb.active.cell(row, 11, "内容无需标注")
        wb.save(output)
        tasks = read_tasks(output)
        self.assertEqual(len(tasks), 3)
        self.assertEqual([task.topic for task in tasks].count("山地骑行的快乐"), 2)

    def test_unknown_model_is_marked_as_unverified(self):
        folder = self.folder / "ABC-999"
        folder.mkdir()
        (folder / "other.mp4").write_bytes(b"video")
        output = self.root / "草稿.xlsx"
        result = generate_folder_draft(self.folder, self.db_path, output)
        self.assertEqual(result.no_reference, 1)
        wb = load_workbook(output, read_only=True)
        unknown = next(row for row in wb.active.iter_rows(min_row=2, values_only=True)
                       if "ABC-999" in row[1])
        self.assertNotIn("ABC-999", unknown[2])
        self.assertNotIn("ABC-999", unknown[3])
        self.assertIn("#ABC-999", unknown[7])
        self.assertIn("参考视频=0", unknown[12])
        self.assertEqual(unknown[7], "骑行;自行车")
        wb.close()

    def test_ranked_references_choose_relevant_theme_and_confirmed_topic(self):
        with closing(sqlite3.connect(self.db_path)) as db:
            db.executescript("""
                ALTER TABLE videos ADD COLUMN platform TEXT DEFAULT '抖音';
                ALTER TABLE videos ADD COLUMN body TEXT DEFAULT '';
                ALTER TABLE video_topics ADD COLUMN extraction TEXT DEFAULT '标题#话题';
                CREATE TABLE guanghe_hot_occurrences(
                    video_key TEXT,time_range TEXT,chart TEXT,rank INTEGER,captured_at TEXT);
                INSERT INTO videos(video_key,title,platform,body) VALUES
                  ('accessory','山地车骑行头盔','淘宝光合','骑行必备装备'),
                  ('mountain','山地车周末骑行记录','淘宝光合','户外骑行，分享周末的想法'),
                  ('general','骑行日常片段','淘宝光合','记录骑行日常');
                INSERT INTO guanghe_hot_occurrences VALUES
                  ('accessory','最近7天','热度榜',1,'2026-09-28'),
                  ('mountain','最近7天','热度榜',2,'2026-09-28'),
                  ('general','最近7天','热度榜',3,'2026-09-28');
                INSERT INTO video_topics(video_key,topic,extraction) VALUES
                  ('accessory','骑行必备装备','平台话题'),
                  ('mountain','山地骑行的快乐','平台话题'),
                  ('general','进来一起享受骑行','平台话题');
            """)
        output = self.root / "智能匹配.xlsx"
        generate_folder_draft(self.folder, self.db_path, output)
        wb = load_workbook(output, read_only=True)
        try:
            rows = list(wb.active.iter_rows(min_row=2, values_only=True))
        finally:
            wb.close()
        model_row = next(row for row in rows if "GT-05" in row[1])
        generic_row = next(row for row in rows if "氛围感视频" in row[1])
        self.assertEqual(model_row[8], "山地骑行的快乐")
        self.assertEqual(model_row[13], "山地车周末骑行记录")
        self.assertIn("光合热门榜第2名", model_row[15])
        self.assertIn("榜单同类参考", model_row[12])
        self.assertNotIn("户外骑行，分享周末的想法", model_row[3])
        self.assertEqual(generic_row[8], "进来一起享受骑行")
        self.assertEqual(generic_row[13], "骑行日常片段")
        self.assertNotIn("山地车", generic_row[2])

    def test_accessory_titles_are_not_whole_bicycle_references(self):
        self.assertEqual(_category("自行车存放神器！放车巨稳", "山地车通用"), "配件或非自行车")
        self.assertEqual(_category("自行车骑行神秘小黑盒", "骑行安全锁"), "配件或非自行车")
        self.assertEqual(_category("骑行好物分享", "骑行眼镜防风镜"), "配件或非自行车")

    def test_quality_video_supplies_reference_when_hot_chart_has_no_match(self):
        with closing(sqlite3.connect(self.db_path)) as db:
            db.executescript("""
                ALTER TABLE videos ADD COLUMN platform TEXT DEFAULT '抖音';
                ALTER TABLE videos ADD COLUMN body TEXT DEFAULT '';
                CREATE TABLE video_sources(video_key TEXT,source_type TEXT);
                INSERT INTO videos(video_key,title,platform,body) VALUES
                  ('quality','GT-05山地车日常骑行','淘宝光合','记录这辆山地车的日常');
                INSERT INTO video_sources VALUES('quality','淘宝优质视频');
            """)
            advice = recommend(db, "GT-05", "山地车")
        self.assertEqual(advice.title, "GT-05山地车日常骑行")
        self.assertEqual(advice.source, "淘宝优质视频")
        self.assertEqual(advice.topic, "山地骑行的快乐")

    def test_generic_folders_do_not_repeat_copy(self):
        another = self.folder / "城市骑行视频"
        another.mkdir()
        (another / "另一个片段.mp4").write_bytes(b"video")
        output = self.root / "不同通用文件夹.xlsx"
        generate_folder_draft(self.folder, self.db_path, output)
        wb = load_workbook(output, read_only=True)
        try:
            rows = list(wb.active.iter_rows(min_row=2, values_only=True))
            self.assertEqual(len({row[2] for row in rows}), len(rows))
            self.assertEqual(len({row[3] for row in rows}), len(rows))
        finally:
            wb.close()


if __name__ == "__main__":
    unittest.main()
