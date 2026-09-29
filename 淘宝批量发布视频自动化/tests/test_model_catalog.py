from pathlib import Path
import sqlite3
import tempfile
import unittest

from openpyxl import Workbook

from taobao_uploader.model_catalog import find_model, import_model_workbooks


class ModelCatalogTests(unittest.TestCase):
    def test_import_is_idempotent_and_model_lookup_is_exact(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            database = root / "reference.sqlite3"
            sqlite3.connect(database).close()
            source = root / "永久车型编码20260924.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(("中文编码", "英文编码", "车型"))
            sheet.append(("CN01-20 20寸 童车", "CN01-20-20C", "CN01-20"))
            sheet.append(("CN01-201 20寸 童车", "CN01-201-20C", "CN01-201"))
            workbook.save(source)
            workbook.close()

            self.assertEqual(import_model_workbooks(database, [source]), {"永久": 2})
            self.assertEqual(import_model_workbooks(database, [source]), {"永久": 2})
            matches = find_model(database, "永久", "CN 01-20")
            self.assertEqual(matches, [("CN01-20", "CN01-20 20寸 童车", "CN01-20-20C")])
            connection = sqlite3.connect(database)
            try:
                self.assertEqual(connection.execute("SELECT count(*) FROM model_catalog").fetchone()[0], 2)
                self.assertEqual(connection.execute(
                    "SELECT source_row FROM model_catalog WHERE model_name='CN01-20'").fetchone()[0], 2)
            finally:
                connection.close()


if __name__ == "__main__":
    unittest.main()
