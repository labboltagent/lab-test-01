import sqlite3
import tempfile
import unittest
from pathlib import Path

from excel_hierarchy_importer import import_hierarchy_excel
from excel_hierarchy_importer.xlsx import XlsxFile, create_simple_xlsx


class TestHierarchyImporter(unittest.TestCase):
    def setUp(self) -> None:
        self.conn = sqlite3.connect(":memory:")
        self.conn.execute("PRAGMA foreign_keys = ON")

        self.parent = "parent"
        self.child = "child"
        self.grandchild = "grandchild"

        self.conn.executescript(
            f"""
            CREATE TABLE "{self.parent}" (
              id TEXT PRIMARY KEY,
              name TEXT
            );
            CREATE TABLE "{self.child}" (
              id TEXT PRIMARY KEY,
              name TEXT
            );
            CREATE TABLE "{self.grandchild}" (
              id TEXT PRIMARY KEY,
              name TEXT
            );

            CREATE TABLE "{self.parent}_{self.child}" (
              "{self.parent}_id" TEXT NOT NULL,
              "{self.child}_id" TEXT NOT NULL,
              PRIMARY KEY ("{self.parent}_id", "{self.child}_id"),
              FOREIGN KEY ("{self.parent}_id") REFERENCES "{self.parent}"(id),
              FOREIGN KEY ("{self.child}_id") REFERENCES "{self.child}"(id)
            );

            CREATE TABLE "{self.child}_{self.grandchild}" (
              "{self.child}_id" TEXT NOT NULL,
              "{self.grandchild}_id" TEXT NOT NULL,
              PRIMARY KEY ("{self.child}_id", "{self.grandchild}_id"),
              FOREIGN KEY ("{self.child}_id") REFERENCES "{self.child}"(id),
              FOREIGN KEY ("{self.grandchild}_id") REFERENCES "{self.grandchild}"(id)
            );
            """
        )

    def tearDown(self) -> None:
        self.conn.close()

    def _make_workbook(self, path: Path) -> None:
        create_simple_xlsx(
            path,
            sheets={
                self.parent: [
                    ["id", "name", "child_ids"],
                    ["p1", "Parent 1", "c1, c2"],
                    ["p2", "Parent 2", "c2"],
                ],
                self.child: [
                    ["id", "name", "grandchild_ids"],
                    ["c1", "Child 1", "g1"],
                    ["c2", "Child 2", "g1, g2"],
                ],
                self.grandchild: [
                    ["id", "name"],
                    ["g1", "Grand 1"],
                    ["g2", "Grand 2"],
                ],
            },
        )

    def test_import_and_status_updates(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            excel_path = Path(tmpdir) / "hierarchy.xlsx"
            self._make_workbook(excel_path)

            import_hierarchy_excel(excel_path, self.conn)

            xlsx = XlsxFile.load(excel_path)
            for sheet_name in (self.parent, self.child, self.grandchild):
                ws = xlsx.get_sheet(sheet_name)
                headers = ws.get_row_values(1)
                self.assertIn("__status__", headers)
                status_col = [str(h).strip() for h in headers].index("__status__") + 1
                max_row, _ = ws.max_row_col()
                for r in range(2, max_row + 1):
                    self.assertEqual(ws.get_row_values(r)[status_col - 1], "inserted")

            parent_count = self.conn.execute(f'SELECT COUNT(*) FROM "{self.parent}"').fetchone()[0]
            child_count = self.conn.execute(f'SELECT COUNT(*) FROM "{self.child}"').fetchone()[0]
            grandchild_count = self.conn.execute(f'SELECT COUNT(*) FROM "{self.grandchild}"').fetchone()[0]
            parent_child_count = self.conn.execute(
                f'SELECT COUNT(*) FROM "{self.parent}_{self.child}"'
            ).fetchone()[0]
            child_grandchild_count = self.conn.execute(
                f'SELECT COUNT(*) FROM "{self.child}_{self.grandchild}"'
            ).fetchone()[0]

            self.assertEqual(parent_count, 2)
            self.assertEqual(child_count, 2)
            self.assertEqual(grandchild_count, 2)
            self.assertEqual(parent_child_count, 3)
            self.assertEqual(child_grandchild_count, 3)
