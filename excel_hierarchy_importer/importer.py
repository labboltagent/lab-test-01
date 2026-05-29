from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .xlsx import XlsxFile


STATUS_HEADER = "__status__"


@dataclass(frozen=True)
class HierarchySheets:
    parent: str
    child: str
    grandchild: str


def _parse_id_list(value: Any) -> list[str]:
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    return [part.strip() for part in text.split(",") if part.strip()]


def import_hierarchy_excel(
    excel_path: str | Path,
    conn: sqlite3.Connection,
    *,
    sheets: HierarchySheets | None = None,
    parent_child_column: str = "child_ids",
    child_grandchild_column: str = "grandchild_ids",
) -> None:
    """
    Minimal importer for a 3-sheet Excel workbook representing a simple hierarchy.

    Expected workbook conventions:
      - Exactly 3 sheets, or provide explicit `sheets=HierarchySheets(...)`.
      - Each sheet has an `id` column.
      - Parent sheet has a `child_ids` column (comma-separated child ids).
      - Child sheet has a `grandchild_ids` column (comma-separated grandchild ids).
      - A `__status__` column is added/updated on each sheet per row.

    Database conventions (tables must already exist):
      - Tables named exactly like the sheet names.
      - Join tables: `{parent}_{child}` and `{child}_{grandchild}` with columns
        `{parent}_id`, `{child}_id`, `{grandchild}_id` (using the sheet names).
    """

    excel_path = Path(excel_path)
    xlsx = XlsxFile.load(excel_path)
    sheet_names = xlsx.sheet_names()

    if sheets is None:
        if len(sheet_names) != 3:
            raise ValueError(
                f"Expected exactly 3 sheets (parent/child/grandchild). Found {len(sheet_names)}: {sheet_names}"
            )
        sheets = HierarchySheets(parent=sheet_names[0], child=sheet_names[1], grandchild=sheet_names[2])

    ws_parent = xlsx.get_sheet(sheets.parent)
    ws_child = xlsx.get_sheet(sheets.child)
    ws_grandchild = xlsx.get_sheet(sheets.grandchild)

    parent_status_col = ws_parent.ensure_header(STATUS_HEADER)
    child_status_col = ws_child.ensure_header(STATUS_HEADER)
    grandchild_status_col = ws_grandchild.ensure_header(STATUS_HEADER)
    xlsx.mark_dirty(ws_parent.name)
    xlsx.mark_dirty(ws_child.name)
    xlsx.mark_dirty(ws_grandchild.name)

    def header_map(ws) -> dict[str, int]:
        headers = ws.get_row_values(1)
        return {str(h).strip(): i + 1 for i, h in enumerate(headers) if h is not None and str(h).strip()}

    parent_headers = header_map(ws_parent)
    child_headers = header_map(ws_child)
    grandchild_headers = header_map(ws_grandchild)

    def existing_status(ws, status_col: int, row: int) -> str | None:
        values = ws.get_row_values(row)
        if status_col <= 0 or status_col > len(values):
            return None
        val = values[status_col - 1]
        return str(val).strip() if val is not None else None

    def row_dict(headers: dict[str, int], ws, row: int) -> dict[str, Any]:
        values = ws.get_row_values(row)
        out: dict[str, Any] = {}
        for name, col in headers.items():
            out[name] = values[col - 1] if col - 1 < len(values) else None
        return out

    def insert_base_row(table: str, row: dict[str, Any], skip_columns: set[str]) -> None:
        columns = [c for c in row.keys() if c not in skip_columns and c != STATUS_HEADER]
        values = [row[c] for c in columns]
        placeholders = ", ".join(["?"] * len(columns))
        col_sql = ", ".join([f'"{c}"' for c in columns])
        conn.execute(f'INSERT INTO "{table}" ({col_sql}) VALUES ({placeholders})', values)

    def insert_join_row(join_table: str, columns: list[str], values: list[Any]) -> None:
        placeholders = ", ".join(["?"] * len(columns))
        col_sql = ", ".join([f'"{c}"' for c in columns])
        conn.execute(f'INSERT OR IGNORE INTO "{join_table}" ({col_sql}) VALUES ({placeholders})', values)

    parent_child_join = f"{sheets.parent}_{sheets.child}"
    child_grandchild_join = f"{sheets.child}_{sheets.grandchild}"
    parent_id_col = f"{sheets.parent}_id"
    child_id_col = f"{sheets.child}_id"
    grandchild_id_col = f"{sheets.grandchild}_id"

    savepoint_counter = 0

    conn.execute("BEGIN")
    try:
        max_grandchild_row, _ = ws_grandchild.max_row_col()
        for i in range(2, max_grandchild_row + 1):
            if existing_status(ws_grandchild, grandchild_status_col, i) == "inserted":
                continue
            row = row_dict(grandchild_headers, ws_grandchild, i)
            if all(v is None for v in row.values()):
                continue
            savepoint_counter += 1
            sp = f"sp_{savepoint_counter}"
            conn.execute(f"SAVEPOINT {sp}")
            try:
                insert_base_row(sheets.grandchild, row, skip_columns=set())
                conn.execute(f"RELEASE {sp}")
                ws_grandchild.set_cell(i, grandchild_status_col, "inserted")
                xlsx.mark_dirty(ws_grandchild.name)
            except Exception as e:  # noqa: BLE001
                conn.execute(f"ROLLBACK TO {sp}")
                conn.execute(f"RELEASE {sp}")
                ws_grandchild.set_cell(i, grandchild_status_col, f"error: {e}")
                xlsx.mark_dirty(ws_grandchild.name)

        max_child_row, _ = ws_child.max_row_col()
        for i in range(2, max_child_row + 1):
            if existing_status(ws_child, child_status_col, i) == "inserted":
                continue
            row = row_dict(child_headers, ws_child, i)
            if all(v is None for v in row.values()):
                continue
            savepoint_counter += 1
            sp = f"sp_{savepoint_counter}"
            conn.execute(f"SAVEPOINT {sp}")
            try:
                insert_base_row(sheets.child, row, skip_columns={child_grandchild_column})
                child_id = str(row.get("id")).strip()
                for gc_id in _parse_id_list(row.get(child_grandchild_column)):
                    insert_join_row(child_grandchild_join, [child_id_col, grandchild_id_col], [child_id, gc_id])
                conn.execute(f"RELEASE {sp}")
                ws_child.set_cell(i, child_status_col, "inserted")
                xlsx.mark_dirty(ws_child.name)
            except Exception as e:  # noqa: BLE001
                conn.execute(f"ROLLBACK TO {sp}")
                conn.execute(f"RELEASE {sp}")
                ws_child.set_cell(i, child_status_col, f"error: {e}")
                xlsx.mark_dirty(ws_child.name)

        max_parent_row, _ = ws_parent.max_row_col()
        for i in range(2, max_parent_row + 1):
            if existing_status(ws_parent, parent_status_col, i) == "inserted":
                continue
            row = row_dict(parent_headers, ws_parent, i)
            if all(v is None for v in row.values()):
                continue
            savepoint_counter += 1
            sp = f"sp_{savepoint_counter}"
            conn.execute(f"SAVEPOINT {sp}")
            try:
                insert_base_row(sheets.parent, row, skip_columns={parent_child_column})
                parent_id = str(row.get("id")).strip()
                for child_id in _parse_id_list(row.get(parent_child_column)):
                    insert_join_row(parent_child_join, [parent_id_col, child_id_col], [parent_id, child_id])
                conn.execute(f"RELEASE {sp}")
                ws_parent.set_cell(i, parent_status_col, "inserted")
                xlsx.mark_dirty(ws_parent.name)
            except Exception as e:  # noqa: BLE001
                conn.execute(f"ROLLBACK TO {sp}")
                conn.execute(f"RELEASE {sp}")
                ws_parent.set_cell(i, parent_status_col, f"error: {e}")
                xlsx.mark_dirty(ws_parent.name)

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        xlsx.save()


def main(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Import a 3-sheet hierarchy Excel workbook into a SQLite DB.")
    parser.add_argument("excel_path", help="Path to .xlsx file")
    parser.add_argument("--sqlite", dest="sqlite_path", required=True, help="SQLite DB path (use ':memory:' for in-mem)")
    args = parser.parse_args(argv)

    conn = sqlite3.connect(args.sqlite_path)
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        import_hierarchy_excel(args.excel_path, conn)
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
