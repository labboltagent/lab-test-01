from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def _col_to_letters(col: int) -> str:
    if col < 1:
        raise ValueError("Column index must be >= 1")
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def _cell_ref(row: int, col: int) -> str:
    return f"{_col_to_letters(col)}{row}"


def _letters_to_col(letters: str) -> int:
    col = 0
    for ch in letters:
        if not ("A" <= ch <= "Z"):
            break
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return col


def _parse_cell_ref(ref: str) -> tuple[int, int]:
    ref = ref.upper()
    letters = "".join(ch for ch in ref if "A" <= ch <= "Z")
    digits = "".join(ch for ch in ref if ch.isdigit())
    return int(digits) if digits else 0, _letters_to_col(letters) if letters else 0


@dataclass(frozen=True)
class SheetInfo:
    name: str
    path: str


class XlsxSheet:
    def __init__(self, info: SheetInfo, tree: ET.ElementTree, shared_strings: list[str]) -> None:
        self.info = info
        self.tree = tree
        self._shared_strings = shared_strings

    @property
    def name(self) -> str:
        return self.info.name

    def _sheetdata(self) -> ET.Element:
        root = self.tree.getroot()
        sheetdata = root.find(_qn(_NS_MAIN, "sheetData"))
        if sheetdata is None:
            sheetdata = ET.SubElement(root, _qn(_NS_MAIN, "sheetData"))
        return sheetdata

    def max_row_col(self) -> tuple[int, int]:
        max_row = 0
        max_col = 0
        for c in self.tree.getroot().iterfind(f".//{_qn(_NS_MAIN, 'c')}"):
            r = c.get("r")
            if not r:
                continue
            row, col = _parse_cell_ref(r)
            max_row = max(max_row, row)
            max_col = max(max_col, col)
        return max_row, max_col

    def get_row_values(self, row: int) -> list[Any]:
        max_row, max_col = self.max_row_col()
        if row > max_row:
            return []
        values: list[Any] = [None] * max_col
        sheetdata = self._sheetdata()
        row_el = sheetdata.find(f"{_qn(_NS_MAIN, 'row')}[@r='{row}']")
        if row_el is None:
            return values
        for c in row_el.findall(_qn(_NS_MAIN, "c")):
            ref = c.get("r")
            if not ref:
                continue
            r_row, r_col = _parse_cell_ref(ref)
            if r_row != row or r_col <= 0:
                continue
            values[r_col - 1] = self._read_cell_value(c)
        return values

    def _read_cell_value(self, c: ET.Element) -> Any:
        t = c.get("t")
        if t == "inlineStr":
            t_el = c.find(f"{_qn(_NS_MAIN, 'is')}/{_qn(_NS_MAIN, 't')}")
            return t_el.text if t_el is not None else ""
        v_el = c.find(_qn(_NS_MAIN, "v"))
        if v_el is None or v_el.text is None:
            return None
        if t == "s":
            try:
                return self._shared_strings[int(v_el.text)]
            except Exception:  # noqa: BLE001
                return v_el.text
        return v_el.text

    def ensure_header(self, header_name: str) -> int:
        header_row = self.get_row_values(1)
        for idx, val in enumerate(header_row, start=1):
            if str(val).strip() == header_name:
                return idx
        new_col = len(header_row) + 1 if header_row else 1
        self.set_cell(1, new_col, header_name)
        return new_col

    def set_cell(self, row: int, col: int, value: str) -> None:
        sheetdata = self._sheetdata()

        row_el = sheetdata.find(f"{_qn(_NS_MAIN, 'row')}[@r='{row}']")
        if row_el is None:
            row_el = ET.SubElement(sheetdata, _qn(_NS_MAIN, "row"), {"r": str(row)})

        ref = _cell_ref(row, col)
        cell_el = None
        for c in row_el.findall(_qn(_NS_MAIN, "c")):
            if c.get("r") == ref:
                cell_el = c
                break
        if cell_el is None:
            cell_el = ET.SubElement(row_el, _qn(_NS_MAIN, "c"), {"r": ref})

        cell_el.set("t", "inlineStr")
        for child in list(cell_el):
            cell_el.remove(child)
        is_el = ET.SubElement(cell_el, _qn(_NS_MAIN, "is"))
        t_el = ET.SubElement(is_el, _qn(_NS_MAIN, "t"))
        t_el.text = value


class XlsxFile:
    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self._parts: dict[str, bytes] = {}
        self._dirty: set[str] = set()
        self.shared_strings: list[str] = []
        self.sheets: list[SheetInfo] = []
        self._sheet_trees: dict[str, ET.ElementTree] = {}

    @classmethod
    def load(cls, path: str | Path) -> "XlsxFile":
        xlsx = cls(path)
        with zipfile.ZipFile(xlsx.path, "r") as zf:
            for name in zf.namelist():
                xlsx._parts[name] = zf.read(name)

        xlsx.shared_strings = xlsx._load_shared_strings()
        xlsx.sheets = xlsx._load_sheets()
        for info in xlsx.sheets:
            data = xlsx._parts[info.path]
            xlsx._sheet_trees[info.name] = ET.ElementTree(ET.fromstring(data))
        return xlsx

    def _load_shared_strings(self) -> list[str]:
        data = self._parts.get("xl/sharedStrings.xml")
        if not data:
            return []
        root = ET.fromstring(data)
        strings: list[str] = []
        for si in root.findall(_qn(_NS_MAIN, "si")):
            t_el = si.find(_qn(_NS_MAIN, "t"))
            if t_el is not None and t_el.text is not None:
                strings.append(t_el.text)
            else:
                strings.append("")
        return strings

    def _load_sheets(self) -> list[SheetInfo]:
        wb_root = ET.fromstring(self._parts["xl/workbook.xml"])
        rels_root = ET.fromstring(self._parts["xl/_rels/workbook.xml.rels"])
        rid_to_target: dict[str, str] = {}
        for rel in rels_root.findall(_qn(_NS_PKG_REL, "Relationship")):
            rid = rel.get("Id")
            target = rel.get("Target")
            if rid and target:
                if target.startswith("/"):
                    target_path = target.lstrip("/")
                else:
                    target_path = f"xl/{target}"
                rid_to_target[rid] = target_path

        sheets_el = wb_root.find(_qn(_NS_MAIN, "sheets"))
        if sheets_el is None:
            return []
        out: list[SheetInfo] = []
        for sh in sheets_el.findall(_qn(_NS_MAIN, "sheet")):
            name = sh.get("name")
            rid = sh.get(_qn(_NS_REL, "id"))
            if not name or not rid or rid not in rid_to_target:
                continue
            out.append(SheetInfo(name=name, path=rid_to_target[rid]))
        return out

    def sheet_names(self) -> list[str]:
        return [s.name for s in self.sheets]

    def get_sheet(self, name: str) -> XlsxSheet:
        if name not in self._sheet_trees:
            raise KeyError(f"Sheet not found: {name}")
        info = next(s for s in self.sheets if s.name == name)
        return XlsxSheet(info, self._sheet_trees[name], self.shared_strings)

    def mark_dirty(self, sheet_name: str) -> None:
        info = next(s for s in self.sheets if s.name == sheet_name)
        self._dirty.add(info.path)

    def save(self) -> None:
        for info in self.sheets:
            if info.path not in self._dirty:
                continue
            tree = self._sheet_trees[info.name]
            xml_bytes = ET.tostring(tree.getroot(), encoding="utf-8", xml_declaration=True)
            self._parts[info.path] = xml_bytes

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for name, data in self._parts.items():
                zf.writestr(name, data)
        self.path.write_bytes(buf.getvalue())


def create_simple_xlsx(path: str | Path, sheets: dict[str, list[list[Any]]]) -> None:
    """
    Create a minimal .xlsx using inline strings for all values.

    This is only intended for tests and simple workbooks.
    """

    path = Path(path)
    sheet_names = list(sheets.keys())

    def sheet_xml(rows: list[list[Any]]) -> bytes:
        ws = ET.Element(_qn(_NS_MAIN, "worksheet"))
        ws.set("xmlns", _NS_MAIN)
        sd = ET.SubElement(ws, _qn(_NS_MAIN, "sheetData"))
        for r_idx, row in enumerate(rows, start=1):
            row_el = ET.SubElement(sd, _qn(_NS_MAIN, "row"), {"r": str(r_idx)})
            for c_idx, value in enumerate(row, start=1):
                if value is None:
                    continue
                c = ET.SubElement(row_el, _qn(_NS_MAIN, "c"), {"r": _cell_ref(r_idx, c_idx), "t": "inlineStr"})
                is_el = ET.SubElement(c, _qn(_NS_MAIN, "is"))
                t_el = ET.SubElement(is_el, _qn(_NS_MAIN, "t"))
                t_el.text = str(value)
        return ET.tostring(ws, encoding="utf-8", xml_declaration=True)

    workbook = ET.Element(_qn(_NS_MAIN, "workbook"))
    workbook.set("xmlns", _NS_MAIN)
    workbook.set("xmlns:r", _NS_REL)
    sheets_el = ET.SubElement(workbook, _qn(_NS_MAIN, "sheets"))
    for idx, name in enumerate(sheet_names, start=1):
        sh = ET.SubElement(
            sheets_el,
            _qn(_NS_MAIN, "sheet"),
            {"name": name, "sheetId": str(idx), _qn(_NS_REL, "id"): f"rId{idx}"},
        )
        _ = sh

    workbook_xml = ET.tostring(workbook, encoding="utf-8", xml_declaration=True)

    rels = ET.Element(_qn(_NS_PKG_REL, "Relationships"))
    rels.set("xmlns", _NS_PKG_REL)
    for idx, _ in enumerate(sheet_names, start=1):
        ET.SubElement(
            rels,
            _qn(_NS_PKG_REL, "Relationship"),
            {
                "Id": f"rId{idx}",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                "Target": f"worksheets/sheet{idx}.xml",
            },
        )
    workbook_rels_xml = ET.tostring(rels, encoding="utf-8", xml_declaration=True)

    root_rels = ET.Element(_qn(_NS_PKG_REL, "Relationships"))
    root_rels.set("xmlns", _NS_PKG_REL)
    ET.SubElement(
        root_rels,
        _qn(_NS_PKG_REL, "Relationship"),
        {
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "Target": "xl/workbook.xml",
        },
    )
    root_rels_xml = ET.tostring(root_rels, encoding="utf-8", xml_declaration=True)

    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>
""".encode("utf-8")

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        for idx, name in enumerate(sheet_names, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", sheet_xml(sheets[name]))

