from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def text_of(node: ET.Element | None) -> str:
    if node is None:
        return ""
    return "".join(part.text or "" for part in node.iter(f"{{{MAIN}}}t"))


def inspect_xlsx(path: Path) -> dict:
    with zipfile.ZipFile(path) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared_strings = [text_of(item) for item in root.findall(f"{{{MAIN}}}si")]

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in relationships.findall(f"{{{PKG_REL}}}Relationship")
        }

        result = {"path": str(path), "sheets": []}
        sheets_node = workbook.find(f"{{{MAIN}}}sheets")
        for sheet in list(sheets_node) if sheets_node is not None else []:
            name = sheet.attrib["name"]
            # Most workbooks use targets relative to ``xl/`` (for example,
            # ``worksheets/sheet1.xml``).  The KiotViet template produced by
            # this Excel version already prefixes its targets with ``xl/``.
            # Accept both valid forms instead of producing ``xl/xl/...``.
            target = targets[sheet.attrib[f"{{{REL}}}id"]].replace("\\", "/").lstrip("/")
            sheet_path = target if target.startswith("xl/") else "xl/" + target
            xml = ET.fromstring(archive.read(sheet_path))
            dimension = xml.find(f"{{{MAIN}}}dimension")
            merges = xml.find(f"{{{MAIN}}}mergeCells")
            rows: list[dict] = []
            sheet_data = xml.find(f"{{{MAIN}}}sheetData")
            for row in (list(sheet_data) if sheet_data is not None else [])[:30]:
                cells: dict[str, object] = {}
                for cell in row.findall(f"{{{MAIN}}}c"):
                    reference = cell.attrib.get("r", "")
                    cell_type = cell.attrib.get("t", "n")
                    value_node = cell.find(f"{{{MAIN}}}v")
                    formula_node = cell.find(f"{{{MAIN}}}f")
                    if cell_type == "inlineStr":
                        value: object = text_of(cell.find(f"{{{MAIN}}}is"))
                    elif value_node is None:
                        value = None
                    elif cell_type == "s":
                        index = int(value_node.text or 0)
                        value = shared_strings[index] if index < len(shared_strings) else index
                    elif cell_type == "b":
                        value = value_node.text == "1"
                    else:
                        raw = value_node.text or ""
                        try:
                            value = float(raw)
                            if value.is_integer():
                                value = int(value)
                        except ValueError:
                            value = raw
                    cells[reference] = {
                        "value": value,
                        "formula": formula_node.text if formula_node is not None else None,
                        "style": int(cell.attrib.get("s", 0)),
                        "type": cell_type,
                    }
                if cells:
                    rows.append({"row": int(row.attrib.get("r", 0)), "cells": cells})
            result["sheets"].append(
                {
                    "name": name,
                    "visibility": sheet.attrib.get("state", "visible"),
                    "dimension": dimension.attrib.get("ref") if dimension is not None else None,
                    "merged_ranges": [m.attrib["ref"] for m in list(merges or [])],
                    "rows": rows,
                }
            )
        return result


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) != 2:
        raise SystemExit("Usage: inspect_xlsx.py <workbook.xlsx>")
    print(json.dumps(inspect_xlsx(Path(sys.argv[1]).resolve()), ensure_ascii=False, indent=2))
