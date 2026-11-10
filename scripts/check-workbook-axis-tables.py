"""Independent native tables, explicit SDK expectations and pinned Calc axis edits."""
from workbook_acceptance import audit_schema
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import importlib.util
import json
import os
import re
import subprocess
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("workbook_table_axis", ROOT / "scripts/check-workbook-axis.py")
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-axis-tables.xlsx"
REPORT = ROOT / "docs/evidence/workbook-axis-tables.json"
BASE = ROOT / "examples/xlsx/05-native-table.xlsx"
REFERENCES = ["https://support.microsoft.com/en-us/excel/resize-a-table-by-adding-or-removing-rows-and-columns-in-excel",
              "https://learn.microsoft.com/en-us/javascript/api/excel/excel.table?view=excel-js-preview"]


def case(name, axis_name, action, index, count, ref, totals, observers, columns=None, rejected=False, header=None):
    return {"name": name, "operation": {"axis": axis_name, "action": action, "index": index, "count": count},
            **({"setHeader": {"column": index, "text": header}} if header else {}),
            "expected": {"rejected": rejected, "ref": ref, "totals": totals, "columns": columns if columns is not None else ["产品", "数量", "单价", "金额"], "observerValues": observers},
            **({"producerRejection": {"code": "InvalidDocument", "reason": "constant-header-mismatch" if name != "delete-all-rows" else "unsafe-formula-syntax"}} if name in ["delete-header-row", "insert-body-column", "delete-all-rows"] else {})}


CASES = [
    case("insert-body-row", "row", "insert", 2, 1, "A1:D7", True, ["10", "240", "240"]),
    case("insert-first-row", "row", "insert", 1, 1, "A1:D7", True, ["10", "240", "240"]),
    case("insert-cross-chunk-rows", "row", "insert", 1, 255, "A1:D261", True, ["10", "240", "240"]),
    case("delete-header-row", "row", "delete", 0, 1, "A1:D6", True, ["10", "240", "240"], rejected=True),
    case("delete-body-rows", "row", "delete", 1, 2, "A1:D4", True, ["5", "160", "160"]),
    case("delete-all-data", "row", "delete", 1, 4, "A1:D2", True, ["0", "0", "0"]),
    case("delete-totals-row", "row", "delete", 5, 1, "A1:D5", False, ["10", "#REF!", "240"]),
    case("insert-body-column", "column", "insert", 2, 1, "A1:E6", True, ["10", "240", "240"], ["产品", "数量", "Column1", "单价", "金额"]),
    case("insert-named-column", "column", "insert", 2, 1, "A1:E6", True, ["10", "240", "240"], ["产品", "数量", "Inserted", "单价", "金额"], header="Inserted"),
    case("delete-quantity-column", "column", "delete", 1, 1, "A1:C6", True, ["#REF!", "#REF!", "#REF!"], ["产品", "单价", "金额"]),
    case("delete-all-rows", "row", "delete", 0, 6, None, False, ["#REF!", "#REF!", "#REF!"], []),
]


def generate():
    # Extend the unchanged independent source using only Python ZIP/XML. The
    # original fixture/evidence is never regenerated or modified by this task.
    with ZipFile(BASE) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    def insert(path, before, data):
        text = parts[path].decode("utf8")
        assert text.count(before) == 1
        parts[path] = text.replace(before, data + before).encode()
    insert("xl/workbook.xml", "</sheets>", '<sheet name="Observer" sheetId="2" r:id="observer"/>')
    insert("xl/_rels/workbook.xml.rels", "</Relationships>", f'<Relationship Id="observer" Type="{axis.source.R}/worksheet" Target="worksheets/sheet2.xml"/>')
    insert("[Content_Types].xml", "</Types>", f'<Override PartName="/xl/worksheets/sheet2.xml" ContentType="{axis.source.TYPES["worksheet"]}"/>')
    cells = ''.join(axis.source.cell(address, value, formula=formula) for address, value, formula in [
        ("A1", 10, "SUM(SalesTable[数量])"), ("B1", 240, "Sales!D6"), ("C1", 240, "SUM(SalesTable[金额])")])
    parts["xl/worksheets/sheet2.xml"] = (axis.source.XML + f'<worksheet xmlns="{axis.source.S}"><sheetFormatPr defaultRowHeight="15" defaultColWidth="12"/><sheetData><row r="1">{cells}</row></sheetData></worksheet>').encode()
    return axis.source.zipped(parts)


def observe(path):
    with ZipFile(path) as archive:
        strings = ["".join(item.itertext()) for item in ET.fromstring(archive.read("xl/sharedStrings.xml"))] if "xl/sharedStrings.xml" in archive.namelist() else []
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        cells = {cell.get("r"): cell for cell in sheet.findall("{*}sheetData/{*}row/{*}c")}
        tables = []
        for name in sorted(archive.namelist()):
            if name.startswith("xl/tables/") and name.endswith(".xml"):
                table = ET.fromstring(archive.read(name))
                start = re.match(r"([A-Z]+)(\d+)", table.get("ref"))
                first = 0
                for char in start[1]:
                    first = first * 26 + ord(char) - 64
                headers = []
                for offset, column in enumerate(table.findall("{*}tableColumns/{*}tableColumn")):
                    address = axis.source.column(first + offset) + start[2]
                    cell = cells.get(address)
                    kind = cell.get("t", "n") if cell is not None else None
                    constant = None
                    if cell is not None and cell.find("{*}f") is None:
                        if kind == "s":
                            constant = strings[int(cell.findtext("{*}v"))]
                        if kind == "inlineStr":
                            constant = "".join(cell.find("{*}is").itertext())
                    headers.append({"address": address, "type": kind, "constantText": constant})
                tables.append({"ref": table.get("ref"), "header": table.get("headerRowCount", "1"), "totals": table.get("totalsRowCount", "0"), "headerCells": headers,
                               "columns": [{"name": column.get("name"), "calculated": column.findtext("{*}calculatedColumnFormula"),
                                            "totalFormula": column.findtext("{*}totalsRowFormula")} for column in table.findall("{*}tableColumns/{*}tableColumn")]})
        formulas = {}
        for index, sheet_name in [(1, "Sales"), (2, "Observer")]:
            sheet = ET.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml"))
            formulas[sheet_name] = {cell.get("r"): {"formula": cell.findtext("{*}f"), "value": cell.findtext("{*}v"), "type": cell.get("t", "n")}
                                    for cell in sheet.findall("{*}sheetData/{*}row/{*}c") if cell.find("{*}f") is not None}
        return {"tables": tables, "formulas": formulas}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    data = generate()
    if not args.record:
        assert FIXTURE.read_bytes() == data, "Independent native table source changed."
        previous = json.loads(REPORT.read_text())
        assert previous["sourceSha256"] == sha256(data).hexdigest()
        assert [{key: item[key] for key in ["name", "operation", "expected", "setHeader", "producerRejection"] if key in item} for item in previous["cases"]] == CASES
        for item in previous["cases"]:
            assert sha256((ROOT / item["outputFile"]).read_bytes()).hexdigest() == item["outputSha256"]
        if not args.verify_native:
            print("Verified eleven independent native-table axis sources and immutable producer hashes.")
            return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    os.environ["LC_ALL"] = "C.UTF-8"  # Acceptance process only; not an SDK locale default.
    directory = ROOT / ".cache/workbook-axis-tables-native"
    directory.mkdir(exist_ok=True)
    source_path = directory / "source.xlsx"
    source_path.write_bytes(data)
    cases, outputs = [], []
    for item in CASES:
        output = directory / (item["name"] + ".xlsx")
        def mutate(document, prop):
            sheet = document.getSheets().getByIndex(0)
            operation = item["operation"]
            axes = sheet.getRows() if operation["axis"] == "row" else sheet.getColumns()
            method = axes.insertByIndex if operation["action"] == "insert" else axes.removeByIndex
            method(operation["index"], operation["count"])
            sheet.getCellByPosition(item["setHeader"]["column"], 0).setString(item["setHeader"]["text"])
        version = axis.native(source_path, output, item["operation"], mutate if "setHeader" in item else None)
        observed = observe(output)
        recorded = FIXTURE.parent / ("workbook-axis-tables-" + item["name"] + "-calc.xlsx")
        output_hash = sha256(output.read_bytes() if args.record else recorded.read_bytes()).hexdigest()
        cases.append({**item, "observed": observed, "outputFile": recorded.relative_to(ROOT).as_posix(), "outputSha256": output_hash,
                      "basis": "SDK source-identity and table-template expectations; actual Calc output remains separate. Header rejection and total-row handling do not claim Excel runtime verification."})
        outputs.append((recorded, output))
    audit_schema(directory)
    report = {"scope": "Independent source and actual pinned Calc edits; producer differences are retained, not repaired. Excel/WPS runtime unverified.",
              "references": REFERENCES, "office": version, "archiveSha256": axis.ARCHIVE_HASH, "nativeLocale": "C.UTF-8",
              "baseSourceFile": BASE.relative_to(ROOT).as_posix(), "baseSourceSha256": sha256(BASE.read_bytes()).hexdigest(), "sourceSha256": sha256(data).hexdigest(), "cases": cases}
    if args.record:
        for recorded, output in outputs:
            recorded.write_bytes(output.read_bytes())
        FIXTURE.write_bytes(data)
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert report == previous, "Pinned native table axis observations changed."
    print("Verified eleven actual pinned Calc table axis edits and official XSD/OPC checks.")


if __name__ == "__main__":
    main()
