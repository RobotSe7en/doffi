"""Whole native table cuts, independently produced by pinned Calc moveRange."""
from workbook_acceptance import audit_schema
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import importlib.util
import json
import subprocess
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("move_table_source", ROOT / "scripts/check-workbook-axis-tables.py")
source = importlib.util.module_from_spec(spec)
spec.loader.exec_module(source)
axis = source.axis
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-axis-tables.xlsx"
REPORT = ROOT / "docs/evidence/workbook-move-definitions.json"
CASES = ["same-sheet", "other-sheet"]


def observe(path):
    with ZipFile(path) as archive:
        tables = []
        for name in sorted(archive.namelist()):
            if name.startswith("xl/tables/") and name.endswith(".xml"):
                table = ET.fromstring(archive.read(name))
                tables.append({"name": table.get("name"), "range": table.get("ref"), "columns": [column.get("name") for column in table.findall("{*}tableColumns/{*}tableColumn")],
                               "filter": table.find("{*}autoFilter").get("ref") if table.find("{*}autoFilter") is not None else None})
        sheets = []
        for index in [1, 2]:
            sheet = ET.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml"))
            sheets.append({"tables": len(sheet.findall("{*}tableParts/{*}tablePart")),
                           "formulas": [{"address": cell.get("r"), "expression": cell.findtext("{*}f"), "type": cell.get("t", "n"), "value": cell.findtext("{*}v")}
                                        for cell in sheet.findall("{*}sheetData/{*}row/{*}c") if cell.find("{*}f") is not None]})
        return {"tables": tables, "sheets": sheets}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    data = FIXTURE.read_bytes()
    assert data == source.generate(), "The reused independent table source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(data).hexdigest()
        assert [case["name"] for case in report["cases"]] == CASES
        for case in report["cases"]:
            producer = FIXTURE.with_name("workbook-move-definitions-" + case["name"] + "-calc.xlsx")
            assert sha256(producer.read_bytes()).hexdigest() == case["producerSha256"]
            assert observe(producer) == case["observed"]
        print("Verified two independent whole-table cut source and producer cases.")
        return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-move-definitions-native"
    directory.mkdir(exist_ok=True)
    input_path = directory / "source.xlsx"
    input_path.write_bytes(data)
    cases = []
    for name in CASES:
        def mutate(document, prop):
            sheet = document.getSheets().getByIndex(0)
            target = document.getSheets().getByIndex(1 if name == "other-sheet" else 0)
            target.moveRange(target.getCellRangeByName("H260").getCellAddress(), sheet.getCellRangeByName("A1:D6").getRangeAddress())
        output = directory / (name + ".xlsx")
        version = axis.native(input_path, output, {}, mutate)
        producer = FIXTURE.with_name("workbook-move-definitions-" + name + "-calc.xlsx")
        if args.record:
            producer.write_bytes(output.read_bytes())
        cases.append({"name": name, "producerSha256": sha256(producer.read_bytes()).hexdigest(), "observed": observe(output)})
    audit_schema(directory, stdout=subprocess.PIPE, text=True)
    report = {"scope": "Independent ZIP/XML source reused without rewriting, actual pinned Calc whole-table moveRange. SDK preserves its source data-only filter and calculated-column recipe; Calc includes totals in its filter and drops the stored recipe. Excel/WPS runtime unverified.",
              "office": version, "archiveSha256": axis.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(), "cases": cases}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Pinned whole-table cut observations changed."
    print("Verified two actual pinned Calc whole-table cuts and official XSD/OPC.")


if __name__ == "__main__":
    main()
