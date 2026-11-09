"""Independent relative-name sources and actual pinned Calc axis edits."""
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
spec = importlib.util.spec_from_file_location("name_axis_native", ROOT / "scripts/check-workbook-axis.py")
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-axis-names.xlsx"
REPORT = ROOT / "docs/evidence/workbook-axis-names.json"
CASES = [("row", "insert", 0), ("row", "insert", 6), ("row", "delete", 0), ("column", "insert", 0), ("column", "delete", 0)]
REFERENCES = ["https://support.microsoft.com/en-us/excel/use-the-name-manager-in-excel",
              "https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/36f59210-dd88-44c8-a24c-95b33d7742d5"]


def generate():
    book = axis.source.Book()
    book.names = [({"name": "Relative"}, "Data!$A1"), ({"name": "Absolute"}, "Data!$A$1"),
                  ({"name": "Local", "localSheetId": 0}, "!$A$1"), ({"name": "Dynamic"}, "!$A$1"),
                  ({"name": "Alias"}, "Relative+10"), ({"name": "WindowSum"}, "SUM(Data!$A1:$A3)"),
                  ({"name": "Mixed"}, "Data!A$1"), ({"name": "Columns"}, "Data!A:C"),
                  ({"name": "Rows"}, "Data!$1:3")]
    for index, title in enumerate(["Data", "Other"]):
        sheet = book.sheet(title)
        for row in range(1, 21):
            sheet.add("A" + str(row), (10 if index == 0 else 100) + row)
        for address, value in [("B1", 20), ("C1", 30), ("D1", 40)]:
            sheet.add(address, value * (index + 1))
        for address, formula in [("C5", "Relative"), ("C10", "Relative"), ("D5", "Absolute"),
                                 ("F5", "Dynamic"), ("G5", "Alias"), ("H5", "WindowSum"), ("D10", "Mixed")]:
            sheet.add(address, 999, formula=formula)
        if index == 0:
            sheet.add("E5", 999, formula="Local")
    return axis.source.zipped(book.finish())


def observe(path):
    with ZipFile(path) as archive:
        names = [{"name": item.get("name"), "sheet": item.get("localSheetId"), "expression": item.text}
                 for item in ET.fromstring(archive.read("xl/workbook.xml")).findall("{*}definedNames/{*}definedName")]
        sheets = []
        for index in [1, 2]:
            sheet = ET.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml"))
            sheets.append([{"address": cell.get("r"), "expression": cell.findtext("{*}f"), "type": cell.get("t", "n"), "value": cell.findtext("{*}v")}
                           for cell in sheet.findall("{*}sheetData/{*}row/{*}c") if cell.find("{*}f") is not None])
        return {"names": names, "sheets": sheets}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    data = generate()
    if not args.record:
        assert FIXTURE.read_bytes() == data, "Independent name-axis source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(data).hexdigest() and report["references"] == REFERENCES
        assert [item["case"] for item in report["cases"]] == [list(case) for case in CASES]
        for item in report["cases"]:
            producer = FIXTURE.with_name("workbook-axis-names-" + "-".join(map(str, item["case"])) + "-calc.xlsx")
            assert sha256(producer.read_bytes()).hexdigest() == item["producerSha256"]
            assert observe(producer) == item["observed"]
        print("Verified five independent relative-name axis source/producer cases.")
        return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-axis-names-native"
    directory.mkdir(exist_ok=True)
    input_path = directory / "source.xlsx"
    input_path.write_bytes(data)
    cases = []
    for direction, action, index in CASES:
        case = [direction, action, index]
        output = directory / ("-".join(map(str, case)) + ".xlsx")
        version = axis.native(input_path, output, {"axis": direction, "action": action, "index": index, "count": 1})
        producer = FIXTURE.with_name("workbook-axis-names-" + "-".join(map(str, case)) + "-calc.xlsx")
        if args.record:
            producer.write_bytes(output.read_bytes())
        cases.append({"case": case, "producerSha256": sha256(producer.read_bytes()).hexdigest(), "observed": observe(output)})
    audit_schema(directory, stdout=subprocess.PIPE, text=True)
    report = {"scope": "Independent Python ZIP/XML names, actual pinned Calc row/column insertions/deletions and official XSD/OPC. Relative name offsets stay recipes while bound absolute components follow cells. Calc's bang interpretation and nested Alias=Relative+10 cache of 10 after column edits are retained differences; the latter loses the calling coordinate (Relative itself is 15 after insertion and #REF! after deletion). Neither is adopted as Excel behavior. Excel web evidence covers calculation only; Excel/WPS structural runtime remains unverified.",
              "references": REFERENCES, "office": version, "archiveSha256": axis.ARCHIVE_HASH,
              "sourceSha256": sha256(data).hexdigest(), "cases": cases}
    if args.record:
        FIXTURE.write_bytes(data)
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Pinned name-axis observations changed."
    print("Verified five pinned native name-axis edits and official XSD/OPC.")


if __name__ == "__main__":
    main()
