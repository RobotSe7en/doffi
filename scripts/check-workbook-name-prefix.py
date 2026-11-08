"""Independent current-workbook name prefixes and pinned Calc copy/reopen differences."""
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
spec = importlib.util.spec_from_file_location("native_name_axis", ROOT / "scripts/check-workbook-axis.py")
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-name-prefix.xlsx"
REPORT = ROOT / "docs/evidence/workbook-name-prefix.json"
CASES = ["resave", "copy-same", "copy-other", "reopen-copy-other"]
REFERENCES = ["https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a469e5f5-a102-49bd-9642-8a8e8aaf1623",
              "https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/3d025add-118d-4413-9856-ab65712ec1b0"]


def generate():
    book = axis.source.Book()
    for index, title in enumerate(["Data", "Other"]):
        sheet = book.sheet(title)
        for address, value in [("A5", 10), ("A6", 30), ("H1", 7)]:
            sheet.add(address, value * 10 ** index)
        for address, formula, cached in [("B5", "Rate+Tax", 17 * 10 ** index), ("B6", "$A$5+A6+Local", 70 * 10 ** index),
                                         ("B7", "Tax", 7 * 10 ** index), ("B8", "[0]!Tax", 7)]:
            sheet.add(address, cached, formula=formula)
    book.names.extend([({"name": "Tax"}, "Data!$H$1"), ({"name": "Rate", "localSheetId": 0}, "Data!$A$5"),
                       ({"name": "Local", "localSheetId": 0}, "Data!$A$6"), ({"name": "Local", "localSheetId": 1}, "Other!$A$6"),
                       ({"name": "Tax", "localSheetId": 1}, "Other!$H$1")])
    return axis.source.zipped(book.finish())


def observe(path):
    with ZipFile(path) as archive:
        names = [{"name": node.get("name"), "sheet": node.get("localSheetId"), "expression": node.text}
                 for node in ET.fromstring(archive.read("xl/workbook.xml")).findall("{*}definedNames/{*}definedName") if not node.get("name").startswith("_xlnm.")]
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
        assert FIXTURE.read_bytes() == data, "Independent current-workbook name source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(data).hexdigest() and report["references"] == REFERENCES
        assert [case["name"] for case in report["cases"]] == CASES
        for case in report["cases"]:
            producer = FIXTURE.with_name("workbook-name-prefix-" + case["name"] + "-calc.xlsx")
            assert sha256(producer.read_bytes()).hexdigest() == case["producerSha256"] and observe(producer) == case["observed"]
        print("Verified four independent current-workbook name source/copy/reopen cases.")
        return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-name-prefix-native"
    directory.mkdir(exist_ok=True)
    input_path = directory / "source.xlsx"
    input_path.write_bytes(data)
    cases = []
    for name in CASES:
        def mutate(document, prop):
            if name in ["copy-same", "copy-other"]:
                source = document.getSheets().getByIndex(0)
                target = document.getSheets().getByIndex(1 if name == "copy-other" else 0)
                target.copyRange(target.getCellRangeByName("B5" if name == "copy-other" else "D5").getCellAddress(), source.getCellRangeByName("B5:B8").getRangeAddress())
        output = directory / (name + ".xlsx")
        origin = directory / "copy-other.xlsx" if name == "reopen-copy-other" else input_path
        version = axis.native(origin, output, {}, mutate)
        producer = FIXTURE.with_name("workbook-name-prefix-" + name + "-calc.xlsx")
        if args.record:
            producer.write_bytes(output.read_bytes())
        cases.append({"name": name, "producerSha256": sha256(producer.read_bytes()).hexdigest(), "observed": observe(output)})
    audit_python = sys.executable if importlib.util.find_spec("lxml") is not None else str(ROOT / ".cache/xlsx-venv/bin/python")
    subprocess.run([audit_python, str(ROOT / "scripts/check-xlsx-output.py"), str(directory)], check=True, stdout=subprocess.PIPE, text=True)
    report = {"scope": "Independent ZIP/XML [0]!Tax source follows the Microsoft current-workbook name prefix. Calc removes the prefix, resolves a local shadow on resave, clones a missing local name on copy, and loses copied global binding on a second reopen. These producer differences are retained; Excel/WPS runtime unverified.",
              "references": REFERENCES, "office": version, "archiveSha256": axis.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(),
              "sourceCurrentWorkbookNameExpected": {"Data!B8": 7, "Other!B8": 7}, "cases": cases}
    if args.record:
        FIXTURE.write_bytes(data)
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Pinned name-prefix observations changed."
    print("Verified four actual pinned Calc name resave/copy/reopen operations and official XSD/OPC.")


if __name__ == "__main__":
    main()
