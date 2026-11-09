"""Independent ZIP/XML sorting source, pinned native edits and explicit producer differences."""
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
spec = importlib.util.spec_from_file_location("workbook_native_sort", ROOT / "scripts/check-workbook-axis.py")
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-sort.xlsx"
REPORT = ROOT / "docs/evidence/workbook-sort.json"
BASIS = "https://support.microsoft.com/en-us/excel/sort-data-in-a-range-or-table-in-excel"
ASCENDING = [2, 3, 1, 16, 17, 6, 5, 4, 12, 13, 14, 15, 7, 8, 9, 10, 11]
CASES = [
    {"name": "numeric-ascending", "sheet": "Numeric", "keys": [{"column": 0, "descending": False}], "caseSensitive": False,
     "expected": [13, 12, 11, 2, 3, 14, 7, 15, 1, 16, 8, 17, 6, 5, 4, 10, 9]},
    {"name": "numeric-descending", "sheet": "Numeric", "keys": [{"column": 0, "descending": True}], "caseSensitive": False,
     "expected": [9, 10, 4, 5, 6, 17, 8, 16, 1, 15, 7, 14, 2, 3, 11, 12, 13]},
    {"name": "ascending", "sheet": "Data", "keys": [{"column": 0, "descending": False}], "caseSensitive": False, "expected": ASCENDING},
    {"name": "descending", "sheet": "Data", "keys": [{"column": 0, "descending": True}], "caseSensitive": False,
     "expected": [9, 10, 8, 7, 14, 15, 13, 12, 4, 5, 6, 16, 17, 1, 2, 3, 11]},
    {"name": "second-key", "sheet": "Data", "keys": [{"column": 0, "descending": False}, {"column": 1, "descending": False}], "caseSensitive": False,
     "expected": [3, 2, 1, 17, 16, 6, 5, 4, 12, 13, 15, 14, 7, 8, 10, 9, 11]},
    {"name": "case-sensitive", "sheet": "Data", "keys": [{"column": 0, "descending": False}], "caseSensitive": True,
     "expected": [2, 3, 1, 16, 12, 13, 17, 6, 5, 4, 14, 15, 7, 8, 9, 10, 11]},
]


def generate():
    book = axis.source.Book()
    sheet = book.sheet("Data")
    for address, label in [("A1", "Key"), ("B1", "Secondary"), ("C1", "Identity")]:
        sheet.add(address, label, style=1)
    values = [3, 1, 1, "A11", "A100", "A1", False, True, "#REF!", "#DIV/0!", None, "ab", "a-b", "O'Neil", "ONeil", "a", "A"]
    for identity, value in enumerate(values, 1):
        row = identity + 1
        if value is not None:
            sheet.add("A" + str(row), value, **({"kind": "e"} if identity in [9, 10] else {}))
        sheet.add("B" + str(row), 18 - identity)
        sheet.add("C" + str(row), f"R{identity:02}")
    sheet.add("F1", 999, formula="A2")
    observer = book.sheet("Observer")
    observer.add("A1", 999, formula="Data!A2")
    observer.add("B1", 999, formula="Numeric!A2")
    numeric = book.sheet("Numeric")
    for address, label in [("A1", "Key"), ("B1", "Secondary"), ("C1", "Identity")]:
        numeric.add(address, label, style=1)
    for identity, value in enumerate([30, 10, 10, 70, 60, 50, 20, 40, 90, 80, 0, -10, -20, 15, 25, 35, 45], 1):
        numeric.add("A" + str(identity + 1), value)
        numeric.add("B" + str(identity + 1), 18 - identity)
        numeric.add("C" + str(identity + 1), f"R{identity:02}")
    numeric.add("F1", 999, formula="A2")
    return axis.source.zipped(book.finish())


def observe(path, case):
    with ZipFile(path) as archive:
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            strings = ["".join(item.itertext()) for item in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
        sheet = ET.fromstring(archive.read("xl/worksheets/" + ("sheet1.xml" if case["sheet"] == "Data" else "sheet3.xml")))
        cells = {cell.get("r"): cell for cell in sheet.findall("{*}sheetData/{*}row/{*}c")}

        def text(cell):
            if cell.get("t") == "s":
                return strings[int(cell.findtext("{*}v"))]
            if cell.get("t") == "inlineStr":
                return "".join(cell.find("{*}is").itertext())
            return cell.findtext("{*}v")

        rows = [int(text(cells[f"C{row}"])[1:]) for row in range(2, 19)]
        references = {case["sheet"] + "!F1": cells["F1"].findtext("{*}f")}
        other = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
        reference = "A1" if case["sheet"] == "Data" else "B1"
        references["Observer!" + reference] = next(cell.findtext("{*}f") for cell in other.findall("{*}sheetData/{*}row/{*}c") if cell.get("r") == reference)
        return {"identities": rows, "outsideFormulas": references}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    source = generate()
    if not args.record:
        assert FIXTURE.read_bytes() == source, "Independent sorting source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(source).hexdigest()
        assert [{key: case[key] for key in ["name", "sheet", "keys", "caseSensitive", "expected"]} for case in report["cases"]] == CASES
        for case in report["cases"]:
            assert sha256((ROOT / case["outputFile"]).read_bytes()).hexdigest() == case["outputSha256"]
        print("Verified six independent sorting sources and declared expectations.")
        return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    import uno
    from com.sun.star.table.TableSortFieldType import AUTOMATIC
    directory = ROOT / ".cache/workbook-sort-native"
    directory.mkdir(exist_ok=True)
    source_path = directory / "source.xlsx"
    source_path.write_bytes(source)
    cases = []
    for case in CASES:
        output = directory / (case["name"] + ".xlsx")

        def mutate(document, prop):
            fields = []
            for key in case["keys"]:
                field = uno.createUnoStruct("com.sun.star.table.TableSortField")
                field.Field, field.IsAscending, field.IsCaseSensitive, field.FieldType = key["column"], not key["descending"], case["caseSensitive"], AUTOMATIC
                field.CollatorLocale = uno.createUnoStruct("com.sun.star.lang.Locale", "en", "US", "")
                fields.append(field)
            target = document.getSheets().getByName(case["sheet"]).getCellRangeByName("A2:C18")
            target.sort((prop("SortFields", uno.Any("[]com.sun.star.table.TableSortField", tuple(fields))), prop("IsSortColumns", False), prop("ContainsHeader", False),
                         prop("BindFormatsToContent", True), prop("IsCaseSensitive", case["caseSensitive"])))
            assert [row[2] for row in target.getDataArray()] != [f"R{identity:02}" for identity in range(1, 18)], "Native sort descriptor did not reorder the source."

        version = axis.native(source_path, output, None, mutate)
        observed = observe(output, case)
        recorded = FIXTURE.parent / ("workbook-sort-" + case["name"] + "-calc.xlsx")
        if args.record:
            recorded.write_bytes(output.read_bytes())
        cases.append({**case, "observed": observed, "matchesProducer": observed["identities"] == case["expected"],
                      "outputFile": recorded.relative_to(ROOT).as_posix(), "outputSha256": sha256(recorded.read_bytes()).hexdigest(),
                      "basis": "Microsoft documented sort orders; stable equal-key records; outside formulas stay at their addresses by user decision. Actual pinned Calc output is retained separately; no Excel/WPS runtime claim."})
    # Committed producer bytes/hashes are immutable reference artifacts. A new
    # run's timestamps may differ; compare its observed semantics separately.
    audit_schema(directory)
    report = {"scope": "Independent input and actual pinned Calc sort edits. Microsoft documented expectations; Excel/WPS runtime unverified.",
              "office": version, "archiveSha256": axis.ARCHIVE_HASH, "sourceSha256": sha256(source).hexdigest(), "reference": BASIS, "cases": cases}
    if args.record:
        FIXTURE.write_bytes(source)
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        prior = json.loads(REPORT.read_text())
        assert report == prior, "Pinned native sorting observations changed."
    print("Verified six actual pinned Calc sorting edits and official XSD/OPC checks.")


if __name__ == "__main__":
    main()
