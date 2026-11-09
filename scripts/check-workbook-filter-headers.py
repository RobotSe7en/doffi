"""Independent ordinary-filter header deletions and actual pinned Calc edits."""
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import importlib.util
import json
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("filter_header_axis", ROOT / "scripts/check-workbook-axis.py")
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-filter-headers.xlsx"
REPORT = ROOT / "docs/evidence/workbook-filter-headers.json"
CASES = [
    {"name": "header", "index": 1, "count": 1, "range": "B2:D5", "expected": {"type": "number", "value": 80}},
    {"name": "above-header-data", "index": 0, "count": 3, "range": "B1:D3", "expected": {"type": "number", "value": 90}},
    {"name": "all", "index": 1, "count": 5, "range": None, "expected": {"type": "error", "value": "#REF!"}},
    {"name": "header-most-data", "index": 1, "count": 4, "range": "B2:D2", "expected": {"type": "number", "value": 40}},
]


def generate():
    book = axis.source.Book()
    sheet = book.sheet("Data")
    for column, value in [("B", "Item"), ("C", "Count"), ("D", "Amount")]:
        sheet.add(column + "2", value)
    for row, item, count, amount in [(3, "keep", 1, 10), (4, "skip", 2, 20), (5, "keep", 3, 30), (6, "keep", 4, 40)]:
        for column, value in [("B", item), ("C", count), ("D", amount)]:
            sheet.add(column + str(row), value)
    sheet.after = '<autoFilter ref="B2:D6"><filterColumn colId="0"><filters><filter val="keep"/></filters></filterColumn><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="1"/></customFilters></filterColumn></autoFilter>'
    observer = book.sheet("Observer")
    observer.add("A1", 999, formula="SUBTOTAL(9,Data!D3:D6)")
    return axis.source.zipped(book.finish())


def observe(path):
    with ZipFile(path) as archive:
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        filtered = sheet.find("{*}autoFilter")
        observer = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
        cell = observer.find("{*}sheetData/{*}row/{*}c")
        value = cell.findtext("{*}v")
        return {"range": filtered.get("ref") if filtered is not None else None,
                "columns": [{"column": int(column.get("colId")), "values": [item.get("val") for item in column.findall("{*}filters/{*}filter")],
                             "tests": [item.attrib for item in column.findall("{*}customFilters/{*}customFilter")]}
                            for column in filtered.findall("{*}filterColumn")] if filtered is not None else [],
                "hiddenRows": [int(row.get("r")) - 1 for row in sheet.findall("{*}sheetData/{*}row") if row.get("hidden") in ["1", "true"]],
                "formula": cell.findtext("{*}f"), "cached": {"type": "error", "value": value} if cell.get("t") == "e" else {"type": "number", "value": float(value)}}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    data = generate()
    if not args.record:
        assert FIXTURE.read_bytes() == data, "Independent filter-header source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(data).hexdigest()
        assert report["operations"] == CASES
        for case in report["cases"]:
            path = ROOT / case["outputFile"]
            assert sha256(path.read_bytes()).hexdigest() == case["outputSha256"]
            assert observe(path) == case["observed"]
        print("Verified independent filter-header source and eight native observations.")
        return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-filter-headers-native"
    directory.mkdir(exist_ok=True)
    source = directory / "source.xlsx"
    source.write_bytes(data)
    cases = []
    for case in CASES:
        for reapply in [False, True]:
            name = case["name"] + ("-reapply" if reapply else "-delete")
            output = directory / (name + ".xlsx")

            def mutate(document, prop):
                sheet = document.getSheets().getByIndex(0)
                sheet.getRows().removeByIndex(case["index"], case["count"])
                if reapply and case["range"]:
                    import uno
                    from com.sun.star.sheet.FilterOperator import EQUAL, GREATER
                    from com.sun.star.sheet.FilterConnection import AND
                    target = sheet.getCellRangeByName(case["range"])
                    descriptor = target.createFilterDescriptor(True)
                    descriptor.ContainsHeader = True
                    descriptor.UseRegularExpressions = False
                    first = uno.createUnoStruct("com.sun.star.sheet.TableFilterField")
                    first.Field, first.Operator, first.IsNumeric, first.StringValue = 0, EQUAL, False, "keep"
                    second = uno.createUnoStruct("com.sun.star.sheet.TableFilterField")
                    second.Connection, second.Field, second.Operator, second.IsNumeric, second.NumericValue = AND, 1, GREATER, True, 1
                    descriptor.setFilterFields((first, second))
                    target.filter(descriptor)

            version = axis.native(source, output, None, mutate)
            recorded = FIXTURE.parent / ("workbook-filter-headers-" + name + "-calc.xlsx")
            if args.record:
                shutil.copyfile(output, recorded)
            observed = observe(output)
            assert observe(recorded) == observed, "Immutable producer observation changed."
            assert observed["range"] == case["range"]
            if reapply:
                assert observed["cached"] == case["expected"], (name, observed)
            cases.append({"name": name, "operation": case["name"], "reapply": reapply, "observed": observed,
                          "matchesCurrent": observed["cached"] == case["expected"], "outputFile": str(recorded.relative_to(ROOT)),
                          "outputSha256": sha256(recorded.read_bytes()).hexdigest()})
    audit = ROOT / ".cache/xlsx-venv/bin/python"
    subprocess.run([str(audit) if audit.is_file() else sys.executable, str(ROOT / "scripts/check-xlsx-output.py"), str(directory)], check=True)
    report = {"scope": "Independent ZIP/XML source, actual row deletions and separate explicit criterion reapplication in pinned Calc. Native cached visibility is not current filtering. Excel/WPS runtime unverified.",
              "compatibility": "The first surviving row is the resulting ordinary filter header. The Microsoft reference documents filter range ownership, not physical header deletion; runtime evidence here is Calc only.",
              "basis": ["https://learn.microsoft.com/en-us/office/vba/api/excel.autofilter.range",
                        "https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XSheetFilterable.html"],
              "office": version, "archiveSha256": axis.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(), "operations": CASES, "cases": cases}
    if args.record:
        FIXTURE.write_bytes(data)
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Pinned filter-header observations changed."
    print("Verified eight actual pinned Calc filter-header edits and official XSD/OPC.")


if __name__ == "__main__":
    main()
