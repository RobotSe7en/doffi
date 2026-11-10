"""Independent filter/SUBTOTAL source and pinned producer observations, without SDK output."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import argparse
import importlib.util
import json
import os
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("filter_native_probe", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec)
spec.loader.exec_module(native)
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-filter-calculation.xlsx"
REPORT = ROOT / "docs/evidence/workbook-filter-calculation.json"
CASES = [
    {"name": "ValuesManualHidden", "values": ["keep", "skip", "keep"], "amounts": [1, 10, 100], "visible": [True, False, True], "manual": [0],
     "filter": '<filters><filter val="keep"/></filters>', "expected": [101, 100]},
    {"name": "CustomBounds", "values": [1, 2, 2, 4], "visible": [False, True, True, False],
     "filter": '<customFilters and="1"><customFilter operator="greaterThanOrEqual" val="2"/><customFilter operator="lessThan" val="4"/></customFilters>', "expected": [4, 4]},
    {"name": "TopTies", "values": [1, 2, 2, 4], "visible": [False, True, True, True],
     "filter": '<top10 top="1" percent="0" val="2" filterVal="2"/>', "expected": [8, 8]},
    {"name": "BottomTies", "values": [1, 2, 2, 4], "visible": [True, True, True, False],
     "filter": '<top10 top="0" percent="0" val="2" filterVal="2"/>', "expected": [5, 5]},
    {"name": "TopPercent", "values": [1, 2, 2, 4], "visible": [False, False, False, True],
     "filter": '<top10 top="1" percent="1" val="25" filterVal="4"/>', "expected": [4, 4]},
    {"name": "AboveAverage", "values": [1, 2, 2, 4], "visible": [False, False, False, True],
     "filter": '<dynamicFilter type="aboveAverage" val="2.25"/>', "expected": [4, 4]},
    {"name": "BelowAverage", "values": [1, 2, 2, 4], "visible": [True, True, True, False],
     "filter": '<dynamicFilter type="belowAverage" val="2.25"/>', "expected": [5, 5]},
    {"name": "EscapedWildcard", "values": ["alpha", "a*b", "abc", "beta"], "amounts": [1, 2, 3, 4], "visible": [False, True, False, False],
     "filter": '<customFilters><customFilter operator="equal" val="a~*b"/></customFilters>', "expected": [2, 2]},
]


def generate(cached=True):
    book = native.fixture.Book()
    for case in CASES:
        sheet = book.sheet(case["name"])
        sheet.add("A1", "Filter")
        sheet.add("B1", "Amount")
        for index, value in enumerate(case["values"]):
            row = index + 2
            sheet.add("A" + str(row), value)
            sheet.add("B" + str(row), case.get("amounts", case["values"])[index])
            if not case["visible"][index] or index in case.get("manual", []):
                sheet.row_props[row] = {"hidden": 1}
        last = len(case["values"]) + 1
        sheet.after = f'<autoFilter ref="A1:B{last}"><filterColumn colId="0">{case["filter"]}</filterColumn></autoFilter>'
        for row, code in [(1, 9), (2, 109)]:
            sheet.add("D" + str(row), 999 if cached else None, formula=f"SUBTOTAL({code},B2:B{last})", formula_attrs={"ca": "1"})
    parts = book.finish()
    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace('fullCalcOnLoad="0"', 'fullCalcOnLoad="1" forceFullCalc="1"')
    return native.fixture.zipped(parts)


def observe(path):
    cases = []
    with ZipFile(path) as archive:
        for index, case in enumerate(CASES, 1):
            root = ET.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml"))
            cells = {cell.get("r"): cell for cell in root.findall("{*}sheetData/{*}row/{*}c")}
            values = [float(cells[f"D{row}"].findtext("{*}v")) for row in [1, 2]]
            basis = "Pinned Calc observation; Excel/WPS runtime unverified" if values == case["expected"] else "ECMA-376 18.17.7.305 and Microsoft SUBTOTAL: codes 1-11 include manually hidden rows, codes 101-111 exclude them. Native row/@hidden has no separate manual/filter reason; Calc excludes the manually hidden matching row in both codes on this input. Current-source criteria and original hidden flags remain separate."
            cases.append({"sheet": case["name"], "expected": case["expected"], "observed": values, "basis": basis,
                          "matchesProducer": values == case["expected"]})
    return cases


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    source, reference = generate(), generate(cached=False)
    if not args.record and not args.verify_native:
        assert FIXTURE.read_bytes() == source
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(source).hexdigest()
        assert [case["expected"] for case in report["cases"]] == [case["expected"] for case in CASES]
        print(f"Verified {len(CASES) * 2} independent filter/SUBTOTAL sources")
        return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    assert native.shared.run(["libreoffice", "--version"]) == native.BUILD
    directory = ROOT / ".cache/workbook-filter-calculation-native"
    directory.mkdir(exist_ok=True)
    with TemporaryDirectory(dir=directory) as temporary:
        work = Path(temporary)
        output = work / "out"
        output.mkdir()
        path = work / "filters.xlsx"
        path.write_bytes(reference)
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(output), str(path)])
        cases = observe(output / "filters.xlsx")
    report = {"scope": "Independent current-filter definitions, explicit native hidden rows, and SUBTOTAL 9/109. Source caches deliberately wrong. Dynamic dates use separate explicit host-clock unit cases. Excel/WPS runtime unverified.",
              "basis": "ECMA-376 Part 1 18.3.2, 18.17.7.305 and 18.18.26; current filtering is independent of native cached visibility",
              "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH,
              "sourceSha256": sha256(source).hexdigest(), "referenceInputSha256": sha256(reference).hexdigest(), "cases": cases}
    if args.record:
        FIXTURE.write_bytes(source)
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert FIXTURE.read_bytes() == source
        assert json.loads(REPORT.read_text()) == report, "Native filter observations changed"
    print(f"Verified {len(CASES) * 2} independent filter/SUBTOTAL observations")


if __name__ == "__main__":
    main()
