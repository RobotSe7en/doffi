"""Independent merged-sort sources, actual pinned Calc edits and Office's size rule."""
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
spec = importlib.util.spec_from_file_location("merged_sort_axis", ROOT / "scripts/check-workbook-axis.py")
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-merged-sort.xlsx"
REPORT = ROOT / "docs/evidence/workbook-merged-sort.json"
BASIS = "https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/error-sort-a-range-merged-cell"
CASES = [
    {"name": "uniform-stable", "sheet": "Uniform", "height": 2, "columns": [0], "descending": [False], "expectedRecords": [1, 2, 0]},
    {"name": "uniform-multikey", "sheet": "Uniform", "height": 2, "columns": [0, 2], "descending": [False, True], "expectedRecords": [2, 1, 0]},
    {"name": "horizontal-stable", "sheet": "Horizontal", "height": 1, "columns": [0], "descending": [False], "expectedRecords": [1, 2, 0]},
    {"name": "mixed-rejected", "sheet": "Mixed", "height": 1, "columns": [0], "descending": [False], "reject": True},
]


def generate():
    book = axis.source.Book()
    for name, height in [("Uniform", 2), ("Horizontal", 1), ("Mixed", 1)]:
        sheet = book.sheet(name)
        merges = []
        for index, (primary, secondary) in enumerate([(3, 30), (1, 20), (1, 25)]):
            row = index * height + 1
            sheet.add("A" + str(row), 999, formula=str(primary))
            sheet.add("C" + str(row), secondary)
            for column in ["A", "C"] if name != "Mixed" else ["A"]:
                merges.append(f'{column}{row}:{chr(ord(column) + 1)}{row + height - 1}')
        sheet.after = '<mergeCells count="' + str(len(merges)) + '">' + ''.join('<mergeCell ref="' + value + '"/>' for value in merges) + '</mergeCells>'
        sheet.add("F1", 999, formula="A1")
    observer = book.sheet("Observer")
    for column, name in [("A", "Uniform"), ("B", "Horizontal"), ("C", "Mixed")]:
        observer.add(column + "1", 999, formula=name + "!A1")
    return axis.source.zipped(book.finish())


def observe(path, case):
    with ZipFile(path) as archive:
        index = {"Uniform": 1, "Horizontal": 2, "Mixed": 3}[case["sheet"]]
        sheet = ET.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml"))
        cells = {cell.get("r"): cell for cell in sheet.findall("{*}sheetData/{*}row/{*}c")}
        records = [[cells.get(column + str(row)).findtext("{*}v") if cells.get(column + str(row)) is not None else None for column in ["A", "C"]]
                   for row in range(1, case["height"] * 3 + 1, case["height"])]
        return {"records": records, "merges": [item.get("ref") for item in sheet.findall("{*}mergeCells/{*}mergeCell")],
                "outsideFormula": cells["F1"].findtext("{*}f")}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    source = generate()
    if not args.record:
        assert FIXTURE.read_bytes() == source, "Independent merged-sort source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(source).hexdigest()
        assert [{key: value for key, value in case.items() if key not in ["nativeCall", "observed", "outputFile", "outputSha256"]} for case in report["cases"]] == CASES
        for case in report["cases"]:
            path = ROOT / case["outputFile"]
            assert sha256(path.read_bytes()).hexdigest() == case["outputSha256"]
            assert observe(path, case) == case["observed"]
        print("Verified four independent merged-sort cases and immutable producer observations.")
        return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    import uno
    from com.sun.star.table.TableSortFieldType import AUTOMATIC
    directory = ROOT / ".cache/workbook-merged-sort-native"
    directory.mkdir(exist_ok=True)
    source_path = directory / "source.xlsx"
    source_path.write_bytes(source)
    cases = []
    for case in CASES:
        output = directory / (case["name"] + ".xlsx")
        call = {"outcome": "returned"}

        def mutate(document, prop):
            document.calculateAll()
            fields = []
            for column, descending in zip(case["columns"], case["descending"]):
                field = uno.createUnoStruct("com.sun.star.table.TableSortField")
                field.Field, field.IsAscending, field.FieldType = column, not descending, AUTOMATIC
                fields.append(field)
            target = document.getSheets().getByName(case["sheet"]).getCellRangeByName("A1:D" + str(case["height"] * 3))
            before = target.getDataArray()
            try:
                target.sort((prop("SortFields", uno.Any("[]com.sun.star.table.TableSortField", tuple(fields))), prop("IsSortColumns", False),
                             prop("ContainsHeader", False), prop("BindFormatsToContent", True)))
            except Exception as error:
                call.update(outcome="rejected", error=type(error).__name__)
            call["changed"] = target.getDataArray() != before

        version = axis.native(source_path, output, None, mutate)
        recorded = FIXTURE.parent / ("workbook-merged-sort-" + case["name"] + "-calc.xlsx")
        if args.record:
            recorded.write_bytes(output.read_bytes())
        cases.append({**case, "nativeCall": call, "observed": observe(output, case), "outputFile": recorded.relative_to(ROOT).as_posix(), "outputSha256": sha256(recorded.read_bytes()).hexdigest()})
    audit_schema(directory)
    report = {"scope": "Independent ZIP/XML and four actual pinned Calc sort calls. Microsoft documents equal merged row/column sizes; record ordering expectations are declared independently. Excel/WPS editing runtime remains unverified.",
              "reference": BASIS, "office": version, "archiveSha256": axis.ARCHIVE_HASH, "sourceSha256": sha256(source).hexdigest(), "cases": cases}
    if args.record:
        FIXTURE.write_bytes(source)
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert report == json.loads(REPORT.read_text()), "Pinned merged-sort producer observations changed."
    print("Verified four actual pinned Calc merged-sort calls and official XSD/OPC.")


if __name__ == "__main__":
    main()
