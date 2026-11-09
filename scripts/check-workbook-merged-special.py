"""Independent merged selective paste/skip-blank and numeric autofill evidence."""
from workbook_acceptance import audit_schema
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import importlib.util
import json
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("merged_special_native", ROOT / "scripts/check-workbook-copy-merges.py")
merged = importlib.util.module_from_spec(spec)
spec.loader.exec_module(merged)
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-merged-special.xlsx"
REPORT = ROOT / "docs/evidence/workbook-merged-special.json"
CASES = [{"mode": mode, "skipBlanks": skip, "target": target} for mode in ["contents", "values", "formulas", "formats"]
         for skip in [False, True] for target in ["plain", "matching"]]
CASES += [{"mode": "formats", "skipBlanks": skip, "target": "blank"} for skip in [False, True]]
CASES += [{"mode": "series", "axis": axis} for axis in ["row", "column"]]
REFERENCES = ["https://support.microsoft.com/en-au/excel/paste-options?nochrome=true",
              "https://learn.microsoft.com/en-us/office/vba/api/excel.range.pastespecial",
              "https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sheet_1_1GlobalSheetSettings.html"]


def key(case):
    return "series-" + case["axis"] if case["mode"] == "series" else case["mode"] + "-" + str(int(case["skipBlanks"])) + "-" + case["target"]


def generate():
    book = merged.axis.source.Book()
    sheet = book.sheet("Data")
    sheet.add("A1", 7)
    sheet.add("A256", kind="inlineStr", rich='<r><rPr><b/><rFont val="Noto Sans CJK SC"/></rPr><t>合并</t></r><r><rPr><i/><rFont val="Liberation Sans"/></rPr><t>😀禰󠄀</t></r>', style=9)
    sheet.add("A258", 999, formula="A1+$A$1")
    for address, value in [("D1", 88), ("E1", 98), ("D256", 88), ("H1", 88), ("L1", 10), ("L3", 20), ("O1", 10), ("Q1", 20)]:
        sheet.add(address, value)
    ranges = ["A256:B257", "D256:E257", "L1:M2", "L3:M4", "O1:P2", "Q1:R2"]
    sheet.after = '<mergeCells count="6">' + ''.join('<mergeCell ref="' + item + '"/>' for item in ranges) + '</mergeCells>'
    sheet.print_area = "$A$1:$Z$270"
    book.sheet("Other").print_area = "$A$1:$Z$270"
    return merged.axis.source.zipped(book.finish())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    data = generate()
    if not args.record:
        assert FIXTURE.read_bytes() == data, "Independent merged-special source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(data).hexdigest() and report["references"] == REFERENCES
        assert [item["case"] for item in report["cases"]] == CASES
        for item in report["cases"]:
            producer = FIXTURE.with_name("workbook-merged-special-" + key(item["case"]) + "-calc.xlsx")
            assert sha256(producer.read_bytes()).hexdigest() == item["producerSha256"] and merged.observe(producer) == item["observed"]
        print("Verified twenty independent merged-special source/producer cases.")
        return
    import uno
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == merged.axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-merged-special-native"
    directory.mkdir(exist_ok=True)
    input_path = directory / "source.xlsx"
    input_path.write_bytes(data)
    cases = []
    for case in CASES:
        def mutate(document, prop):
            controller, sheet = document.getCurrentController(), document.getSheets().getByIndex(0)
            # This changes only the disposable acceptance process. It expresses
            # the test's deliberate overwrite instead of cancelling a hidden UI
            # prompt. Always copy current formula results, not the 999 caches.
            document.createInstance("com.sun.star.sheet.GlobalSheetSettings").setPropertyValue("ReplaceCellsWarning", False)
            document.calculateAll()
            if case["mode"] == "series":
                address, direction = ("L1:M10", "TO_BOTTOM") if case["axis"] == "row" else ("O1:X2", "TO_RIGHT")
                sheet.getCellRangeByName(address).fillAuto(uno.Enum("com.sun.star.sheet.FillDirection", direction), 4)
                return
            dispatcher = uno.getComponentContext().ServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", uno.getComponentContext())
            controller.select(sheet.getCellRangeByName("A256:B258"))
            dispatcher.executeDispatch(controller.getFrame(), ".uno:Copy", "", 0, ())
            target = {"plain": "D1:E3", "matching": "D256:E258", "blank": "H1:I3"}[case["target"]]
            controller.select(sheet.getCellRangeByName(target))
            flags = {"contents": "A", "values": "SVD", "formulas": "SVDF", "formats": "T"}[case["mode"]]
            dispatcher.executeDispatch(controller.getFrame(), ".uno:InsertContents", "", 0, tuple(prop(name, value) for name, value in [
                ("Flags", flags), ("FormulaCommand", 0), ("SkipEmptyCells", case["skipBlanks"]), ("Transpose", False), ("AsLink", False), ("MoveMode", 4)]))
        output = directory / (key(case) + ".xlsx")
        version = merged.axis.native(input_path, output, None, mutate)
        producer = FIXTURE.with_name("workbook-merged-special-" + key(case) + "-calc.xlsx")
        if args.record:
            producer.write_bytes(output.read_bytes())
        cases.append({"case": case, "producerSha256": sha256(producer.read_bytes()).hexdigest(), "observed": merged.observe(output)})
    audit_schema(directory, stdout=subprocess.PIPE, text=True)
    report = {"scope": "Independent ZIP/XML source, eighteen actual selective clipboard pastes with explicit overwrite/current calculation, two actual numeric fillAuto operations, and official XSD/OPC. SDK values/formulas preserve target merge formatting; Calc removes a matching target merge. Calc formats/skip-blank content can retain populated covered cells; SDK rejects the whole candidate to protect source content. Source and producer bytes remain distinct. No Excel/WPS structural runtime claim.",
              "references": REFERENCES, "office": version, "archiveSha256": merged.axis.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(), "cases": cases}
    if args.record:
        FIXTURE.write_bytes(data)
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Pinned merged-special observations changed."
    print("Verified twenty actual pinned merged-special operations and official XSD/OPC.")


if __name__ == "__main__":
    main()
