"""Pinned native structured-reference fill observations, including Calc differences."""
from hashlib import sha256
from pathlib import Path
import argparse
import importlib.util
import json
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("structured_fill_tables", ROOT / "scripts/check-workbook-axis-tables.py")
base = importlib.util.module_from_spec(spec)
spec.loader.exec_module(base)
SOURCE = ROOT / "packages/office/tests/fixtures/workbook-axis-tables.xlsx"
REPORT = ROOT / "docs/evidence/workbook-structured-fill.json"
CASES = [
    {"name": "quantity-right", "seed": "H3", "range": "H3:O3", "direction": "TO_RIGHT", "formula": "=SUM(SalesTable[数量])", "sdk": [10, 100, 240, 0, 10, 100, 240, 0]},
    {"name": "amount-right", "seed": "H5", "range": "H5:O5", "direction": "TO_RIGHT", "formula": "=SUM(SalesTable[金额])", "sdk": [240, 0, 10, 100, 240, 0, 10, 100]},
    {"name": "quantity-left", "seed": "O7", "range": "H7:O7", "direction": "TO_LEFT", "formula": "=SUM(SalesTable[数量])", "sdk": [100, 240, 0, 10, 100, 240, 0, 10]},
    {"name": "range-right", "seed": "H9", "range": "H9:O9", "direction": "TO_RIGHT", "formula": "=SUM(SalesTable[[数量]:[单价]])", "sdk": [110] * 8},
    {"name": "locked-right", "seed": "H11", "range": "H11:O11", "direction": "TO_RIGHT", "formula": "=SUM(SalesTable[[数量]:[数量]])", "sdk": [10] * 8},
    {"name": "quantity-down", "seed": "H13", "range": "H13:H20", "direction": "TO_BOTTOM", "formula": "=SUM(SalesTable[数量])", "sdk": [10] * 8},
    {"name": "all-right", "seed": "H22", "range": "H22:O22", "direction": "TO_RIGHT", "formula": "=SUM(SalesTable[[#All],[数量]])", "sdk": [10, 100, 480, 0, 10, 100, 480, 0]},
]
REFERENCES = ["https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables"]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    data = SOURCE.read_bytes()
    assert data == base.generate(), "Independent source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(data).hexdigest() and report["cases"] == CASES and report["references"] == REFERENCES
        producer = ROOT / report["producer"]
        assert report["producerSha256"] == sha256(producer.read_bytes()).hexdigest() and report["observed"] == base.observe(producer)
        print("Verified seven independent structured-fill records and explicit producer differences.")
        return
    import uno
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == base.axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-structured-fill-native"
    directory.mkdir(exist_ok=True)
    output = directory / "filled.xlsx"

    def mutate(document, prop):
        sheet = document.getSheets().getByIndex(1)
        for case in CASES:
            sheet.getCellRangeByName(case["seed"]).setFormula(case["formula"])
            sheet.getCellRangeByName(case["range"]).fillAuto(uno.Enum("com.sun.star.sheet.FillDirection", case["direction"]), 1)
        document.calculateAll()

    version = base.axis.native(SOURCE, output, None, mutate)
    audit_python = sys.executable if importlib.util.find_spec("lxml") is not None else str(ROOT / ".cache/xlsx-venv/bin/python")
    subprocess.run([audit_python, str(ROOT / "scripts/check-xlsx-output.py"), str(directory)], check=True, stdout=subprocess.PIPE, text=True)
    producer = SOURCE.with_name("workbook-structured-fill-calc.xlsx")
    if args.record:
        producer.write_bytes(output.read_bytes())
    report = {"scope": "Independent ZIP/XML table source and seven actual pinned Calc fillAuto cases with official XSD/OPC. Microsoft documents that ordinary copy retains structured columns and horizontal fill adjusts fully qualified column specifiers as a series. SDK single-column fill cycles through the table in both directions; explicit column ranges retain their bindings. Those boundary/range details are implementation behavior, not independently observed Excel acceptance. Calc retains all original structured columns and records #N/A for SUM of the All/quantity column; SDK sums its numeric body and ignores the text header and blank total. Both producer differences remain preserved. Anonymous Excel Preview opened the immutable source but did not accept temporary UI input; it supplies no editing acceptance. Excel/WPS fill runtime unverified.",
              "references": REFERENCES, "source": SOURCE.relative_to(ROOT).as_posix(), "sourceSha256": sha256(data).hexdigest(), "cases": CASES,
              "office": version, "archiveSha256": base.axis.ARCHIVE_HASH, "producer": producer.relative_to(ROOT).as_posix(),
              "producerSha256": sha256(producer.read_bytes()).hexdigest(), "observed": base.observe(output)}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Pinned structured-fill observations changed."
    print("Verified seven pinned native structured fills and official XSD/OPC; producer differences retained.")


if __name__ == "__main__":
    main()
