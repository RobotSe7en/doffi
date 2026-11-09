"""Independent partial native table cuts and actual producer reopen differences."""
from workbook_acceptance import audit_schema
from hashlib import sha256
from pathlib import Path
import argparse
import importlib.util
import json
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("partial_table_move", ROOT / "scripts/check-workbook-move-definitions.py")
base = importlib.util.module_from_spec(spec)
spec.loader.exec_module(base)
SOURCE = ROOT / "packages/office/tests/fixtures/workbook-axis-tables.xlsx"
REPORT = ROOT / "docs/evidence/workbook-partial-table-move.json"
CASES = [
    {"name": "body-same", "source": "A2:D3", "target": "H260", "sheet": 0},
    {"name": "body-other", "source": "A2:D3", "target": "H260", "sheet": 1},
    {"name": "column-other", "source": "D2:D3", "target": "H260", "sheet": 1},
    {"name": "inside-body", "source": "B2:C3", "target": "B4", "sheet": 0},
    {"name": "total-other", "source": "D6", "target": "H260", "sheet": 1},
]
REFERENCES = ["https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables",
              "https://support.microsoft.com/en-gb/excel/move-or-copy-cells-rows-and-columns"]


def observe(path):
    result = base.observe(path)
    # Retain calculated/totals recipes and native filter/header details as well
    # as cell caches; a producer loss must not disappear behind value equality.
    result["nativeDefinitions"] = base.source.observe(path)["tables"]
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    data = SOURCE.read_bytes()
    assert data == base.source.generate(), "The independent table source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(data).hexdigest() and report["references"] == REFERENCES and report["sourceObserved"] == observe(SOURCE)
        assert [item["case"] for item in report["cases"]] == CASES
        for item in report["cases"]:
            for step in ["moved", "reopened"]:
                producer = ROOT / item[step]["file"]
                assert sha256(producer.read_bytes()).hexdigest() == item[step]["sha256"] and observe(producer) == item[step]["observed"]
        print("Verified five independent partial-table cuts and five actual reopen records.")
        return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == base.axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-partial-table-move-native"
    directory.mkdir(exist_ok=True)
    input_path = directory / "source.xlsx"
    input_path.write_bytes(data)
    cases = []
    for case in CASES:
        def mutate(document, prop):
            document.calculateAll()
            sheet, target = document.getSheets().getByIndex(0), document.getSheets().getByIndex(case["sheet"])
            target.moveRange(target.getCellRangeByName(case["target"]).getCellAddress(), sheet.getCellRangeByName(case["source"]).getRangeAddress())
        item = {"case": case}
        for step in ["moved", "reopened"]:
            output = directory / (case["name"] + "-" + step + ".xlsx")
            origin = input_path if step == "moved" else directory / (case["name"] + "-moved.xlsx")
            version = base.axis.native(origin, output, None, mutate if step == "moved" else lambda document, prop: None)
            producer = SOURCE.with_name("workbook-partial-table-move-" + case["name"] + "-" + step + "-calc.xlsx")
            if args.record:
                producer.write_bytes(output.read_bytes())
            item[step] = {"file": producer.relative_to(ROOT).as_posix(), "sha256": sha256(producer.read_bytes()).hexdigest(), "observed": observe(output)}
        cases.append(item)
    audit_schema(directory, stdout=subprocess.PIPE, text=True)
    report = {"scope": "Immutable independent ZIP/XML table source; five actual pinned Calc partial moveRange edits and five real reopens, with official XSD/OPC. Definitions remain on the original table; producer loss of calculated recipes and native filter/header differences remain recorded. Calc's moved current-row structured spellings can retain old caches initially and change after reopen; both versions remain original bytes. SDK applies the user's exact cut-reference mapping to current-row inputs of moved formulas, retaining unmoved inputs and table definitions. Dynamic structured consumers remaining inside the table continue to use its current data. Excel/WPS structural runtime unverified.",
              "references": REFERENCES, "source": SOURCE.relative_to(ROOT).as_posix(), "sourceSha256": sha256(data).hexdigest(), "sourceObserved": observe(SOURCE),
              "office": version, "archiveSha256": base.axis.ARCHIVE_HASH, "cases": cases}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Pinned partial-table move/reopen observations changed."
    print("Verified five actual pinned partial-table cuts, five actual reopens and official XSD/OPC.")


if __name__ == "__main__":
    main()
