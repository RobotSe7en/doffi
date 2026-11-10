"""Whole native table copy provenance and actual producer reopen differences."""
from workbook_acceptance import audit_schema
from hashlib import sha256
from pathlib import Path
import argparse
import importlib.util
import json
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("copy_tables_source", ROOT / "scripts/check-workbook-move-definitions.py")
base = importlib.util.module_from_spec(spec)
spec.loader.exec_module(base)
SOURCE = ROOT / "packages/office/tests/fixtures/workbook-axis-tables.xlsx"
REPORT = ROOT / "docs/evidence/workbook-copy-tables.json"
CASES = ["same-sheet", "other-sheet"]


def observe(path):
    value = base.observe(path)
    value["nativeDefinitions"] = base.source.observe(path)["tables"]
    return value


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    data = SOURCE.read_bytes()
    assert data == base.source.generate(), "Independent source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(data).hexdigest() and [item["name"] for item in report["cases"]] == CASES
        for item in report["cases"]:
            for step in ["copied", "reopened"]:
                producer = ROOT / item[step]["file"]
                assert sha256(producer.read_bytes()).hexdigest() == item[step]["sha256"] and observe(producer) == item[step]["observed"]
        print("Verified two independent whole-table copies and two actual producer reopens.")
        return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == base.axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-copy-tables-native"
    directory.mkdir(exist_ok=True)
    cases = []
    for name in CASES:
        def mutate(document, prop):
            document.calculateAll()
            source = document.getSheets().getByIndex(0)
            target = document.getSheets().getByIndex(1 if name == "other-sheet" else 0)
            target.copyRange(target.getCellRangeByName("H260").getCellAddress(), source.getCellRangeByName("A1:D6").getRangeAddress())
            document.calculateAll()
        item = {"name": name}
        for step in ["copied", "reopened"]:
            output = directory / (name + "-" + step + ".xlsx")
            origin = SOURCE if step == "copied" else directory / (name + "-copied.xlsx")
            version = base.axis.native(origin, output, None, mutate if step == "copied" else lambda document, prop: None)
            producer = SOURCE.with_name("workbook-copy-tables-" + name + "-" + step + "-calc.xlsx")
            if args.record:
                producer.write_bytes(output.read_bytes())
            item[step] = {"file": producer.relative_to(ROOT).as_posix(), "sha256": sha256(producer.read_bytes()).hexdigest(), "observed": observe(output)}
        cases.append(item)
    audit_schema(directory, stdout=subprocess.PIPE, text=True)
    report = {"scope": "Immutable independent ZIP/XML native table source, two actual pinned Calc whole-range copyRange edits and two actual reopens with official XSD/OPC. Producer copies cells without a new table definition, retains original structured references and yields VALUE errors for copied current-row formulas outside the original table. SDK complete contents copies create independent table/column identities and map self references to the new table; partial and selective copies keep ordinary bindings. Table recipes remain metadata, actual copied exceptions and rich headers remain source cells. Producer loss of original calculated recipes/filter differences remains explicit. Excel/WPS whole-table copy runtime unverified.",
              "source": SOURCE.relative_to(ROOT).as_posix(), "sourceSha256": sha256(data).hexdigest(), "sourceObserved": observe(SOURCE),
              "office": version, "archiveSha256": base.axis.ARCHIVE_HASH, "cases": cases}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Pinned table copy/reopen observations changed."
    print("Verified two pinned whole-table copies, two actual reopens and official XSD/OPC.")


if __name__ == "__main__":
    main()
