"""Independent ZIP/XML inputs for an explicitly recorded Excel web probe.

These synthetic inputs are not acceptance evidence until the producer and any
repair are inspected. Original reading fixtures remain unchanged.
"""
import argparse
import hashlib
import importlib.util
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DESTINATION = ROOT / "docs/evidence/excel-web-inputs"
spec = importlib.util.spec_from_file_location("excel_web_probe_inputs", ROOT / "scripts/generate-xlsx-fixtures.py")
fixture = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fixture)


def sources():
    book = fixture.Book()
    sheet = book.sheet("Data")
    sheet.add("A1", 2)
    sheet.add("C1", 999, formula="A1+1")
    sheet.add("C2", 999, formula="C1*2")
    yield "recalculation.xlsx", fixture.zipped(book.finish()), [{"sheet": "Data", "cell": "C1", "formula": "A1+1", "cache": 999}, {"sheet": "Data", "cell": "C2", "formula": "C1*2", "cache": 999}]

    book = fixture.Book()
    definitions = [
        ({"name": "ProbeFixed"}, "Data!$A$1"),
        ({"name": "ProbeCaller"}, "!$A$1"),
        ({"name": "ProbeRel"}, "Data!$A1"),
        ({"name": "ProbeRelCaller"}, "!$A1"),
        ({"name": "ProbeRelColumn"}, "Data!B$1"),
        ({"name": "ProbeShadow"}, "30"),
        ({"name": "ProbeShadow", "localSheetId": 0}, "10"),
        ({"name": "ProbeShadow", "localSheetId": 1}, "20"),
        ({"name": "ProbePlainAlias"}, "ProbeShadow"),
        ({"name": "ProbeBangAlias"}, "!ProbeShadow"),
        ({"name": "ProbeGlobal"}, "Data!$A$1+1000"),
        ({"name": "ProbeBangGlobal"}, "!ProbeGlobal"),
        ({"name": "ProbeQualGlobal"}, "Data!ProbeGlobal"),
        ({"name": "ProbeBookGlobal"}, "[0]!ProbeGlobal"),
        ({"name": "ProbeLocalPlain", "localSheetId": 0}, "ProbeShadow"),
        ({"name": "ProbeLocalBang", "localSheetId": 0}, "!ProbeShadow"),
        ({"name": "ProbeNested"}, "ProbePlainAlias"),
        ({"name": "ProbeNestedBang"}, "!ProbeBangAlias"),
    ]
    formulas = ["ProbeFixed", "ProbeCaller", "ProbeRel", "ProbeRelCaller", "ProbeRelColumn", "ProbeShadow", "ProbePlainAlias", "ProbeBangAlias",
                "ProbeBangGlobal", "ProbeQualGlobal", "ProbeBookGlobal", "Data!ProbeLocalPlain", "Data!ProbeLocalBang", "ProbeNested", "ProbeNestedBang"]
    cells = []
    for name, base in [("Data", 10), ("Other", 100), ("Third", 1000)]:
        sheet = book.sheet(name)
        sheet.before += '<cols><col min="2" max="2" width="34" customWidth="1"/></cols>'
        for row in range(1, 31):
            sheet.add(f"A{row}", base + row)
        sheet.add("B1", base + 200)
        sheet.add("D1", base + 400)
        for row, formula in enumerate(formulas, 5):
            sheet.add(f"B{row}", formula)
            sheet.add(f"C{row}", 999, formula=formula)
            cells.append({"sheet": name, "cell": f"C{row}", "formula": formula, "cache": 999})
    book.names = definitions
    yield "name-semantics.xlsx", fixture.zipped(book.finish()), cells

    book = fixture.Book()
    book.names = [entry for entry in definitions if entry[0]["name"] in {"ProbeCaller", "ProbeRelCaller", "ProbeShadow", "ProbeBangAlias", "ProbeGlobal", "ProbePlainAlias"}]
    book.names += [({"name": "ProbeLocalCaller", "localSheetId": 0}, "!$A$1"), ({"name": "ProbeLocalRelCaller", "localSheetId": 0}, "!$A1")]
    formulas = ["Data!ProbeCaller", "Other!ProbeCaller", "Data!ProbeRelCaller", "Other!ProbeRelCaller", "Data!ProbeBangAlias", "Other!ProbeBangAlias",
                "[0]!ProbeCaller", "Data!ProbeLocalCaller", "Data!ProbeLocalRelCaller", "[0]!ProbePlainAlias"]
    cells = []
    for name, base in [("Data", 10), ("Other", 100)]:
        sheet = book.sheet(name)
        sheet.before += '<cols><col min="2" max="2" width="34" customWidth="1"/></cols>'
        for row in range(1, 31):
            sheet.add(f"A{row}", base + row)
        for row, formula in enumerate(formulas, 5):
            sheet.add(f"B{row}", formula)
            sheet.add(f"C{row}", 999, formula=formula)
            cells.append({"sheet": name, "cell": f"C{row}", "formula": formula, "cache": 999})
    yield "qualified-context.xlsx", fixture.zipped(book.finish()), cells

    book = fixture.Book()
    book.names = [({"name": "ProbeNegRow"}, "Data!$A1048576"), ({"name": "ProbeNegColumn"}, "Data!XFD$1"),
                  ({"name": "ProbeNegBoth"}, "!XFD1048576"), ({"name": "ProbeNegRange"}, "Data!$A1048575:$A1048576"),
                  ({"name": "ProbeNegColumns"}, "Data!XFC$1:XFD$1"), ({"name": "ProbeNegWholeRow"}, "Data!1048576:1048576")]
    formulas = [("C5", "ProbeNegRow"), ("C6", "ProbeNegColumn"), ("C7", "ProbeNegBoth"), ("C8", "SUM(ProbeNegRange)"),
                ("C9", "ROW(ProbeNegWholeRow)"), ("C10", "ROWS(ProbeNegWholeRow)"), ("C11", "COLUMNS(ProbeNegWholeRow)"),
                ("G1", "SUM(ProbeNegRange)"), ("G2", "SUM(ProbeNegRange)"), ("G3", "SUM(ProbeNegRange)"), ("G4", "SUM(ProbeNegRange)"),
                ("B2", "SUM(ProbeNegColumns)"), ("C2", "SUM(ProbeNegColumns)"), ("A31", "SUM(ProbeNegColumns)")]
    cells = []
    for name, base in [("Data", 10), ("Other", 100)]:
        sheet = book.sheet(name)
        sheet.before += '<cols><col min="2" max="2" width="34" customWidth="1"/></cols>'
        for row in range(1, 31):
            sheet.add(f"A{row}", base + row)
        for cell, value in [("B1", base + 200), ("XFC1", 701), ("XFD1", 702), ("A1048575", 777), ("A1048576", 888)]:
            sheet.add(cell, value)
        for cell, formula in formulas:
            if cell.startswith("C") and int(cell[1:]) >= 5:
                sheet.add("B" + cell[1:], formula)
            sheet.add(cell, 999, formula=formula)
            cells.append({"sheet": name, "cell": cell, "formula": formula, "cache": 999})
    yield "relative-bounds.xlsx", fixture.zipped(book.finish()), cells


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    files = {}
    for name, data, cells in sources():
        files[DESTINATION / name] = data
        files[DESTINATION / (name + ".json")] = (json.dumps({"status": "unverified-input", "producer": "independent Python ZIP/XML, without SDK conversion",
            "sha256": hashlib.sha256(data).hexdigest(), "cells": cells}, ensure_ascii=False, indent=2) + "\n").encode()
    for path, data in files.items():
        if args.check:
            if not path.exists() or path.read_bytes() != data:
                raise SystemExit("Excel web probe input differs: " + str(path.relative_to(ROOT)))
        else:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(data)
    print("Verified independent Excel web probe bytes" if args.check else "Wrote independent Excel web probe inputs; producer results remain unverified")


if __name__ == "__main__":
    main()
