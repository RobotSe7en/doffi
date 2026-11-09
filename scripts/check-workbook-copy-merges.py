"""Independent merged copies and actual autofill in the pinned Calc producer."""
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
spec = importlib.util.spec_from_file_location("copy_merge_axis", ROOT / "scripts/check-workbook-axis.py")
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-copy-merges.xlsx"
REPORT = ROOT / "docs/evidence/workbook-copy-merges.json"
CASES = ["copy-empty", "copy-matching", "copy-other", "fill-down", "fill-up", "copy-incompatible"]


def generate():
    book = axis.source.Book()
    sheet = book.sheet("Data")
    sheet.add("A1", 7)
    sheet.add("A256", kind="inlineStr", rich='<r><rPr><b/><rFont val="Noto Sans CJK SC"/></rPr><t>合并</t></r><r><rPr><i/><rFont val="Liberation Sans"/></rPr><t>😀禰󠄀</t></r>', style=9)
    sheet.add("A258", 999, formula="A1+$A$1")
    sheet.after = '<mergeCells count="3"><mergeCell ref="A256:B257"/><mergeCell ref="D256:E257"/><mergeCell ref="H256:H257"/></mergeCells>'
    sheet.print_area = "$A$1:$L$270"
    book.sheet("Other").print_area = "$A$1:$L$270"
    return axis.source.zipped(book.finish())


def observe(path):
    with ZipFile(path) as archive:
        strings = ["".join(node.itertext()) for node in ET.fromstring(archive.read("xl/sharedStrings.xml"))] if "xl/sharedStrings.xml" in archive.namelist() else []
        result = []
        for index in [1, 2]:
            sheet = ET.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml"))
            cells = []
            for cell in sheet.findall("{*}sheetData/{*}row/{*}c"):
                kind, value, formula = cell.get("t", "n"), cell.findtext("{*}v"), cell.findtext("{*}f")
                if value is None and formula is None:
                    continue
                cells.append({"address": cell.get("r"), "type": "text" if kind == "s" else kind,
                              "value": strings[int(value)] if kind == "s" else value, "formula": formula})
            result.append({"merges": [node.get("ref") for node in sheet.findall("{*}mergeCells/{*}mergeCell")], "cells": cells})
        return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    data = generate()
    if not args.record:
        assert FIXTURE.read_bytes() == data, "Independent merged source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(data).hexdigest()
        assert [case["name"] for case in report["cases"]] == CASES
        for case in report["cases"]:
            producer = FIXTURE.with_name("workbook-copy-merges-" + case["name"] + "-calc.xlsx")
            assert sha256(producer.read_bytes()).hexdigest() == case["producerSha256"]
            assert observe(producer) == case["observed"]
        print("Verified six independent merged-copy/autofill source and producer cases.")
        return
    import uno
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-copy-merges-native"
    directory.mkdir(exist_ok=True)
    input_path = directory / "source.xlsx"
    input_path.write_bytes(data)
    cases = []
    for name in CASES:
        def mutate(document, prop):
            source = document.getSheets().getByIndex(0)
            if name.startswith("fill-"):
                address, direction = ("A256:B261", "TO_BOTTOM") if name == "fill-down" else ("A252:B257", "TO_TOP")
                source.getCellRangeByName(address).fillAuto(uno.Enum("com.sun.star.sheet.FillDirection", direction), 2)
            else:
                destination = document.getSheets().getByIndex(1 if name == "copy-other" else 0)
                address = "D256" if name == "copy-matching" else "H256" if name == "copy-incompatible" else "D1"
                destination.copyRange(destination.getCellRangeByName(address).getCellAddress(), source.getCellRangeByName("A256:B258").getRangeAddress())
        output = directory / (name + ".xlsx")
        version = axis.native(input_path, output, {}, mutate)
        producer = FIXTURE.with_name("workbook-copy-merges-" + name + "-calc.xlsx")
        if args.record:
            producer.write_bytes(output.read_bytes())
        # ZIP timestamps and producer UUIDs may vary; compare semantic records.
        cases.append({"name": name, "producerSha256": sha256(producer.read_bytes()).hexdigest(), "observed": observe(output)})
    audit_schema(directory, stdout=subprocess.PIPE, text=True)
    report = {"scope": "Independent ZIP/XML source, actual pinned Calc copyRange and fillAuto. Mismatched destination merges are rejected by the SDK; Calc retains its old merge. Excel/WPS runtime unverified.",
              "office": version, "archiveSha256": axis.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(), "cases": cases}
    if args.record:
        FIXTURE.write_bytes(data)
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Pinned merged copy/autofill observations changed."
    print("Verified six actual pinned Calc merged-copy/autofill operations and official XSD/OPC.")


if __name__ == "__main__":
    main()
