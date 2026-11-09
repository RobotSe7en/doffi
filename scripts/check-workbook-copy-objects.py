"""Independent cell/object copies, complete chart vectors and partial-vector retention."""
from hashlib import sha256
from pathlib import Path
import argparse
import importlib.util
import json
import os
import re
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("object_copy_native", ROOT / "scripts/check-workbook-axis-drawings.py")
base = importlib.util.module_from_spec(spec)
spec.loader.exec_module(base)
REPORT = ROOT / "docs/evidence/workbook-copy-objects.json"
SOURCES = [("charts", "examples/xlsx/07-three-charts.xlsx", [
    ("full-same", "A1:O40", False), ("full-other", "A1:O40", True), ("chart-only", "A7:E19", True),
    ("partial-vector", "A3:E19", True), ("single-full", "A1:E19", True)]),
    ("images-shapes", "examples/xlsx/08-images-shapes.xlsx", [
        ("full-same", "A1:O40", False), ("full-other", "A1:O40", True), ("partial-box", "A1:I11", True)])]


def coordinate(value):
    match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", value)
    column = 0
    for letter in match[1]:
        column = column * 26 + ord(letter) - ord("A") + 1
    return {"row": int(match[2]) - 1, "column": column - 1}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    fonts = json.loads((ROOT / "examples/xlsx/reference.json").read_text())["environment"]["fonts"]
    if args.record or args.verify_native:
        assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == base.axis.ARCHIVE_HASH
        os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
        os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
        environment = base.shared.environment()
        assert environment["libreOffice"] == base.axis.BUILD and environment["fonts"] == fonts, "Pinned Office/fonts changed."
    directory = ROOT / ".cache/workbook-copy-objects-native"
    directory.mkdir(exist_ok=True)
    sources = []
    for name, path, definitions in SOURCES:
        source = ROOT / path
        cases = []
        for case, area, other in definitions:
            producer = ROOT / ("packages/office/tests/fixtures/workbook-copy-objects-" + name + "-" + case + "-calc.xlsx")
            if args.record or args.verify_native:
                def mutate(document, prop):
                    sheet = document.getSheets().getByIndex(0)
                    target = sheet
                    if other:
                        sheets = document.getSheets()
                        sheets.insertNewByName("Other", sheets.getCount())
                        target = sheets.getByName("Other")
                        target.setPropertyValue("PageStyle", sheet.getPropertyValue("PageStyle"))
                    target.copyRange(target.getCellRangeByName("H60").getCellAddress(), sheet.getCellRangeByName(area).getRangeAddress())
                output = directory / (name + "-" + case + ".xlsx")
                base.axis.native(source, output, None, mutate)
                if args.record:
                    producer.write_bytes(output.read_bytes())
                observed = base.observe(output)
            else:
                observed = base.observe(producer)
            first, last = area.split(":")
            cases.append({"name": case, "producerAction": "UNO copyRange", "sourceRange": {"from": coordinate(first), "to": coordinate(last)}, "target": coordinate("H60"), "otherSheet": other,
                          "outputFile": str(producer.relative_to(ROOT)), "outputSha256": sha256(producer.read_bytes()).hexdigest(), "observed": observed})
        sources.append({"name": name, "source": path, "sourceSha256": sha256(source.read_bytes()).hexdigest(), "original": base.observe(source), "cases": cases})
    if args.record or args.verify_native:
        python = sys.executable if importlib.util.find_spec("lxml") else str(ROOT / ".cache/xlsx-venv/bin/python")
        subprocess.run([python, str(ROOT / "scripts/check-xlsx-output.py"), str(directory)], check=True, stdout=subprocess.PIPE, text=True)
    report = {"scope": "Two immutable independent ZIP/XML sources and eight actual native range copies, with pictures, shapes, absolute boxes and three basic charts. Complete and partial direct input vectors are observed separately. Excel/WPS runtime unverified.",
              "office": base.axis.BUILD, "archiveSha256": base.axis.ARCHIVE_HASH,
              "fontManifest": {"source": "examples/xlsx/reference.json#environment.fonts", "sha256": sha256(json.dumps(fonts, sort_keys=True, separators=(",", ":")).encode()).hexdigest()},
              "references": ["https://learn.microsoft.com/en-us/office/vba/api/excel.range.copy", "https://support.microsoft.com/en-au/excel/paste-options?nochrome=true"],
              "producerDifferences": [
                  "Calc redirects a complete copied direct chart vector, but keeps the whole original vector when only some or none of its cells are copied; these are observations, not an Excel runtime claim.",
                  "Calc normalizes native anchor containers, transformed frames, marker offsets and drawing names. SDK copies retain source kinds/flags, measured extents and inert rich source with fresh identities.",
                  "SDK keeps symbolic defined-name and structured-reference bindings rather than guessing rewritten definitions; this probe covers direct native vectors, not producer verification of every symbolic-reference copy.",
                  "Absolute SDK boxes require complete measured footprint coverage; cell-anchored objects follow their logical start marker. Original source markers and the producer's normalized geometry remain separate."], "sources": sources}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Native object copy evidence changed."
    print("Verified eight independent object/chart copies" + (" with pinned Calc, fixed fonts and official XSD/OPC." if args.record or args.verify_native else "."))


if __name__ == "__main__":
    main()
