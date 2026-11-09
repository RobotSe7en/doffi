"""Independent drawing moves/sorts with native normalization and copy differences."""
from hashlib import sha256
from pathlib import Path
import argparse
import importlib.util
import json
import os
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent


def module(name, file):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / file)
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value


base = module("drawing_transport_native", "check-workbook-axis-drawings.py")
REPORT = ROOT / "docs/evidence/workbook-drawing-transport.json"
SOURCES = [("images-shapes", "examples/xlsx/08-images-shapes.xlsx", "A1:I11", 0, 0, 10, 8),
           ("charts", "examples/xlsx/07-three-charts.xlsx", "A7:E19", 6, 0, 18, 4)]
CASES = ["cut-same", "cut-other", "sort"]


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
        # pyuno configures its own runtime environment; read the independent
        # CLI/font identity before importing it (native() isolates that state).
        import uno
        from com.sun.star.table.TableSortFieldType import AUTOMATIC
    directory = ROOT / ".cache/workbook-drawing-transport-native"
    directory.mkdir(exist_ok=True)
    records = []
    for name, path, area, first_row, first_col, last_row, last_col in SOURCES:
        source = ROOT / path
        cases = []
        for case in CASES:
            producer = ROOT / ("packages/office/tests/fixtures/workbook-drawing-transport-" + name + "-" + case + "-calc.xlsx")
            if args.record or args.verify_native:
                def mutate(document, prop):
                    sheet = document.getSheets().getByIndex(0)
                    if case == "sort":
                        start, column = (2, 1) if name == "images-shapes" else (6, 0)
                        for offset, value in enumerate([3, 1, 2]):
                            sheet.getCellByPosition(column, start + offset).setValue(value)
                        field = uno.createUnoStruct("com.sun.star.table.TableSortField")
                        field.Field, field.IsAscending, field.IsCaseSensitive, field.FieldType = 0, True, False, AUTOMATIC
                        field.CollatorLocale = uno.createUnoStruct("com.sun.star.lang.Locale", "en", "US", "")
                        sheet.getCellRangeByPosition(column, start, column + 1, start + 2).sort((
                            prop("SortFields", uno.Any("[]com.sun.star.table.TableSortField", (field,))), prop("IsSortColumns", False),
                            prop("ContainsHeader", False), prop("BindFormatsToContent", True)))
                        return
                    target = sheet
                    if case == "cut-other":
                        sheets = document.getSheets()
                        sheets.insertNewByName("Other", sheets.getCount())
                        target = sheets.getByName("Other")
                        # A genuine producer page-style choice; no output patch.
                        target.setPropertyValue("PageStyle", sheet.getPropertyValue("PageStyle"))
                    target.moveRange(target.getCellRangeByName("H20").getCellAddress(), sheet.getCellRangeByName(area).getRangeAddress())
                output = directory / (name + "-" + case + ".xlsx")
                base.axis.native(source, output, None, mutate)
                if args.record:
                    producer.write_bytes(output.read_bytes())
                observed = base.observe(output)
            else:
                observed = base.observe(producer)
            operation = {"kind": "sort", "from": [2, 1], "to": [4, 2], "rows": [4, 2, 3]} if case == "sort" and name == "images-shapes" else \
                {"kind": "sort", "from": [6, 0], "to": [8, 1], "rows": [8, 6, 7]} if case == "sort" else \
                {"kind": "cut", "from": [first_row, first_col], "to": [last_row, last_col], "target": [19, 7], "otherSheet": case == "cut-other"}
            cases.append({"name": case, "producerAction": "UNO sort" if case == "sort" else "UNO moveRange", "sdkOperation": operation,
                          "outputFile": str(producer.relative_to(ROOT)), "outputSha256": sha256(producer.read_bytes()).hexdigest(), "observed": observed})
        records.append({"name": name, "source": path, "sourceSha256": sha256(source.read_bytes()).hexdigest(), "original": base.observe(source), "cases": cases})
    if args.record or args.verify_native:
        python = sys.executable if importlib.util.find_spec("lxml") else str(ROOT / ".cache/xlsx-venv/bin/python")
        subprocess.run([python, str(ROOT / "scripts/check-xlsx-output.py"), str(directory)], check=True, stdout=subprocess.PIPE, text=True)
    report = {"scope": "Two immutable independent ZIP/XML sources and six actual native cell moves/sorts. SDK cut/sort semantics, producer normalization and producer copy behavior remain separate. Excel/WPS runtime unverified.",
              "office": base.axis.BUILD, "archiveSha256": base.axis.ARCHIVE_HASH,
              "fontManifest": {"source": "examples/xlsx/reference.json#environment.fonts", "sha256": sha256(json.dumps(fonts, sort_keys=True, separators=(",", ":")).encode()).hexdigest()},
              "references": base.REFERENCES[:2], "producerDifferences": [
                  "Calc normalizes one-cell/absolute containers, transformed frames and physical marker offsets; SDK explicit transport retains source identities, physical extents and displacement from the moving anchor cell.",
                  "The chart source's UNO moveRange retains the three original charts and adds a fourth chart. This is an observed producer copy, not proof of SDK cut semantics or Excel behavior.",
                  "Calc row sorting moves the first chart to A9 but may retain a transformed image at its normalized original row. Both native observations remain recorded separately from original source markers.",
                  "Native output can change drawing array order. SDK cell transport retains source object/layer order and object, paragraph, chart and series identities."], "sources": records}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Independent drawing transport evidence changed."
    print("Verified six drawing move/sort records" + (" with pinned Calc, fixed fonts and official XSD/OPC." if args.record or args.verify_native else "."))


if __name__ == "__main__":
    main()
