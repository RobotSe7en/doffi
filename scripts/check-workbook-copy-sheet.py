"""Independent sources and pinned Calc whole-worksheet copy observations."""
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
spec = importlib.util.spec_from_file_location("copy_sheet_axis", ROOT / "scripts/check-workbook-axis.py")
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)
REPORT = ROOT / "docs/evidence/workbook-copy-sheet.json"
FIXTURES = ROOT / "packages/office/tests/fixtures"
CASES = [("names", "packages/office/tests/fixtures/workbook-name-prefix.xlsx"),
         ("tables", "packages/office/tests/fixtures/workbook-axis-tables.xlsx"),
         *[(name, f"examples/xlsx/{file}.xlsx") for name, file in [
             ("structure", "03-structure"), ("charts", "07-three-charts"), ("pictures", "08-images-shapes"),
             ("classic", "09-classic-notes"), ("modern", "10-modern-threads"),
             ("print", "11-print-layout"), ("filtered", "12-multi-sheet-filter")]],
         ("hyperlinks", "packages/office/tests/fixtures/workbook-copy-sheet-hyperlinks.xlsx")]
REFERENCES = ["https://support.microsoft.com/en-us/excel/get-started/move-or-copy-worksheets-or-worksheet-data",
              "https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a469e5f5-a102-49bd-9642-8a8e8aaf1623",
              "https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.subaddress"]


def hyperlink_source():
    book = axis.source.Book()
    sheet = book.sheet("Data")
    for address, value in [("A1", "Internal"), ("A2", "Here"), ("A3", "Named"), ("A4", "Other"), ("A5", 5), ("A6", "External")]:
        sheet.add(address, value)
    sheet.after = ('<hyperlinks><hyperlink ref="A1" location="Data!$A$5"/><hyperlink ref="A2" location="A5"/>'
                   '<hyperlink ref="A3" location="Rate"/><hyperlink ref="A4" location="Other!A1"/>'
                   '<hyperlink ref="A6" r:id="external"/></hyperlinks>')
    book.relate(sheet.path, "external", axis.source.R + "/hyperlink", "https://example.com/authorized-link", True)
    book.names.append(({"name": "Rate", "localSheetId": 0}, "Data!$A$5"))
    book.sheet("Other").add("A1", 7)
    return axis.source.zipped(book.finish())


def observe(path):
    """Exclude producer GUIDs/timestamps, retaining data, references and geometry."""
    with ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        result = {"sheets": [{"name": node.get("name"), "state": node.get("state", "visible")}
                             for node in workbook.findall("{*}sheets/{*}sheet")],
                  "names": [{"name": node.get("name"), "scope": node.get("localSheetId"), "expression": node.text}
                            for node in workbook.findall("{*}definedNames/{*}definedName")], "parts": {}}
        for part in sorted(archive.namelist()):
            if not part.endswith(".xml"):
                continue
            if part.startswith("xl/worksheets/sheet"):
                sheet = ET.fromstring(archive.read(part))
                result["parts"][part] = {
                    "cells": [{"address": cell.get("r"), "formula": cell.findtext("{*}f"), "value": cell.findtext("{*}v"),
                               "type": cell.get("t", "n"), "text": "".join(cell.find("{*}is").itertext()) if cell.find("{*}is") is not None else None}
                              for cell in sheet.findall("{*}sheetData/{*}row/{*}c")],
                    "merges": [node.get("ref") for node in sheet.findall("{*}mergeCells/{*}mergeCell")],
                    "panes": [dict(node.attrib) for node in sheet.findall("{*}sheetViews/{*}sheetView/{*}pane")],
                    "rows": [{key: value for key, value in node.attrib.items() if key not in {"spans", "dyDescent"}}
                             for node in sheet.findall("{*}sheetData/{*}row")],
                    "columns": [dict(node.attrib) for node in sheet.findall("{*}cols/{*}col")],
                    "filters": [dict(node.attrib) for node in sheet.findall("{*}autoFilter")],
                    "links": [{key: value for key, value in node.attrib.items() if not key.endswith("}id")}
                              for node in sheet.findall("{*}hyperlinks/{*}hyperlink")],
                    "print": {tag: [dict(node.attrib) for node in sheet.findall("{*}" + tag)]
                              for tag in ["printOptions", "pageMargins", "pageSetup"]}}
            elif part.startswith("xl/tables/"):
                table = ET.fromstring(archive.read(part))
                result["parts"][part] = {"name": table.get("name"), "displayName": table.get("displayName"), "range": table.get("ref"),
                                         "columns": [{"name": node.get("name"), "calculated": node.findtext("{*}calculatedColumnFormula"), "total": node.findtext("{*}totalsRowFormula")}
                                                     for node in table.findall("{*}tableColumns/{*}tableColumn")]}
            elif part.startswith("xl/charts/chart"):
                chart = ET.fromstring(archive.read(part))
                result["parts"][part] = {"references": [node.text for node in chart.findall(".//{*}f")],
                                         "vectors": [[node.text for node in cache.findall("{*}pt/{*}v")]
                                                     for cache in chart.iter() if cache.tag.rsplit("}", 1)[-1] in {"numCache", "strCache", "numLit", "strLit"}]}
            elif part.startswith("xl/comments"):
                comments = ET.fromstring(archive.read(part))
                result["parts"][part] = {"authors": [node.text for node in comments.findall("{*}authors/{*}author")],
                                         "comments": [{"address": node.get("ref"), "author": node.get("authorId"), "text": "".join(node.find("{*}text").itertext())}
                                                      for node in comments.findall("{*}commentList/{*}comment")]}
        result["counts"] = {kind: sum(part.startswith(prefix) and part.endswith(".xml") for part in archive.namelist())
                            for kind, prefix in [("charts", "xl/charts/chart"), ("drawings", "xl/drawings/drawing"),
                                                 ("classic", "xl/comments"), ("modern", "xl/threadedComments/")]}
        result["mediaSha256"] = sorted(sha256(archive.read(part)).hexdigest() for part in archive.namelist() if part.startswith("xl/media/"))
        return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    generated = hyperlink_source()
    hyperlink_fixture = ROOT / CASES[-1][1]
    if args.record:
        hyperlink_fixture.write_bytes(generated)
    else:
        assert hyperlink_fixture.read_bytes() == generated, "Independent hyperlink copy source changed."
    recorded = json.loads(REPORT.read_text()) if not args.record else None
    if args.record or args.verify_native:
        assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-copy-sheet-native"
    directory.mkdir(exist_ok=True)
    cases = []
    for name, source in CASES:
        producer = FIXTURES / ("workbook-copy-sheet-" + name + "-calc.xlsx")
        if args.record or args.verify_native:
            def mutate(document, prop):
                sheets = document.getSheets()
                sheets.copyByName(sheets.getByIndex(0).Name, "Copied", len(sheets.getElementNames()))
                copied = sheets.getByName("Copied")
                if name == "names":
                    for address, value in [("A5", 99), ("A6", 123), ("H1", 88)]:
                        copied.getCellRangeByName(address).Value = value
                elif name == "tables":
                    copied.getCellRangeByName("B2").Value = 20
            output = directory / (name + ".xlsx")
            axis.native(ROOT / source, output, {}, mutate)
            if args.record:
                producer.write_bytes(output.read_bytes())
            observed = observe(output)
        else:
            observed = observe(producer)
        cases.append({"name": name, "source": source, "sourceSha256": sha256((ROOT / source).read_bytes()).hexdigest(),
                      "producerSha256": sha256(producer.read_bytes()).hexdigest(), "observed": observed})
    if args.record or args.verify_native:
        audit_python = sys.executable if importlib.util.find_spec("lxml") is not None else str(ROOT / ".cache/xlsx-venv/bin/python")
        subprocess.run([audit_python, str(ROOT / "scripts/check-xlsx-output.py"), str(directory)], check=True, stdout=subprocess.PIPE, text=True)
    report = {"scope": "Ten independent native inputs, actual whole-sheet copies in pinned Calc, and stable data/name/table/chart/comment/print/hyperlink observations. Static hyperlink destinations retain their original text rather than being rewritten as formulas. Excel/WPS runtime unverified.",
              "office": axis.BUILD, "archiveSha256": axis.ARCHIVE_HASH, "references": REFERENCES,
              "producerDifferences": ["Calc clones source-bound global names as copied-sheet local definitions and removes the explicit current-workbook name prefix.",
                                      "Calc creates a new table name but copied cell formulas retain the original table reference; editing copied B2 to 20 leaves D2 at 20. SDK references the copied table and calculates 200.",
                                      "Calc drops modern threads on this source; SDK preserves complete copied threads with fresh identities."], "cases": cases}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert recorded == report, "Pinned whole-sheet copy source or producer observations changed."
    print("Verified ten whole-sheet native copy cases" + (" with actual pinned Calc and official XSD/OPC." if args.record or args.verify_native else "."))


if __name__ == "__main__":
    main()
