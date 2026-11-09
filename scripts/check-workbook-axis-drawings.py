"""Pinned native axis edits for drawings and classic note boxes, retaining producer differences."""
from workbook_acceptance import audit_schema
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import importlib.util
import json
import os
import subprocess
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent


def module(name, file):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / file)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


axis = module("drawing_axis_native", "check-workbook-axis.py")
shared = module("drawing_axis_fonts", "pptx-reference.py")
FIXTURES = ROOT / "packages/office/tests/fixtures"
REPORT = ROOT / "docs/evidence/workbook-axis-drawings.json"
SOURCES = ["07-three-charts", "08-images-shapes", "09-classic-notes"]
CASES = [("resave", None),
         ("insert-rows-before", {"axis": "row", "action": "insert", "index": 1, "count": 2}),
         ("insert-rows-inside", {"axis": "row", "action": "insert", "index": 3, "count": 2}),
         ("delete-row-before", {"axis": "row", "action": "delete", "index": 0, "count": 1}),
         ("delete-row-inside", {"axis": "row", "action": "delete", "index": 3, "count": 1}),
         ("insert-cols-before", {"axis": "column", "action": "insert", "index": 1, "count": 2}),
         ("insert-col-inside", {"axis": "column", "action": "insert", "index": 4, "count": 1}),
         ("delete-col-before", {"axis": "column", "action": "delete", "index": 0, "count": 1}),
         ("delete-col-inside", {"axis": "column", "action": "delete", "index": 4, "count": 1})]
REFERENCES = ["https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.onecellanchor?view=openxml-3.0.1",
              "https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.editasvalues?view=openxml-3.0.1",
              "https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.spreadsheet.resizewithcells?view=openxml-3.0.1"]


def observe(path):
    """Record source geometry and flags; exclude GUIDs and timestamps."""
    result = {"drawings": [], "notes": [], "charts": [], "comments": [], "sheets": []}
    with ZipFile(path) as archive:
        for part in sorted(archive.namelist()):
            if part.startswith("xl/drawings/drawing") and part.endswith(".xml"):
                for anchor in ET.fromstring(archive.read(part)):
                    identity = anchor.find(".//{*}cNvPr")
                    value = {"name": identity.get("name") if identity is not None else None, "kind": anchor.tag.rsplit("}", 1)[-1], "editAs": anchor.get("editAs")}
                    for key in ["from", "to", "pos", "ext"]:
                        node = anchor.find("{*}" + key)
                        if node is not None:
                            value[key] = dict(node.attrib) if node.attrib else {child.tag.rsplit("}", 1)[-1]: child.text for child in node}
                    frame = anchor.find(".//{*}xfrm")
                    if frame is not None:
                        value["frame"] = {"attributes": dict(frame.attrib), "geometry": {child.tag.rsplit("}", 1)[-1]: dict(child.attrib) for child in frame}}
                    result["drawings"].append(value)
            elif part.endswith(".vml"):
                for node in ET.fromstring(archive.read(part)).findall(".//{urn:schemas-microsoft-com:office:excel}ClientData"):
                    result["notes"].append({child.tag.rsplit("}", 1)[-1]: child.text for child in node})
            elif part.startswith("xl/charts/chart") and part.endswith(".xml"):
                result["charts"].append([node.text for node in ET.fromstring(archive.read(part)).findall(".//{*}f")])
            elif part.startswith("xl/comments") and part.endswith(".xml"):
                result["comments"].extend({"cell": node.get("ref"), "text": "".join(node.find("{*}text").itertext())}
                                          for node in ET.fromstring(archive.read(part)).findall("{*}commentList/{*}comment"))
            elif part.startswith("xl/worksheets/sheet") and part.endswith(".xml"):
                sheet = ET.fromstring(archive.read(part))
                result["sheets"].append({"defaults": [dict(node.attrib) for node in sheet.findall("{*}sheetFormatPr")],
                                          "columns": [dict(node.attrib) for node in sheet.findall("{*}cols/{*}col")],
                                          "rows": [{key: value for key, value in node.attrib.items() if key in {"r", "ht", "customHeight", "hidden"}}
                                                   for node in sheet.findall("{*}sheetData/{*}row")]})
        result["mediaSha256"] = sorted(sha256(archive.read(part)).hexdigest() for part in archive.namelist() if part.startswith("xl/media/"))
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    fonts = json.loads((ROOT / "examples/xlsx/reference.json").read_text())["environment"]["fonts"]
    font_digest = sha256(json.dumps(fonts, sort_keys=True, separators=(",", ":")).encode()).hexdigest()
    if args.record or args.verify_native:
        assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
        os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
        os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
        environment = shared.environment()
        assert environment["libreOffice"] == axis.BUILD and environment["fonts"] == fonts, "The pinned Office/font environment differs."
    directory = ROOT / ".cache/workbook-axis-drawings-native"
    directory.mkdir(exist_ok=True)
    sources = []
    for source_name in SOURCES:
        source = ROOT / ("examples/xlsx/" + source_name + ".xlsx")
        cases = []
        for case_name, operation in CASES:
            producer = FIXTURES / ("workbook-axis-drawings-" + source_name + "-" + case_name + "-calc.xlsx")
            if args.record or args.verify_native:
                output = directory / (source_name + "-" + case_name + ".xlsx")
                axis.native(source, output, operation or {}, (lambda document, prop: None) if operation is None else None)
                if args.record:
                    producer.write_bytes(output.read_bytes())
                observed = observe(output)
            else:
                observed = observe(producer)
            cases.append({"name": case_name, "operation": operation, "producerSha256": sha256(producer.read_bytes()).hexdigest(), "observed": observed})
        sources.append({"name": source_name, "path": str(source.relative_to(ROOT)), "sha256": sha256(source.read_bytes()).hexdigest(), "original": observe(source), "cases": cases})
    if args.record or args.verify_native:
        audit_schema(directory, stdout=subprocess.PIPE, text=True)
    report = {"scope": "Three independent sources, a resave baseline and eight actual axis edits per source, with pinned fonts and official XSD/OPC. Source anchors and native producer geometry are recorded separately. Excel/WPS runtime unverified.",
              "office": axis.BUILD, "archiveSha256": axis.ARCHIVE_HASH, "fontManifest": {"source": "examples/xlsx/reference.json#environment.fonts", "sha256": font_digest},
              "references": REFERENCES, "producerDifferences": [
                  "Calc converts one-cell and absolute source containers to two-cell forms and rewrites transformed geometry/cache coordinates on resave; original native source anchors remain separate evidence.",
                  "Some first-open one-cell edits retain the old physical position in Calc. SDK movement follows the Microsoft oneCellAnchor source contract, which moves with its cell and retains physical extents.",
                  "Calc may retain a classic box size on an internal row deletion despite SizeWithCells. SDK applies both moving/resizing markers when both source flags are true; native observations are not replaced with SDK expectations."], "sources": sources}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Pinned drawing/note axis sources or producer observations changed."
    print("Verified 27 drawing/note source/producer cases" + (" with actual pinned Calc, fixed fonts and official XSD/OPC." if args.record or args.verify_native else "."))


if __name__ == "__main__":
    main()
