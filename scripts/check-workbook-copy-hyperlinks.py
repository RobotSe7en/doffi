"""Pinned native range-copy observations for immutable independent hyperlinks."""
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
SOURCE = ROOT / "examples/xlsx/01-values-text.xlsx"
FIXTURES = ROOT / "packages/office/tests/fixtures"
REPORT = ROOT / "docs/evidence/workbook-copy-hyperlinks.json"
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def observe(path):
    with ZipFile(path) as archive:
        strings = ["".join(item.itertext()) for item in ET.fromstring(archive.read("xl/sharedStrings.xml"))] if "xl/sharedStrings.xml" in archive.namelist() else []
        result = {}
        for name in sorted(archive.namelist()):
            if not name.startswith("xl/worksheets/sheet") or not name.endswith(".xml"):
                continue
            sheet = ET.fromstring(archive.read(name))
            relationships = name.rsplit("/", 1)[0] + "/_rels/" + name.rsplit("/", 1)[1] + ".rels"
            links = {item.get("Id"): item.get("Target") for item in ET.fromstring(archive.read(relationships))} if relationships in archive.namelist() else {}
            values = {}
            for cell in sheet.findall("s:sheetData/s:row/s:c", NS):
                if cell.get("r") not in ["A2", "A3", "A4", "AG257"]:
                    continue
                value = cell.findtext("s:v", default="", namespaces=NS)
                if cell.get("t") == "s":
                    value = strings[int(value)]
                elif cell.get("t") == "inlineStr":
                    value = "".join(text.text or "" for text in cell.findall("s:is//s:t", NS))
                values[cell.get("r")] = {"type": cell.get("t", "n"), "value": value, "formula": cell.findtext("s:f", default="", namespaces=NS)}
            result[name] = {"links": [{**{key: value for key, value in item.attrib.items() if key != REL_ID}, "url": links.get(item.get(REL_ID))}
                                      for item in sheet.findall("s:hyperlinks/s:hyperlink", NS)], "cells": values, "printBands": {band.tag.rsplit("}", 1)[-1]: band.text or "" for band in sheet.find("s:headerFooter", NS)} if sheet.find("s:headerFooter", NS) is not None else {}}
        return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    spec = importlib.util.spec_from_file_location("copy_hyperlinks_calc", ROOT / "scripts/check-workbook-axis.py")
    axis = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(axis)
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    os.environ["LC_ALL"] = os.environ["LANG"] = "C.UTF-8"
    source_hash = sha256(SOURCE.read_bytes()).hexdigest()
    directory = ROOT / ".cache/workbook-copy-hyperlinks-native"
    directory.mkdir(exist_ok=True)
    (directory / "source.xlsx").write_bytes(SOURCE.read_bytes())
    cases, outputs = [], []
    for name in ["copy", "overlap", "other", "new-sheet"]:
        recorded = FIXTURES / ("workbook-copy-hyperlinks-" + name + "-calc.xlsx")
        output = directory / (name + ".xlsx")
        if args.record or args.verify_native:
            def mutate(document, prop):
                sheets = document.getSheets()
                source = sheets.getByIndex(0)
                target = source
                if name in ["other", "new-sheet"]:
                    if name == "other":
                        sheets.copyByName(source.getName(), "Other", 1)
                    else:
                        sheets.insertNewByName("Other", 1)
                    target = sheets.getByIndex(1)
                source_range = source.getCellRangeByName("A2:A3" if name == "overlap" else "A2").getRangeAddress()
                at = target.getCellByPosition(0 if name == "overlap" else 32, 2 if name == "overlap" else 256).getCellAddress()
                target.copyRange(at, source_range)
            axis.native(SOURCE, output, {}, mutate)
        else:
            output.write_bytes(recorded.read_bytes())
        assert sha256(SOURCE.read_bytes()).hexdigest() == source_hash
        current = observe(output)
        destination = current["xl/worksheets/sheet2.xml" if name in ["other", "new-sheet"] else "xl/worksheets/sheet1.xml"]["links"]
        assert any(link["ref"] == ("A3" if name == "overlap" else "AG257") and link["url"] == "https://example.com/authorized-link" for link in destination)
        cases.append({"name": name, "sdkAcceptance": "rejected: formatted print header/footer controls" if name == "new-sheet" else "supported", "output": recorded.relative_to(ROOT).as_posix(), "sha256": sha256(output.read_bytes() if args.record else recorded.read_bytes()).hexdigest(), "observed": current})
        outputs.append((recorded, output))
    audit_python = ROOT / ".cache/xlsx-venv/bin/python"
    audit = subprocess.run([str(audit_python) if audit_python.is_file() else sys.executable, str(ROOT / "scripts/check-xlsx-output.py"), str(directory)], check=True, capture_output=True, text=True)
    report = {"scope": "Immutable Python ZIP/XML source and actual pinned Calc ordinary, overlapping and cross-sheet copies, plus an explicit negative producer case for new-sheet default print font controls. Display-cache regeneration and omitted tooltip are retained as producer differences; no Excel/WPS runtime or SDK native XLSX export claim.",
              "source": SOURCE.relative_to(ROOT).as_posix(), "sourceSha256": source_hash, "sourceObserved": observe(SOURCE), "office": axis.BUILD, "archiveSha256": axis.ARCHIVE_HASH, "nativeLocale": "C.UTF-8",
              "officialSchemasSha256": json.loads(audit.stdout)["schemasSha256"], "cases": cases}
    if args.record:
        for recorded, output in outputs:
            recorded.write_bytes(output.read_bytes())
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert report == json.loads(REPORT.read_text()), "Native copied hyperlink observations changed"
    print("Verified four actual pinned Calc hyperlink copies and five official XSD/OPC inputs; producer differences and unsupported print controls retained")


if __name__ == "__main__":
    main()
