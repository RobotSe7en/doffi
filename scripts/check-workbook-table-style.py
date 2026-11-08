"""Reproduce custom-style producer losses and real edits with pinned Calc."""
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
FIXTURES = ROOT / "packages/office/tests/fixtures"
SOURCE = FIXTURES / "workbook-table-style.xlsx"
REPORT = ROOT / "docs/evidence/workbook-table-style-native.json"
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def module(name, path):
    spec = importlib.util.spec_from_file_location(name, ROOT / path)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


def observe(path):
    with ZipFile(path) as archive:
        styles = ET.fromstring(archive.read("xl/styles.xml"))
        definitions = styles.find("s:tableStyles", NS)
        tables = []
        for name in sorted(archive.namelist()):
            if name.startswith("xl/tables/") and name.endswith(".xml"):
                table = ET.fromstring(archive.read(name))
                style = table.find("s:tableStyleInfo", NS)
                tables.append({"name": table.get("name"), "displayName": table.get("displayName"), "range": table.get("ref"),
                               "style": None if style is None else style.attrib})
        cells = {}
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        for cell in sheet.findall("s:sheetData/s:row/s:c", NS):
            if cell.get("r") in ["B2", "D2"]:
                cells[cell.get("r")] = {"type": cell.get("t", "n"), "style": cell.get("s", "0"),
                                        "value": cell.findtext("s:v", default="", namespaces=NS),
                                        "formula": cell.findtext("s:f", default="", namespaces=NS)}
        return {"customStyles": [] if definitions is None else [
            {"attributes": style.attrib, "elements": [element.attrib for element in style]} for style in definitions],
                "defaultStyle": None if definitions is None else definitions.get("defaultTableStyle"), "tables": tables, "cells": cells}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    generator = module("table_style_source", "scripts/generate-workbook-table-style-fixture.py")
    source, evidence = generator.authored_source()
    assert SOURCE.read_bytes() == source
    assert generator.REPORT.read_bytes() == evidence
    axis = module("table_style_calc", "scripts/check-workbook-axis.py")
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    os.environ["LC_ALL"] = os.environ["LANG"] = "C.UTF-8"
    directory = ROOT / ".cache/workbook-table-style-native"
    directory.mkdir(exist_ok=True)
    (directory / "source.xlsx").write_bytes(source)
    cases, outputs = [], []
    for name in ["resave", "edit"]:
        recorded = FIXTURES / ("workbook-table-style-" + name + "-calc.xlsx")
        output = directory / (name + ".xlsx")
        if args.record or args.verify_native:
            def mutate(document, prop):
                if name == "edit":
                    document.getSheets().getByIndex(0).getCellByPosition(1, 1).setValue(9)
            axis.native(SOURCE, output, {}, mutate)
        else:
            output.write_bytes(recorded.read_bytes())
        observed = observe(output)
        if name == "edit":
            assert observed["cells"]["B2"]["value"] == "9"
        cases.append({"name": name, "output": recorded.relative_to(ROOT).as_posix(),
                      "sha256": sha256(output.read_bytes() if args.record else recorded.read_bytes()).hexdigest(), "observed": observed})
        outputs.append((recorded, output))
    audit_python = ROOT / ".cache/xlsx-venv/bin/python"
    audit = subprocess.run([str(audit_python) if audit_python.is_file() else sys.executable, str(ROOT / "scripts/check-xlsx-output.py"), str(directory)], check=True, capture_output=True, text=True)
    schemas = json.loads(audit.stdout)["schemasSha256"]
    report = {"scope": "Independent native custom table styles, pinned Calc resave and an actual B2 edit. Producer loss of custom definitions/default/associations is retained explicitly; not an SDK lossless-export or Excel/WPS runtime claim.",
              "source": SOURCE.relative_to(ROOT).as_posix(), "sourceSha256": sha256(source).hexdigest(), "sourceObserved": observe(SOURCE),
              "office": axis.BUILD, "archiveSha256": axis.ARCHIVE_HASH, "nativeLocale": "C.UTF-8", "officialSchemasSha256": schemas, "cases": cases}
    if args.record:
        for recorded, output in outputs:
            recorded.write_bytes(output.read_bytes())
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert report == json.loads(REPORT.read_text()), "Native table style source/producer observations changed"
    print("Verified independent source, two pinned Calc cases and official XSD/OPC checks; producer losses remain explicit")


if __name__ == "__main__":
    main()
