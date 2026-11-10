"""Pinned Calc consumes SDK HTML through its actual native transfer controller."""
from workbook_acceptance import audit_schema
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
REPORT = ROOT / "docs/evidence/workbook-system-clipboard.json"
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-system-clipboard-calc.xlsx"
SOURCE = ROOT / "examples/xlsx/01-values-text.xlsx"


def observe(path):
    with ZipFile(path) as archive:
        shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        strings = ["".join(item.itertext()) for item in shared]
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        cells = {}
        for cell in sheet.findall("{*}sheetData/{*}row/{*}c"):
            address = cell.attrib["r"]
            if address not in [column + str(row) for row in range(12, 15) for column in "ABCD"]:
                continue
            value = cell.findtext("{*}v")
            cells[address] = {"type": cell.get("t", "n"), "value": strings[int(value)] if cell.get("t") == "s" else value,
                              "formula": cell.findtext("{*}f")}
        merges = [item.attrib["ref"] for item in sheet.findall("{*}mergeCells/{*}mergeCell")]
        return {"cells": cells, "merges": merges}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert sha256(SOURCE.read_bytes()).hexdigest() == report["sourceSha256"]
        assert sha256(FIXTURE.read_bytes()).hexdigest() == report["fixtureSha256"]
        assert sha256((ROOT / "scripts/workbook-system-clipboard-source.mjs").read_bytes()).hexdigest() == report["generatorSha256"]
        assert observe(FIXTURE) == report["observed"]
        print("Verified native system clipboard consumer values and immutable independent source.")
        return
    import uno
    import unohelper
    from com.sun.star.datatransfer import XTransferable, DataFlavor, UnsupportedFlavorException

    spec = importlib.util.spec_from_file_location("system_clipboard_axis", ROOT / "scripts/check-workbook-axis.py")
    axis = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(axis)
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    os.environ["LC_ALL"] = os.environ["LANG"] = "C.UTF-8"
    directory = ROOT / ".cache/workbook-system-clipboard-native"
    subprocess.run(["node", "scripts/workbook-system-clipboard-source.mjs"], cwd=ROOT, check=True)
    html = (directory / "clipboard.html").read_bytes()
    text = (directory / "clipboard.txt").read_text()

    class Transfer(unohelper.Base, XTransferable):
        def getTransferDataFlavors(self):
            return (DataFlavor("text/html", "HTML", uno.getTypeByName("string")),
                    DataFlavor("text/plain;charset=utf-16", "Text", uno.getTypeByName("string")))

        def isDataFlavorSupported(self, flavor):
            return flavor.MimeType in [item.MimeType for item in self.getTransferDataFlavors()]

        def getTransferData(self, flavor):
            if flavor.MimeType == "text/html":
                return html.decode("utf-8")
            if flavor.MimeType == "text/plain;charset=utf-16":
                return text
            raise UnsupportedFlavorException()

    def paste(document, prop):
        sheet = document.getSheets().getByIndex(0)
        controller = document.getCurrentController()
        controller.setActiveSheet(sheet)
        controller.select(sheet.getCellRangeByName("A12"))
        controller.insertTransferable(Transfer())
        document.calculateAll()
        assert sheet.getCellRangeByName("B12").Value == 25
        assert sheet.getCellRangeByName("D14").Value == 12345.6789, "Native precision must not come from rounded display text."
        assert sheet.getCellRangeByName("B14").Value == 60, "Native date must retain its original numeric serial."
        assert sheet.getCellRangeByName("C14").Formula.startswith("="), "A Boolean formula must remain editable as a formula."
        sheet.getCellRangeByName("A12").Value = 15
        document.calculateAll()
        assert sheet.getCellRangeByName("B12").Value == 30

    output = directory / "consumer.xlsx"
    version = axis.native(SOURCE, output, None, paste)
    result = observe(output)
    assert result["cells"]["C12"]["value"] == "001234567890123456789"
    assert result["cells"]["D13"]["value"] == "=1+1" and result["cells"]["D13"]["formula"] is None
    assert "A13:B13" in result["merges"]
    audit = audit_schema(directory, capture_output=True, text=True)
    report = {"scope": "Independent XLSX imported and edited through SDK transactions; generated portable HTML inserted through the actual pinned native controller, then a native value edit recalculates the pasted formula. Original HTML/SDK text retains CR; producer CR/LF and tab normalization and constant-error formula conversion are recorded. No Excel/WPS runtime claim.",
              "sourceSha256": sha256(SOURCE.read_bytes()).hexdigest(), "htmlSha256": sha256(html).hexdigest(), "office": version,
              "fixtureSha256": sha256((output if args.record else FIXTURE).read_bytes()).hexdigest(),
              "generatorSha256": sha256((ROOT / "scripts/workbook-system-clipboard-source.mjs").read_bytes()).hexdigest(),
              "archiveSha256": axis.ARCHIVE_HASH, "officialSchemasSha256": json.loads(audit.stdout)["schemasSha256"], "observed": result}
    if args.record:
        shutil.copyfile(output, FIXTURE)
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert report == json.loads(REPORT.read_text()), "Pinned system clipboard consumer evidence changed."
    print("Verified pinned native clipboard insertion, formula editing/recalculation and official XSD/OPC.")


if __name__ == "__main__":
    main()
