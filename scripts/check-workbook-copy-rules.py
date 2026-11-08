"""Actual bounded copies of immutable conditional formats and validation rules."""
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
SOURCE = ROOT / "examples/xlsx/06-rules-validation.xlsx"
FIXTURES = ROOT / "packages/office/tests/fixtures"
REPORT = ROOT / "docs/evidence/workbook-copy-rules.json"
CASES = {"ordinary": ("A2:E2", "AG257"), "partial-scale": ("A3:D4", "AG257"), "overlap": ("B2:B4", "B3"),
         "validations": ("E4:E8", "AG257"), "other": ("A2:E2", "AG257"), "invalid-stop": ("E3", "AG257")}


def observe(path):
    with ZipFile(path) as archive:
        result = {}
        for name in sorted(archive.namelist()):
            if not name.startswith("xl/worksheets/sheet") or not name.endswith(".xml"):
                continue
            sheet = ET.fromstring(archive.read(name))
            def node(item):
                return {"kind": item.tag.rsplit("}", 1)[-1], "attributes": item.attrib, "text": item.text or "", "children": [node(child) for child in item]}
            result[name] = {"conditionalRules": [node(item) for item in sheet.findall("{*}conditionalFormatting")],
                            "dataValidations": [node(item) for item in sheet.findall("{*}dataValidations/{*}dataValidation")],
                            "copiedValues": [node(item) for item in sheet.findall("{*}sheetData/{*}row/{*}c") if item.get("r", "").startswith(("AG", "AH", "AI", "AJ", "AK"))]}
        return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    spec = importlib.util.spec_from_file_location("copy_rules_axis", ROOT / "scripts/check-workbook-axis.py")
    axis = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(axis)
    source_hash = sha256(SOURCE.read_bytes()).hexdigest()
    if args.record or args.verify_native:
        assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    os.environ["LC_ALL"] = os.environ["LANG"] = "C.UTF-8"
    directory = ROOT / ".cache/workbook-copy-rules-native"
    directory.mkdir(exist_ok=True)
    (directory / "source.xlsx").write_bytes(SOURCE.read_bytes())
    cases = []
    for name, (source_range, destination) in CASES.items():
        recorded = FIXTURES / ("workbook-copy-rules-" + name + "-calc.xlsx")
        output = directory / (name + ".xlsx")
        if args.record or args.verify_native:
            def mutate(document, prop):
                sheets = document.getSheets()
                source = sheets.getByIndex(0)
                target = source
                if name == "other":
                    sheets.copyByName(source.getName(), "Other", 1)
                    target = sheets.getByIndex(1)
                target.copyRange(target.getCellRangeByName(destination).getCellAddress(), source.getCellRangeByName(source_range).getRangeAddress())
            axis.native(SOURCE, output, {}, mutate)
        else:
            output.write_bytes(recorded.read_bytes())
        assert sha256(SOURCE.read_bytes()).hexdigest() == source_hash
        current = observe(output)
        if args.record:
            recorded.write_bytes(output.read_bytes())
        cases.append({"name": name, "sourceRange": source_range, "destination": destination, "sdkCopy": "rejected: copied stop violation" if name == "invalid-stop" else "supported",
                      "output": recorded.relative_to(ROOT).as_posix(), "sha256": sha256(recorded.read_bytes()).hexdigest(), "observed": current})
    audit_python = ROOT / ".cache/xlsx-venv/bin/python"
    audit = subprocess.run([str(audit_python) if audit_python.is_file() else sys.executable, str(ROOT / "scripts/check-xlsx-output.py"), str(directory)], check=True, capture_output=True, text=True)
    report = {"scope": "Immutable independent ZIP/XML source, six actual pinned Calc copies and official Transitional XSD/OPC. SDK uses uniform transaction validation: copied stop violations are rejected even though native copy accepts them. Native conditional-rule grouping and producer changes are preserved. No Excel/WPS runtime claim.",
              "source": SOURCE.relative_to(ROOT).as_posix(), "sourceSha256": source_hash, "sourceObserved": observe(SOURCE), "office": axis.BUILD, "archiveSha256": axis.ARCHIVE_HASH,
              "officialSchemasSha256": json.loads(audit.stdout)["schemasSha256"], "nativeLocale": "C.UTF-8", "cases": cases}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert report == json.loads(REPORT.read_text()), "Native copied-rule observations changed"
    print("Verified six pinned Calc rule copies and seven official XSD/OPC inputs; stop-violation difference retained")


if __name__ == "__main__":
    main()
