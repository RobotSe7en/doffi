"""Independent ZIP/XML source and actual pinned Calc row/column edits."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import argparse
import importlib.util
import json
import os
import signal
import socket
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parent.parent
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-axis.xlsx"
REPORT = ROOT / "docs/evidence/workbook-axis.json"
ARCHIVE_HASH = "f367c434aee8352e9742e192bedcdaef194ffcbcfeaf760c6300fb4fecbba73c"
BUILD = "LibreOffice 25.8.1.1 54047653041915e595ad4e45cccea684809c77b5"
CASES = [
    {"axis": "row", "action": "insert", "index": 1, "count": 2},
    {"axis": "row", "action": "delete", "index": 1, "count": 2},
    {"axis": "column", "action": "insert", "index": 1, "count": 2},
    {"axis": "column", "action": "delete", "index": 1, "count": 1},
]
spec = importlib.util.spec_from_file_location("independent_axis_source", ROOT / "scripts/generate-xlsx-fixtures.py")
source = importlib.util.module_from_spec(spec)
spec.loader.exec_module(source)


def generate():
    book = source.Book()
    sheet = book.sheet("Data")
    for row, value in enumerate([10, 20, 30, 40], 1):
        sheet.add("A" + str(row), value)
    sheet.add("A256", 7)
    sheet.add("A257", 8)
    sheet.add("AF256", 3)
    sheet.add("AG257", 4)
    sheet.add("B1", 100, formula="SUM(A1:A4)")
    sheet.add("B2", 31, formula="'Observer'!A1+$A$3")
    sheet.add("C1", "中文😀禰󠄀", style=9)
    sheet.row_props[2] = {"ht": 20, "customHeight": 1}
    sheet.before = '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15" defaultColWidth="12"/>'
    sheet.after = '<mergeCells count="1"><mergeCell ref="D3:E4"/></mergeCells><conditionalFormatting sqref="A1:A4"><cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="custom" allowBlank="1" showErrorMessage="1" errorStyle="stop" sqref="A1:A4"><formula1>A1&gt;0</formula1></dataValidation></dataValidations>'
    sheet.print_area = "$A$1:$AG$260"
    other = book.sheet("Observer")
    other.add("A1", 1)
    other.add("B1", 30, formula="Data!A3")
    return source.zipped(book.finish())


def observe(path):
    ns = {"s": source.S}
    result = []
    with ZipFile(path) as archive:
        for name in ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"]:
            root = ET.fromstring(archive.read(name))
            formulas = [{"address": cell.attrib["r"], "formula": cell.find("s:f", ns).text,
                         "value": cell.findtext("s:v", namespaces=ns), "type": cell.attrib.get("t", "n")}
                        for cell in root.findall(".//s:sheetData/s:row/s:c", ns) if cell.find("s:f", ns) is not None]
            result.append({"formulas": formulas,
                           "merges": [item.attrib["ref"] for item in root.findall("s:mergeCells/s:mergeCell", ns)],
                           "conditional": [{"ranges": item.attrib["sqref"], "formulas": [node.text for node in item.findall(".//s:formula", ns)]}
                                           for item in root.findall("s:conditionalFormatting", ns)],
                           "validation": [{"ranges": item.attrib["sqref"], "formula": item.findtext("s:formula1", namespaces=ns)}
                                          for item in root.findall("s:dataValidations/s:dataValidation", ns)]})
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        names = [{"name": item.attrib["name"], "scope": item.attrib.get("localSheetId"), "expression": item.text}
                 for item in workbook.findall("s:definedNames/s:definedName", ns)]
    return {"sheets": result, "names": names}


def native(source_path, output, case):
    # pyuno is an acceptance-only client. No Office runtime enters the SDK.
    import uno
    from com.sun.star.beans import PropertyValue

    def prop(name, value):
        result = PropertyValue()
        result.Name, result.Value = name, value
        return result

    with TemporaryDirectory(prefix="doffi-axis-") as temporary:
        with socket.socket() as port:
            port.bind(("127.0.0.1", 0))
            number = port.getsockname()[1]
        environment = {key: value for key, value in os.environ.items() if key not in {"URE_BOOTSTRAP", "UNO_PATH", "UNO_TYPES", "UNO_SERVICES"}}
        environment["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
        binary = str(ROOT / ".cache/acceptance-office/bin/libreoffice")
        version = subprocess.check_output([binary, "--version"], env=environment, text=True).strip()
        assert version == BUILD, version
        process = subprocess.Popen([binary, "-env:UserInstallation=" + Path(temporary).as_uri(), "--headless", "--norestore", "--nodefault",
                                    "--nofirststartwizard", f"--accept=socket,host=127.0.0.1,port={number};urp;StarOffice.ComponentContext"],
                                   env=environment, stdout=subprocess.PIPE, stderr=subprocess.PIPE, start_new_session=True)
        document = None
        try:
            resolver = uno.getComponentContext().ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", uno.getComponentContext())
            deadline = time.monotonic() + 25
            while True:
                try:
                    context = resolver.resolve(f"uno:socket,host=127.0.0.1,port={number};urp;StarOffice.ComponentContext")
                    break
                except Exception:
                    if time.monotonic() >= deadline or process.poll() is not None:
                        raise RuntimeError("The pinned Calc edit process did not become ready.")
                    time.sleep(0.1)
            desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
            document = desktop.loadComponentFromURL(source_path.resolve().as_uri(), "_blank", 0,
                                                    (prop("Hidden", True), prop("ReadOnly", False), prop("UpdateDocMode", 0), prop("MacroExecutionMode", 4)))
            if document is None:
                raise RuntimeError("The independent source did not open in Calc.")
            sheet = document.getSheets().getByIndex(0)
            axes = sheet.getRows() if case["axis"] == "row" else sheet.getColumns()
            method = axes.insertByIndex if case["action"] == "insert" else axes.removeByIndex
            method(case["index"], case["count"])
            document.calculateAll()
            document.storeAsURL(output.resolve().as_uri(), (prop("FilterName", "Calc MS Excel 2007 XML"), prop("Overwrite", True)))
            return version
        finally:
            if document is not None:
                try:
                    document.close(True)
                except Exception:
                    pass  # The isolated process must still be reaped.
            try:
                os.killpg(process.pid, signal.SIGTERM)
            except ProcessLookupError:
                pass
            try:
                process.communicate(timeout=3)
            except subprocess.TimeoutExpired:
                os.killpg(process.pid, signal.SIGKILL)
                process.communicate()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    data = generate()
    if not args.record:
        assert FIXTURE.read_bytes() == data, "Independent axis source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(data).hexdigest()
        assert [item["operation"] for item in report["cases"]] == CASES
        print("Verified four independent axis-edit source cases.")
        return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-axis-native"
    directory.mkdir(exist_ok=True)
    input_path = directory / "source.xlsx"
    input_path.write_bytes(data)
    cases = []
    for index, operation in enumerate(CASES):
        output = directory / (str(index) + ".xlsx")
        version = native(input_path, output, operation)
        cases.append({"operation": operation, "observed": observe(output)})
    # The pyuno client and lxml may live in separate local Python installs.
    # CI provides both to /usr/bin/python3; local schema checks can reuse the
    # already provisioned acceptance venv without changing system packages.
    audit_python = sys.executable
    local_python = ROOT / ".cache/xlsx-venv/bin/python"
    if importlib.util.find_spec("lxml") is None and local_python.is_file():
        audit_python = str(local_python)
    subprocess.run([audit_python, str(ROOT / "scripts/check-xlsx-output.py"), str(directory)], check=True,
                   stdout=subprocess.PIPE, text=True)
    report = {"scope": "Independent ZIP/XML input, actual row/column edits in pinned Calc and native formula/range observations. Excel/WPS unverified.",
              "office": version, "archiveSha256": ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(), "cases": cases}
    if args.record:
        FIXTURE.write_bytes(data)
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Pinned native axis observations changed."
    print("Verified four actual pinned Calc axis edits.")


if __name__ == "__main__":
    main()
