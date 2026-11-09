"""Independent note cuts/sorts retaining both native box sources and producer loss."""
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


def module(name, file):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / file)
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value


base = module("note_transport_drawing", "check-workbook-axis-drawings.py")
dual = module("note_transport_dual", "check-xlsx-dual-note-boxes.py")
REPORT = ROOT / "docs/evidence/workbook-note-transport.json"
SOURCES = [("classic", "examples/xlsx/09-classic-notes.xlsx"), ("dual", "packages/office/tests/fixtures/xlsx-dual-note-boxes-conflict.xlsx")]
CASES = ["cut-same", "cut-other-default-header", "cut-other", "sort"]


def observe(path):
    value = base.observe(path)
    value["nativeBoxes"] = dual.observe(Path(path).read_bytes())
    with ZipFile(path) as archive:
        value["printControls"] = {name: {child.tag.rsplit("}", 1)[-1]: child.text for child in ET.fromstring(archive.read(name)).findall("{*}headerFooter/*")}
                                  for name in sorted(archive.namelist()) if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")}
    return value


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    assert (ROOT / SOURCES[1][1]).read_bytes() == dual.generate("conflict"), "Independent alternate-box source changed."
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert [(item["name"], item["source"]) for item in report["sources"]] == SOURCES
        for item in report["sources"]:
            data = (ROOT / item["source"]).read_bytes()
            assert item["sourceSha256"] == sha256(data).hexdigest() and item["original"] == observe(ROOT / item["source"])
            assert [case["name"] for case in item["cases"]] == CASES
            for case in item["cases"]:
                producer = ROOT / case["outputFile"]
                assert sha256(producer.read_bytes()).hexdigest() == case["outputSha256"] and observe(producer) == case["observed"]
        print("Verified eight independent native note records, both original box sources and two default-header negatives.")
        return
    import uno
    from com.sun.star.table.TableSortFieldType import AUTOMATIC
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == base.axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-note-transport-native"
    directory.mkdir(exist_ok=True)
    sources = []
    for name, path in SOURCES:
        source = ROOT / path
        cases = []
        for case in CASES:
            call = {}

            def mutate(document, prop):
                sheet = document.getSheets().getByIndex(0)
                if case == "sort":
                    for row, key in enumerate([3, 1, 2], 1):
                        sheet.getCellByPosition(1, row).setValue(key)
                        sheet.getCellByPosition(2, row).setString("record-" + str(key))
                    field = uno.createUnoStruct("com.sun.star.table.TableSortField")
                    field.Field, field.IsAscending, field.IsCaseSensitive, field.FieldType = 0, True, False, AUTOMATIC
                    field.CollatorLocale = uno.createUnoStruct("com.sun.star.lang.Locale", "en", "US", "")
                    target = sheet.getCellRangeByName("B2:C4")
                    target.sort((prop("SortFields", uno.Any("[]com.sun.star.table.TableSortField", (field,))), prop("IsSortColumns", False), prop("ContainsHeader", False), prop("BindFormatsToContent", True)))
                    call["sortedKeys"] = [row[0] for row in target.getDataArray()]
                    assert call["sortedKeys"] == [1, 2, 3], "Native source records did not sort."
                else:
                    text = sheet.getCellRangeByName("B2").getString()
                    assert text, "Native note cut requires actual source data."
                    if case.startswith("cut-other"):
                        document.getSheets().insertNewByName("Other", document.getSheets().getCount())
                        target = document.getSheets().getByName("Other")
                        if case == "cut-other":
                            # Select the source's supported native page style
                            # through the actual producer before editing. Raw
                            # new-sheet styled headers remain separate negatives.
                            target.setPropertyValue("PageStyle", sheet.getPropertyValue("PageStyle"))
                    else:
                        target = sheet
                    target.moveRange(target.getCellRangeByName("H20").getCellAddress(), sheet.getCellRangeByName("B2").getRangeAddress())
                    call.update(sourceText=text, targetText=target.getCellRangeByName("H20").getString(), clearedSource=sheet.getCellRangeByName("B2").getString())
                    call.update(sourcePageStyle=sheet.getPropertyValue("PageStyle"), targetPageStyle=target.getPropertyValue("PageStyle"))
                    assert call["targetText"] == text and call["clearedSource"] == "", "Native note cell did not move."

            output = directory / (name + "-" + case + ".xlsx")
            version = base.axis.native(source, output, None, mutate)
            producer = ROOT / "packages/office/tests/fixtures" / ("workbook-note-transport-" + name + "-" + case + "-calc.xlsx")
            if args.record and not (case == "cut-other-default-header" and producer.exists()):
                producer.write_bytes(output.read_bytes())
            cases.append({"name": case, "nativeCall": call, "outputFile": producer.relative_to(ROOT).as_posix(), "outputSha256": sha256(producer.read_bytes()).hexdigest(), "observed": observe(output)})
        sources.append({"name": name, "source": path, "sourceSha256": sha256(source.read_bytes()).hexdigest(), "original": observe(source), "cases": cases})
    audit_python = sys.executable if importlib.util.find_spec("lxml") is not None else str(ROOT / ".cache/xlsx-venv/bin/python")
    subprocess.run([audit_python, str(ROOT / "scripts/check-xlsx-output.py"), str(directory)], check=True, stdout=subprocess.PIPE, text=True)
    report = {"scope": "Two immutable independent native classic-note sources, eight actual pinned Calc cell cuts/sorts with data identity proof and official XSD/OPC. Raw newly created target sheets introduce styled header/footer controls outside the SDK's declared basic print-band support; both original producer files remain unsupported negative cases. Six positive outputs use same-sheet edits or explicitly select the original native page style on the new target before moving data, without patching output bytes. Both VML and commentPr are recorded before editing; producer loss/normalization remains explicit. SDK preserves both sources and physically transports boxes with valid mapped note cells, retaining extents and native/body identities. Synthetic core point metrics are separate from fixed-font browser/native geometry acceptance. Excel/WPS structural runtime unverified.",
              "references": base.REFERENCES, "office": version, "archiveSha256": base.axis.ARCHIVE_HASH, "sources": sources}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert json.loads(REPORT.read_text()) == report, "Pinned native note observations changed."
    print("Verified eight actual pinned note edits, two retained default-header negatives and official XSD/OPC.")


if __name__ == "__main__":
    main()
