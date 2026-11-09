"""Actual pinned Calc comment copies from independent native sources."""
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
spec = importlib.util.spec_from_file_location("copy_comment_axis", ROOT / "scripts/check-workbook-axis.py")
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)
REPORT = ROOT / "docs/evidence/workbook-copy-comments.json"
CASES = [("classic", "examples/xlsx/09-classic-notes.xlsx"), ("classic-producer", "examples/xlsx/producer/09-classic-notes.xlsx"),
         ("modern", "examples/xlsx/10-modern-threads.xlsx"), ("dual-box", "packages/office/tests/fixtures/xlsx-dual-note-boxes-conflict.xlsx")]


def observe(path):
    with ZipFile(path) as archive:
        comments, modern, duplicate_ids = [], [], []
        for name in archive.namelist():
            if name.endswith(".vml"):
                counts = {}
                for node in ET.fromstring(archive.read(name)).iter():
                    identity = node.get("id")
                    if identity:
                        counts[identity] = counts.get(identity, 0) + 1
                duplicate_ids += [{"part": name, "id": identity, "count": count} for identity, count in sorted(counts.items()) if count > 1]
            if not name.endswith(".xml"):
                continue
            if name.startswith("xl/comments"):
                root = ET.fromstring(archive.read(name))
                authors = [node.text or "" for node in root.findall("{*}authors/{*}author")]
                for node in root.findall("{*}commentList/{*}comment"):
                    comments.append({"cell": node.get("ref"), "author": authors[int(node.get("authorId", "0"))],
                                     "text": "".join(value.text or "" for value in node.findall("{*}text//{*}t")), "commentPr": node.find("{*}commentPr") is not None})
            if name.startswith("xl/threadedComments"):
                modern += [{"cell": node.get("ref"), "text": node.findtext("{*}text")} for node in ET.fromstring(archive.read(name))]
        return {"classic": comments, "modern": modern, "duplicateVmlIds": duplicate_ids}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    if not args.record and not args.verify_native:
        report = json.loads(REPORT.read_text())
        assert [(case["name"], case["source"]) for case in report["cases"]] == CASES
        for case in report["cases"]:
            assert sha256((ROOT / case["source"]).read_bytes()).hexdigest() == case["sourceSha256"]
            assert sha256((ROOT / case["outputFile"]).read_bytes()).hexdigest() == case["outputSha256"]
            assert observe(ROOT / case["outputFile"]) == case["observed"]
        print("Verified four independent comment-copy sources and immutable producer observations.")
        return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-copy-comments-native"
    directory.mkdir(exist_ok=True)
    cases = []
    for name, source in CASES:
        output = directory / (name + ".xlsx")
        call = {}

        def mutate(document, prop):
            document.calculateAll()
            sheet = document.getSheets().getByIndex(0)
            source_cell, target_cell = sheet.getCellRangeByName("B2"), sheet.getCellRangeByName("G7")
            text = source_cell.getString()
            assert text, "Independent copy source must contain real cell data."
            sheet.copyRange(target_cell.getCellAddress(), source_cell.getRangeAddress())
            assert target_cell.getString() == text, "The native copy did not copy its source cell."
            call.update(sourceText=text, targetText=target_cell.getString())

        version = axis.native(ROOT / source, output, None, mutate)
        source_notes, observed = observe(ROOT / source), observe(output)
        original_note = next((note for note in source_notes["classic"] if note["cell"] == "B2"), None)
        copied_note = next((note for note in observed["classic"] if note["cell"] == "G7"), None)
        call.update(hadClassicNote=original_note is not None, copiedClassicNote=original_note is not None and copied_note is not None and original_note["text"] == copied_note["text"] and original_note["author"] == copied_note["author"])
        recorded = ROOT / "packages/office/tests/fixtures" / ("workbook-copy-comments-" + name + "-calc.xlsx")
        if args.record:
            recorded.write_bytes(output.read_bytes())
        cases.append({"name": name, "source": source, "sourceSha256": sha256((ROOT / source).read_bytes()).hexdigest(), "nativeCall": call,
                      "observed": observed, "outputFile": recorded.relative_to(ROOT).as_posix(), "outputSha256": sha256(recorded.read_bytes()).hexdigest()})
    audit = ROOT / ".cache/xlsx-venv/bin/python"
    subprocess.run([str(audit) if audit.is_file() else sys.executable, str(ROOT / "scripts/check-xlsx-output.py"), str(directory)], check=True)
    report = {"scope": "Independent native inputs and four real pinned Calc copyRange edits. Cell strings prove each copy; persisted XML proves classic note copies. Duplicate VML IDs remain malformed producer negative cases and are rejected by the SDK. Modern thread and alternate-box losses remain explicit producer differences; no Excel/WPS editing runtime claim.",
              "references": ["https://support.microsoft.com/en-au/excel/paste-options?nochrome=true", "https://support.microsoft.com/en-us/excel/insert-comments-and-notes-in-excel"],
              "office": version, "archiveSha256": axis.ARCHIVE_HASH, "cases": cases}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert report == json.loads(REPORT.read_text()), "Pinned native comment-copy observations changed."
    print("Verified four actual pinned Calc comment copies and official XSD/OPC.")


if __name__ == "__main__":
    main()
