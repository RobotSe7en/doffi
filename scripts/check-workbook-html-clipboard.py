"""Pinned native Calc controller HTML transfers from independent XLSX inputs."""
from hashlib import sha256
from html.parser import HTMLParser
from pathlib import Path
import argparse
import importlib.util
import json

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("html_clipboard_axis", ROOT / "scripts/check-workbook-axis.py")
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)
REPORT = ROOT / "docs/evidence/workbook-html-clipboard.json"
CASES = [("values", "examples/xlsx/01-values-text.xlsx"), ("styles", "examples/xlsx/02-styles-dates.xlsx"), ("structure", "examples/xlsx/03-structure.xlsx")]


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
            assert sha256((ROOT / case["htmlFile"]).read_bytes()).hexdigest() == case["htmlSha256"]
        print("Verified three independent native HTML clipboard transfers and immutable source hashes.")
        return
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == axis.ARCHIVE_HASH
    directory = ROOT / ".cache/workbook-html-clipboard-native"
    directory.mkdir(exist_ok=True)
    cases = []
    for name, source in CASES:
        captured = []
        visibility = {}

        def read_transfer(document, prop):
            controller = document.getCurrentController()
            sheet = document.getSheets().getByIndex(0)
            visibility.update(sourceRows=[row for row in range(8) if sheet.getRows().getByIndex(row).IsVisible],
                              sourceColumns=[column for column in range(6) if sheet.getColumns().getByIndex(column).IsVisible])
            controller.select(sheet.getCellRangeByName("A1:F8"))
            transfer = controller.getTransferable()
            flavors = [flavor for flavor in transfer.getTransferDataFlavors() if flavor.MimeType == "text/html"]
            assert len(flavors) == 1, "Native controller did not provide one HTML transfer."
            value = transfer.getTransferData(flavors[0])
            captured.append(value if isinstance(value, str) else bytes(value.value if hasattr(value, "value") else value).decode("utf-8"))

        version = axis.native(ROOT / source, directory / (name + ".xlsx"), None, read_transfer)
        assert len(captured) == 1 and "<table" in captured[0]
        data = captured[0].encode("utf-8")
        target = ROOT / "packages/office/tests/fixtures" / ("workbook-html-clipboard-" + name + "-calc.html")
        if args.record:
            target.write_bytes(data)
        else:
            assert target.read_bytes() == data, "Pinned native HTML transfer changed."
        conflicts = []

        class Values(HTMLParser):
            def handle_starttag(self, tag, attrs):
                attrs = dict(attrs)
                if tag not in {"td", "th"} or "sdval" not in attrs or "data-sheets-value" not in attrs:
                    return
                value = json.loads(attrs["data-sheets-value"])
                if value.get("1") == 3 and float(attrs["sdval"]) != value.get("3"):
                    conflicts.append({"sdval": attrs["sdval"], "sheetsNumber": value.get("3")})

        Values().feed(captured[0])
        cases.append({"name": name, "source": source, "sourceSha256": sha256((ROOT / source).read_bytes()).hexdigest(),
                      "sheetIndex": 0, "range": "A1:F8", "mime": "text/html", "htmlFile": target.relative_to(ROOT).as_posix(),
                      "htmlSha256": sha256(data).hexdigest(), "bytes": len(data), **visibility, "numericMetadataConflicts": conflicts})
    report = {"scope": "Actual pinned Calc controller getTransferable HTML data from independent inputs. Hidden rows/columns are omitted; emitted merge spans remain original producer bytes. Conflicting truncated Sheets metadata and precise sdval are recorded, with the SDK preserving the precise native number. No SDK export, global OS clipboard, or Excel/WPS editing runtime claim.",
              "office": version, "archiveSha256": axis.ARCHIVE_HASH, "cases": cases}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert report == json.loads(REPORT.read_text()), "Pinned HTML clipboard evidence changed."
    print("Verified three actual pinned Calc controller HTML transfers.")


if __name__ == "__main__":
    main()
