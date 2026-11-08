"""Author native table styles with Python ZIP/XML, independently of the SDK."""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_STORED, ZipFile, ZipInfo
import argparse
import json
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
BASE = ROOT / "examples/xlsx/05-native-table.xlsx"
TARGET = ROOT / "packages/office/tests/fixtures/workbook-table-style.xlsx"
REPORT = ROOT / "docs/evidence/workbook-table-style-source.json"
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def authored_source():
    with ZipFile(BASE) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    styles = files["xl/styles.xml"].decode("utf-8")
    if len(ET.fromstring(styles).findall("s:dxfs/s:dxf", NS)) < 2:
        raise ValueError("The immutable independent base must supply its two differentials")
    definition = ('<tableStyles count="2" defaultTableStyle="CustomBand">'
                  '<tableStyle name="CustomBand" table="1" pivot="0" count="4">'
                  '<tableStyleElement type="wholeTable" dxfId="0"/>'
                  '<tableStyleElement type="headerRow" dxfId="1"/>'
                  '<tableStyleElement type="totalRow"/>'
                  '<tableStyleElement type="firstRowStripe" size="2" dxfId="0"/>'
                  '</tableStyle><tableStyle name="SharedPaint" table="1" pivot="0" count="1">'
                  '<tableStyleElement type="wholeTable" dxfId="0"/>'
                  '</tableStyle></tableStyles>')
    styles, count = re.subn(r"<tableStyles\b[^>]*/>", definition, styles)
    if count != 1:
        raise ValueError("The independent source table style container changed")
    table = files["xl/tables/table1.xml"].decode("utf-8")
    table, count = re.subn(r'(<tableStyleInfo\b[^>]*\bname=")[^"]+("[^>]*/>)', r'\g<1>CustomBand\2', table)
    if count != 1:
        raise ValueError("The independent source table style association changed")
    files["xl/styles.xml"] = styles.encode("utf-8")
    files["xl/tables/table1.xml"] = table.encode("utf-8")
    output = BytesIO()
    with ZipFile(output, "w", compression=ZIP_STORED) as archive:
        for name, data in sorted(files.items()):
            info = ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_STORED
            info.external_attr = 0o600 << 16
            archive.writestr(info, data)
    result = output.getvalue()
    evidence = {"generator": "Python standard-library ZIP/XML; no SDK exporter", "base": str(BASE.relative_to(ROOT)),
                "baseSha256": sha256(BASE.read_bytes()).hexdigest(), "source": str(TARGET.relative_to(ROOT)), "sha256": sha256(result).hexdigest(),
                "styles": ["CustomBand", "SharedPaint"], "regions": ["wholeTable", "headerRow", "totalRow", "firstRowStripe"],
                "scope": "Independent native input; producer observations are recorded separately", "producerEvidence": "docs/evidence/workbook-table-style-native.json"}
    return result, (json.dumps(evidence, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    source, report = authored_source()
    if args.check:
        if TARGET.read_bytes() != source or REPORT.read_bytes() != report:
            raise SystemExit("Independent table style fixture/evidence differs; regenerate and review")
        print("Independent native table style source and hashes passed")
    else:
        TARGET.write_bytes(source)
        REPORT.write_bytes(report)
        print(f"Wrote {TARGET.relative_to(ROOT)}")
