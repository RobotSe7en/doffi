"""Combine independent actual-data and drawing sources without SDK conversion."""
import argparse
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "packages/office/tests/fixtures/workbook-editing-scale.xlsx"
SHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", SHEET)
ET.register_namespace("r", OFFICE)


def generate():
    with ZipFile(ROOT / "examples/xlsx/15-ten-thousand-rows.xlsx") as source:
        files = {name: source.read(name) for name in source.namelist()}
    with ZipFile(ROOT / "examples/xlsx/08-images-shapes.xlsx") as source:
        drawing = {name: source.read(name) for name in source.namelist()}
    for name, content in drawing.items():
        if name.startswith(("xl/drawings/", "xl/media/")):
            files[name] = content
    # Exercise a fixed-size, moving two-marker owner, which requires the
    # editing boundary to preserve physical geometry across row reflow.
    native_drawing = ET.fromstring(files["xl/drawings/drawing1.xml"])
    for anchor in native_drawing:
        if anchor.tag.endswith("}twoCellAnchor"):
            anchor.set("editAs", "oneCell")
    files["xl/drawings/drawing1.xml"] = ET.tostring(native_drawing, encoding="utf-8", xml_declaration=True)
    types = ET.fromstring(files["[Content_Types].xml"])
    for item in ET.fromstring(drawing["[Content_Types].xml"]):
        if (item.get("PartName", "").startswith(("/xl/drawings/", "/xl/media/")) or item.get("Extension") == "png") and not any(old.attrib == item.attrib for old in types):
            types.append(item)
    files["[Content_Types].xml"] = ET.tostring(types, encoding="utf-8", xml_declaration=True)
    sheet = ET.fromstring(files["xl/worksheets/sheet1.xml"])
    cells = sheet.findall(".//{" + SHEET + "}c")
    assert len(cells) == 200000
    assert sum(cell.find("{" + SHEET + "}f") is not None for cell in cells) == 20000
    ET.SubElement(sheet, "{" + SHEET + "}drawing", {"{" + OFFICE + "}id": "editing-scale-drawing"})
    files["xl/worksheets/sheet1.xml"] = ET.tostring(sheet, encoding="utf-8", xml_declaration=True)
    rels = ET.Element("{" + RELS + "}Relationships")
    ET.SubElement(rels, "{" + RELS + "}Relationship", {"Id": "editing-scale-drawing", "Type": OFFICE + "/drawing", "Target": "../drawings/drawing1.xml"})
    files["xl/worksheets/_rels/sheet1.xml.rels"] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)
    result = BytesIO()
    with ZipFile(result, "w", compression=ZIP_DEFLATED) as target:
        for name, content in sorted(files.items()):
            info = ZipInfo(name, date_time=(2024, 1, 1, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            target.writestr(info, content)
    return result.getvalue()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    args = parser.parse_args()
    content = generate()
    if args.record:
        OUTPUT.write_bytes(content)
    else:
        with ZipFile(OUTPUT) as saved, ZipFile(BytesIO(content)) as generated:
            assert sorted(saved.namelist()) == sorted(generated.namelist())
            assert all(saved.read(name) == generated.read(name) for name in saved.namelist()), "Editing capacity parts differ from their independent inputs."
    print("Verified 200000 native cells, 20000 formulas and the independent picture/shape drawing parts.")
