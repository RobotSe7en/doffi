"""Independent official Transitional XSD and selected MS-XLSX thread checks.

This acceptance tool does not import the SDK. It checks native structures and
relationships; the SDK's declared feature gate is tested separately.
"""
from contextlib import contextmanager
from copy import deepcopy
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.parse import unquote
from zipfile import ZipFile
import json
import posixpath
import re
import sys
from lxml import etree as E

ROOT = Path(__file__).resolve().parent.parent
ARCHIVE = ROOT / "scripts/fixtures/ooxml/OfficeOpenXML-XMLSchema-Transitional.zip"
EXPECTED = "d34187520749998af306faf1b730e568b0ca6d88ad24638a407c0a9bb4ca04fc"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
D = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
ROW_VISIBILITY = "urn:doffi:xlsx:row-visibility:1"
GUID = re.compile(r"\{[0-9A-Fa-f]{8}-(?:[0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}\}")


def coordinate(address):
    match = re.fullmatch(r"([A-Z]{1,3})([1-9]\d{0,6})", address or "")
    assert match, "Invalid cell coordinate: " + str(address)
    col = 0
    for char in match[1]: col = col * 26 + ord(char) - 64
    row = int(match[2])
    assert 1 <= col <= 16384 and 1 <= row <= 1048576, "Cell outside native bounds"
    return row, col


@contextmanager
def schemas():
    assert sha256(ARCHIVE.read_bytes()).hexdigest() == EXPECTED
    with TemporaryDirectory(prefix="doffi-xlsx-schemas-") as directory:
        path = Path(directory)
        with ZipFile(ARCHIVE) as archive:
            for name in archive.namelist():
                assert "/" not in name and "\\" not in name
                (path / name).write_bytes(archive.read(name))
        parser = E.XMLParser(resolve_entities=False, no_network=True)
        yield {namespace: E.XMLSchema(E.parse(str(path / (name + ".xsd")), parser))
               for namespace, name in [(S, "sml"), (A, "dml-main"), (D, "dml-spreadsheetDrawing"), (C, "dml-chart")]}


def check_threads(roots):
    people, messages, mentions = {}, {}, {}
    for root in roots.values():
        if E.QName(root).namespace != TC: continue
        assert E.QName(root).localname in ["personList", "ThreadedComments"]
        for item in root:
            identifier = item.get("id", "")
            assert GUID.fullmatch(identifier), "Invalid thread/person identity"
            if E.QName(root).localname == "personList":
                assert item.tag == "{" + TC + "}person" and identifier not in people
                assert item.get("displayName") is not None
                assert not list(item), "Fixture profile has no person extensions"
                people[identifier] = item
            else:
                assert item.tag == "{" + TC + "}threadedComment" and identifier not in messages
                if item.get("ref") is not None: coordinate(item.get("ref"))
                assert re.fullmatch(r"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d+)?(?:Z|[+-]\d\d:\d\d)", item.get("dT", ""))
                assert item.get("done", "0") in ["0", "1", "false", "true"]
                assert [E.QName(child).localname for child in item] in [[], ['text'], ['mentions'], ['text', 'mentions']]
                for child in item:
                    assert E.QName(child).namespace == TC
                    if E.QName(child).localname != 'mentions': continue
                    assert not child.attrib
                    for mention in child:
                        assert mention.tag == '{' + TC + '}mention' and not list(mention)
                        assert set(mention.attrib) == {'mentionpersonId', 'mentionId', 'startIndex', 'length'}
                        mention_id = mention.get('mentionId'); assert GUID.fullmatch(mention_id) and mention_id not in mentions
                        assert GUID.fullmatch(mention.get('mentionpersonId', ''))
                        for field in ['startIndex', 'length']:
                            assert re.fullmatch(r'[+]?[0-9]+', mention.get(field, '')) and int(mention.get(field)) <= 4294967295
                        mentions[mention_id] = mention
                messages[identifier] = item
    for mention in mentions.values(): assert mention.get('mentionpersonId') in people, 'Unknown mentioned person'
    for identifier, item in messages.items():
        assert item.get("personId") in people, "Unknown comment author"
        seen = {identifier}
        current = item
        while current.get("parentId"):
            parent = current.get("parentId")
            assert parent in messages and parent not in seen, "Invalid thread parent"
            seen.add(parent); current = messages[parent]
            # CT_ThreadedComment ref is optional. Conflicting native locations
            # remain schema-valid source facts for explicit SDK recovery.


def audit(file, standards):
    with ZipFile(file) as archive:
        assert len(archive.namelist()) == len(set(archive.namelist())), "Duplicate ZIP entry"
        parts = {name: archive.read(name) for name in archive.namelist()}
    roots = {}
    parser = E.XMLParser(resolve_entities=False, no_network=True)
    for name, data in parts.items():
        if not name.endswith((".xml", ".rels", ".vml")): continue
        assert b"<!DOCTYPE" not in data and b"<!ENTITY" not in data, "DTD/entity declaration prohibited"
        root = E.fromstring(data, parser); roots[name] = root
        namespace = E.QName(root).namespace
        if namespace in standards:
            # The official electronic insert gives sml:t/a:t simple types, while
            # Office's documented shared strings use xml:space. Check that XML
            # attribute separately. Selected revision UIDs are also checked as
            # MS-XLSX GUIDs before removing only those known extension attrs
            # from the schema projection. Source roots remain unchanged.
            projection = root
            spaces = [n for n in root.iter() if n.get("{http://www.w3.org/XML/1998/namespace}space") is not None]
            revision_attrs = []
            for node in root.iter():
                if node.get('{' + XR + '}uid') is not None:
                    assert node.tag == '{' + S + '}comment' and GUID.fullmatch(node.get('{' + XR + '}uid')), 'Invalid comment revision UID'
                    revision_attrs.append((node, '{' + XR + '}uid'))
                if node.get('{' + MC + '}Ignorable') is not None:
                    assert node.tag in ['{' + S + '}comments', '{' + S + '}comment'], 'Unexpected compatibility declaration'
                    prefixes = re.split(r'[\t\r\n ]+', node.get('{' + MC + '}Ignorable').strip(' \t\r\n'))
                    assert all(prefix and node.nsmap.get(prefix) == XR for prefix in prefixes), 'Unknown ignorable namespace'
                    revision_attrs.append((node, '{' + MC + '}Ignorable'))
            if spaces:
                for node in spaces:
                    assert node.tag in ["{" + S + "}t", "{" + A + "}t"] and node.get("{http://www.w3.org/XML/1998/namespace}space") in ["preserve", "default"], "Unexpected xml:space"
            if spaces or revision_attrs:
                projection = deepcopy(root)
                for node in projection.iter():
                    node.attrib.pop("{http://www.w3.org/XML/1998/namespace}space", None)
                    node.attrib.pop('{' + XR + '}uid', None); node.attrib.pop('{' + MC + '}Ignorable', None)
            try: standards[namespace].assertValid(projection)
            except E.DocumentInvalid as error: raise AssertionError(str(file) + " / " + name + ": " + str(error)) from error
    types = roots["[Content_Types].xml"]
    assert types.tag == "{" + CT + "}Types"
    overrides = {unquote(n.get("PartName", "").lstrip("/")): n.get("ContentType") for n in types if n.tag == "{" + CT + "}Override"}
    defaults = {n.get("Extension"): n.get("ContentType") for n in types if n.tag == "{" + CT + "}Default"}
    for name in parts:
        assert name == "[Content_Types].xml" or overrides.get(name) or defaults.get(name.rsplit(".", 1)[-1]), "Missing content type"
    for name in overrides: assert name in parts, "Content type names missing part"
    relations = {}
    for name, root in roots.items():
        if not name.endswith(".rels"): continue
        assert root.tag == "{" + P + "}Relationships"
        source = "" if name == "_rels/.rels" else posixpath.join(posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5])
        assert not source or source in parts, "Missing relationship owner"
        found = {}; relations[source] = found
        for relation in root:
            identifier = relation.get("Id"); assert identifier and identifier not in found
            target = unquote(relation.get("Target"))
            if relation.get("TargetMode") == "External":
                assert relation.get("Type") == R + "/hyperlink", "Unexpected external resource"
            else:
                target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(source), target))
                assert target in parts, "Missing relationship target: " + target
            found[identifier] = (relation.get("Type"), target)
    for name, root in roots.items():
        for node in root.iter():
            for key, value in node.attrib.items():
                if key.startswith("{" + R + "}"):
                    assert value in relations.get(name, {}), "Unresolved explicit relationship"
    styles = next((root for root in roots.values() if root.tag == "{" + S + "}styleSheet"), None)
    style_count = len(styles.find("{" + S + "}cellXfs")) if styles is not None else 1
    strings = next((root for root in roots.values() if root.tag == "{" + S + "}sst"), [])
    for root in roots.values():
        if root.tag == "{" + S + "}worksheet":
            seen = set()
            for row in root.findall("{" + S + "}sheetData/{" + S + "}row"):
                markers = list(row.iter("{" + ROW_VISIBILITY + "}filterHidden"))
                if markers:
                    assert len(markers) == 1 and row.get("hidden") in {"1", "true"}, "Invalid filter-only row provenance"
                    marker = markers[0]
                    extension = marker.getparent()
                    assert extension.tag == "{" + S + "}ext" and extension.get("uri") == ROW_VISIBILITY
                    assert not marker.attrib and not len(marker) and not (marker.text or "").strip()
                for cell in row.findall("{" + S + "}c"):
                    address = cell.get("r"); r, _ = coordinate(address)
                    assert r == int(row.get("r")) and address not in seen, "Repeated or misplaced cell"
                    seen.add(address)
                    assert 0 <= int(cell.get("s", "0")) < style_count, "Invalid style index"
                    if cell.get("t") == "s":
                        assert 0 <= int(cell.findtext("{" + S + "}v")) < len(strings), "Invalid shared string index"
        if root.tag == "{" + D + "}wsDr":
            ids = [n.get("id") for n in root.iter("{" + D + "}cNvPr")]
            assert len(ids) == len(set(ids)), "Duplicate drawing identity"
    check_threads(roots)
    return roots, relations


def check(directory):
    records = []
    with schemas() as standards:
        for file in sorted(directory.glob("*.xlsx")):
            roots, _ = audit(file, standards)
            records.append({"file": file.name, "standardParts": sum(E.QName(root).namespace in standards for root in roots.values()), "modernParts": sum(E.QName(root).namespace == TC for root in roots.values()),
                            "xmlSpaceAttributesChecked": sum(node.get("{http://www.w3.org/XML/1998/namespace}space") is not None for root in roots.values() for node in root.iter())})
    assert records, "No XLSX inputs found"
    print(json.dumps({"schemasSha256": EXPECTED, "lxml": E.LXML_VERSION,
                      "profile": "Unmodified official Transitional XSD; sml:t/a:t xml:space and selected MS-XLSX comment revision UID/compatibility declarations checked separately and removed only from validation projection. OPC references and selected MS-XLSX threads checked separately. VML is checked for XML well-formedness and relationship integrity, not its full schema.",
                      "cases": records}, indent=2))


if __name__ == "__main__": check(Path(sys.argv[1]))
