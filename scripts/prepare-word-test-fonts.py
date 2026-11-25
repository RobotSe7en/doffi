"""Pack local Microsoft Office fonts as same-origin WOFF assets for internal Playground builds."""

import argparse
from hashlib import sha256
import json
from pathlib import Path
import struct
import zlib

try:
    from fontTools import version as fonttools_version
    from fontTools.ttLib import TTFont
except ImportError:
    fonttools_version = None
    TTFont = None


FACES = [
    # Core Chinese Office faces
    ("simsun", "Simsun.ttc", 0, ["宋体", "华文宋体", "STSong"]),
    ("nsimsun", "Simsun.ttc", 1, ["新宋体"]),
    ("simhei", "SimHei.ttf", 0, ["黑体", "华文黑体", "STHeiti"]),
    ("fangsong", "Fangsong.ttf", 0, ["仿宋", "仿宋_GB2312", "FangSong_GB2312", "华文仿宋", "STFangsong"]),
    ("kaiti", "Kaiti.ttf", 0, ["楷体", "楷体_GB2312", "KaiTi_GB2312", "SimKai", "华文楷体", "STKaiti"]),
    ("dengxian", "Deng.ttf", 0, ["等线"]),
    ("dengxian-bold", "Dengb.ttf", 0, ["等线"]),
    ("dengxian-light", "Dengl.ttf", 0, ["等线 Light"]),
    ("yahei", "msyh.ttc", 0, ["微软雅黑"]),
    ("yahei-bold", "msyhbd.ttc", 0, ["微软雅黑"]),
    ("yahei-light", "msyhl.ttc", 0, ["微软雅黑 Light"]),
    ("yahei-ui", "msyh.ttc", 1, ["微软雅黑 UI"]),
    ("yahei-ui-bold", "msyhbd.ttc", 1, ["微软雅黑 UI"]),
    ("yahei-ui-light", "msyhl.ttc", 1, ["微软雅黑 UI Light"]),
    ("jhenghei", "MSJH.ttf", 0, ["微軟正黑體", "微软正黑体"]),
    ("jhenghei-bold", "MSJHBD.ttf", 0, ["微軟正黑體", "微软正黑体"]),
    ("stzhongsong", "STZHONGS.ttf", 0, ["华文中宋"]),
    ("stxinwei", "STXINWEI.ttf", 0, ["华文新魏"]),
    ("stxingkai", "STXINGKA.ttf", 0, ["华文行楷"]),
    ("stliti", "STLITI.ttf", 0, ["华文隶书"]),
    ("sthupo", "STHUPO.ttf", 0, ["华文琥珀"]),
    # Core Latin & symbol Office faces
    ("times-new-roman", "times.ttf", 0, []),
    ("times-new-roman-bold", "timesbd.ttf", 0, []),
    ("times-new-roman-italic", "timesi.ttf", 0, []),
    ("times-new-roman-bold-italic", "timesbi.ttf", 0, []),
    ("aptos", "Aptos.ttf", 0, []),
    ("aptos-bold", "Aptos-Bold.ttf", 0, []),
    ("aptos-italic", "Aptos-Italic.ttf", 0, []),
    ("aptos-bold-italic", "Aptos-Bold-Italic.ttf", 0, []),
    ("aptos-light", "Aptos-Light.ttf", 0, []),
    ("aptos-semibold", "Aptos-SemiBold.ttf", 0, []),
    ("calibri", "Calibri.ttf", 0, []),
    ("calibri-bold", "Calibrib.ttf", 0, []),
    ("calibri-italic", "Calibrii.ttf", 0, []),
    ("calibri-bold-italic", "Calibriz.ttf", 0, []),
    ("calibri-light", "calibril.ttf", 0, []),
    ("calibri-light-italic", "calibrili.ttf", 0, []),
    ("cambria", "Cambria.ttc", 0, []),
    ("cambria-bold", "Cambriab.ttf", 0, []),
    ("cambria-italic", "Cambriai.ttf", 0, []),
    ("cambria-bold-italic", "Cambriaz.ttf", 0, []),
    ("arial", "arial.ttf", 0, []),
    ("arial-bold", "arialbd.ttf", 0, []),
    ("arial-italic", "ariali.ttf", 0, []),
    ("arial-bold-italic", "arialbi.ttf", 0, []),
    ("arial-black", "ariblk.ttf", 0, []),
    ("consolas", "Consola.ttf", 0, []),
    ("consolas-bold", "Consolab.ttf", 0, []),
    ("consolas-italic", "Consolai.ttf", 0, []),
    ("consolas-bold-italic", "Consolaz.ttf", 0, []),
    ("tahoma", "tahoma.ttf", 0, []),
    ("tahoma-bold", "tahomabd.ttf", 0, []),
    ("verdana", "Verdana.ttf", 0, []),
    ("verdana-bold", "Verdana Bold.ttf", 0, []),
    ("verdana-italic", "Verdana Italic.ttf", 0, []),
    ("verdana-bold-italic", "Verdana Bold Italic.ttf", 0, []),
    ("symbol", "symbol.ttf", 0, []),
    ("wingdings", "Wingdings.ttf", 0, []),
    ("wingdings-2", "Wingdings 2.ttf", 0, []),
    ("wingdings-3", "Wingdings 3.ttf", 0, []),
]


def _sfnt_checksum(data: bytes) -> int:
    pad = (-len(data)) & 3
    if pad:
        data += b"\x00" * pad
    total = 0
    for (val,) in struct.iter_unpack(">I", data):
        total = (total + val) & 0xFFFFFFFF
    return total


def _convert_stdlib_woff(source: Path, face_index: int, target: Path) -> dict:
    raw = source.read_bytes()
    if raw[:4] == b"ttcf":
        face_offset = struct.unpack_from(">I", raw, 12 + 4 * face_index)[0]
    else:
        face_offset = 0
    sfnt_version, raw_num_tables = struct.unpack_from(">4sH", raw, face_offset)
    entries = []
    for i in range(raw_num_tables):
        tag, checksum, offset, length = struct.unpack_from(">4sIII", raw, face_offset + 12 + 16 * i)
        if length == 0:
            continue
        entries.append((tag, checksum, raw[offset : offset + length]))
    entries.sort(key=lambda item: item[0])
    num_tables = len(entries)

    # Extract OpenType metadata from OS/2, head, and name tables.
    by_tag = {tag: data for tag, _, data in entries}
    weight = struct.unpack_from(">H", by_tag[b"OS/2"], 4)[0]
    head = bytearray(by_tag[b"head"])
    font_rev = struct.unpack_from(">I", head, 4)[0]
    major_ver = (font_rev >> 16) & 0xFFFF
    minor_ver = font_rev & 0xFFFF
    italic = bool(struct.unpack_from(">H", head, 44)[0] & 2)

    name_table = by_tag[b"name"]
    _, count, str_off = struct.unpack_from(">HHH", name_table, 0)
    names = {}
    for i in range(count):
        pid, eid, lid, nid, length, soff = struct.unpack_from(">HHHHHH", name_table, 6 + 12 * i)
        chunk = name_table[str_off + soff : str_off + soff + length]
        if pid == 3 and eid in (1, 10) and lid == 0x0409:
            names[nid] = chunk.decode("utf-16-be", errors="replace")
        elif nid not in names and pid == 1 and eid == 0 and lid == 0:
            names[nid] = chunk.decode("mac_roman", errors="replace")

    # Recompute head.checkSumAdjustment so extracted TTC faces pass OTS checksum checks.
    struct.pack_into(">I", head, 8, 0)
    rebuilt = []
    for tag, _, data in entries:
        tdata = bytes(head) if tag == b"head" else data
        rebuilt.append((tag, _sfnt_checksum(tdata), tdata))
    entry_selector = (num_tables).bit_length() - 1
    search_range = (1 << entry_selector) * 16
    range_shift = num_tables * 16 - search_range
    sfnt_header = struct.pack(">4sHHHH", sfnt_version, num_tables, search_range, entry_selector, range_shift)
    sfnt_offset = 12 + 16 * num_tables
    sfnt_dir = bytearray()
    total_sum = _sfnt_checksum(sfnt_header)
    for tag, csum, tdata in rebuilt:
        sfnt_dir.extend(struct.pack(">4sIII", tag, csum, sfnt_offset, len(tdata)))
        total_sum = (total_sum + csum) & 0xFFFFFFFF
        sfnt_offset += (len(tdata) + 3) & ~3
    total_sum = (total_sum + _sfnt_checksum(bytes(sfnt_dir))) & 0xFFFFFFFF
    struct.pack_into(">I", head, 8, (0xB1B0AFBA - total_sum) & 0xFFFFFFFF)
    head_bytes = bytes(head)
    head_csum = _sfnt_checksum(head_bytes)

    woff_entries = []
    woff_body = bytearray()
    data_offset = 44 + 20 * num_tables
    for tag, csum, tdata in rebuilt:
        if tag == b"head":
            tdata = head_bytes
            csum = head_csum
        comp = zlib.compress(tdata, 9)
        stored = comp if len(comp) < len(tdata) else tdata
        woff_entries.append((tag, data_offset, len(stored), len(tdata), csum))
        woff_body.extend(stored)
        pad = (-len(stored)) & 3
        if pad:
            woff_body.extend(b"\x00" * pad)
        data_offset += len(stored) + pad

    woff_length = data_offset
    woff_header = struct.pack(
        ">4s4sIHHIHHIIIII",
        b"wOFF", sfnt_version, woff_length, num_tables, 0, sfnt_offset,
        major_ver, minor_ver, 0, 0, 0, 0, 0,
    )
    woff_dir = b"".join(struct.pack(">4sIIII", *entry) for entry in woff_entries)
    target.write_bytes(woff_header + woff_dir + bytes(woff_body))
    return {
        "family": names.get(1),
        "weight": str(weight),
        "style": "italic" if italic else "normal",
        "version": names.get(5),
        "license": names.get(13),
        "licenseUrl": names.get(14),
    }


def _convert_face(source: Path, face_index: int, target: Path) -> dict:
    if TTFont is not None:
        with TTFont(source, fontNumber=face_index, recalcTimestamp=False) as font:
            names = font["name"]
            font.flavor = "woff"
            font.save(target)
            return {
                "family": names.getDebugName(1),
                "weight": str(font["OS/2"].usWeightClass),
                "style": "italic" if font["head"].macStyle & 2 else "normal",
                "version": names.getDebugName(5),
                "license": names.getDebugName(13),
                "licenseUrl": names.getDebugName(14),
            }
    return _convert_stdlib_woff(source, face_index, target)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--font-dir", type=Path,
        default=Path("/Applications/Microsoft Word.app/Contents/Resources/DFonts"),
    )
    args = parser.parse_args()
    output = Path(__file__).resolve().parents[1] / "fonts/local"
    output.mkdir(parents=True, exist_ok=True)
    sources = []
    for key, filename, index, aliases in FACES:
        source = args.font_dir / filename
        target = output / (key + ".woff")
        meta = _convert_face(source, index, target)
        sources.append({
            "family": meta["family"],
            "aliases": aliases,
            "weight": meta["weight"],
            "style": meta["style"],
            "file": target.name,
            "bytes": target.stat().st_size,
            "sha256": sha256(target.read_bytes()).hexdigest(),
            "sourceFile": filename,
            "sourceFace": index,
            "sourceSha256": sha256(source.read_bytes()).hexdigest(),
            "version": meta["version"],
            "license": meta["license"],
            "licenseUrl": meta["licenseUrl"],
        })
    conversion = (
        f"FontTools {fonttools_version}, full WOFF faces"
        if fonttools_version
        else "Python stdlib WOFF 1.0, full WOFF faces"
    )
    manifest = {"source": "Local Microsoft Office font installation", "conversion": conversion, "fonts": sources}
    (output / "sources.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")
    print(f"Prepared {len(sources)} internal Office font faces.")


if __name__ == "__main__":
    main()
