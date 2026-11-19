"""Pack local Word fonts as same-origin assets for internal Playground builds."""

import argparse
from hashlib import sha256
import json
from pathlib import Path

from fontTools import version
from fontTools.ttLib import TTFont


FACES = [
    ("simsun", "Simsun.ttc", 0, ["宋体"]),
    ("nsimsun", "Simsun.ttc", 1, ["新宋体"]),
    ("simhei", "SimHei.ttf", 0, ["黑体"]),
    ("fangsong", "Fangsong.ttf", 0, ["仿宋"]),
    ("kaiti", "Kaiti.ttf", 0, ["楷体", "SimKai"]),
    ("dengxian", "Deng.ttf", 0, ["等线"]),
    ("dengxian-bold", "Dengb.ttf", 0, ["等线"]),
    ("dengxian-light", "Dengl.ttf", 0, ["等线 Light"]),
    ("yahei", "msyh.ttc", 0, ["微软雅黑"]),
    ("yahei-bold", "msyhbd.ttc", 0, ["微软雅黑"]),
    ("yahei-ui", "msyh.ttc", 1, ["微软雅黑 UI"]),
    ("yahei-ui-bold", "msyhbd.ttc", 1, ["微软雅黑 UI"]),
    ("times-new-roman", "times.ttf", 0, []),
    ("times-new-roman-bold", "timesbd.ttf", 0, []),
    ("times-new-roman-italic", "timesi.ttf", 0, []),
    ("times-new-roman-bold-italic", "timesbi.ttf", 0, []),
]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--font-dir", type=Path,
        default=Path("/Applications/Microsoft Word.app/Contents/Resources/DFonts"),
    )
    args = parser.parse_args()
    output = Path(__file__).resolve().parents[1] / "apps/playground/src/assets/word-fonts/internal"
    output.mkdir(parents=True, exist_ok=True)
    sources = []
    for key, filename, index, aliases in FACES:
        source = args.font_dir / filename
        target = output / (key + ".woff")
        with TTFont(source, fontNumber=index, recalcTimestamp=False) as font:
            names = font["name"]
            font.flavor = "woff"
            font.save(target)
            sources.append({
                "family": names.getDebugName(1),
                "aliases": aliases,
                "weight": str(font["OS/2"].usWeightClass),
                "style": "italic" if font["head"].macStyle & 2 else "normal",
                "file": target.name,
                "bytes": target.stat().st_size,
                "sha256": sha256(target.read_bytes()).hexdigest(),
                "sourceFile": filename,
                "sourceFace": index,
                "sourceSha256": sha256(source.read_bytes()).hexdigest(),
                "version": names.getDebugName(5),
                "license": names.getDebugName(13),
                "licenseUrl": names.getDebugName(14),
            })
    manifest = {"source": "Local Microsoft Word font installation", "conversion": f"FontTools {version}, full WOFF faces", "fonts": sources}
    (output / "sources.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")
    print(f"Prepared {len(sources)} internal Word font faces.")


if __name__ == "__main__":
    main()
