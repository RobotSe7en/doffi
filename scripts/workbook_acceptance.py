"""Development-only native acceptance environment and independent XML audit."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import atexit
import json
import os
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parent.parent


def pin_native_fonts(sources, reference_fonts, gate_name):
    """Isolate source-declared fonts and emoji fallback at original reference bytes.

    The historical reference recorded the author's entire OS font inventory.
    Unrelated installed fonts must neither be prerequisites nor affect fallback.
    Keep inherited Fontconfig rendering/alias rules, but replace all search dirs.
    """
    families = {"Noto Color Emoji"}
    inputs = []
    for source in sources:
        inputs.append({"source": str(source.relative_to(ROOT)), "sha256": sha256(source.read_bytes()).hexdigest()})
        with ZipFile(source) as archive:
            for name in archive.namelist():
                if not name.endswith((".xml", ".vml")):
                    continue
                for node in ET.fromstring(archive.read(name)).iter():
                    for key, value in node.attrib.items():
                        if key in {"typeface", "face"} or (key == "val" and node.tag.rsplit("}", 1)[-1] in {"name", "rFont"}):
                            if value:
                                families.add(value)
                    for match in re.finditer(r"(?:^|;)\s*font-family\s*:\s*([^;]+)", node.get("style", ""), re.I):
                        families.update(value.strip().strip("\"'") for value in match[1].split(","))

    reference = {(item["file"], item["sha256"]) for item in reference_fonts}
    selected = {}
    resolutions = []
    for family in sorted(families):
        for style in ["Regular", "Bold", "Italic", "Bold Italic"]:
            pattern = family + ":style=" + style
            output = subprocess.check_output(["fc-match", "-f", "%{family}\n%{file}", pattern], text=True).splitlines()
            if len(output) != 2 or family not in [name.strip() for name in output[0].split(",")]:
                raise RuntimeError("Required native source font unavailable: " + pattern)
            path = Path(output[1])
            digest = sha256(path.read_bytes()).hexdigest()
            if (path.name, digest) not in reference:
                raise RuntimeError("Native source font differs from the original reference bytes: " + str(path))
            if path.name in selected and selected[path.name][1] != digest:
                raise RuntimeError("Conflicting native font basenames: " + path.name)
            selected[path.name] = (path, digest)
            resolutions.append({"pattern": pattern, "file": path.name, "sha256": digest})

    temporary = TemporaryDirectory(prefix="doffi-native-fonts-")
    atexit.register(temporary.cleanup)
    directory = Path(temporary.name)
    fonts = directory / "fonts"
    fonts.mkdir()
    for name, (path, _) in selected.items():
        shutil.copyfile(path, fonts / name)
    inherited = Path(os.environ.get("FONTCONFIG_FILE", "/etc/fonts/fonts.conf")).resolve()
    config = directory / "fonts.conf"
    config.write_text('<?xml version="1.0"?>\n<!DOCTYPE fontconfig SYSTEM "urn:fontconfig:fonts.dtd">\n'
                      '<fontconfig><include>' + escape(str(inherited)) + '</include><reset-dirs/>'
                      '<dir>' + escape(str(fonts)) + '</dir><cachedir>' + escape(str(directory / "cache")) +
                      '</cachedir></fontconfig>\n')
    os.environ["FONTCONFIG_FILE"] = str(config)
    subprocess.run(["fc-cache", "--force", str(fonts)], check=True, timeout=60)
    expected = sorted([{"file": name, "sha256": digest} for name, (_, digest) in selected.items()],
                      key=lambda item: (item["file"], item["sha256"]))
    active = []
    for file in sorted(set(subprocess.check_output(["fc-list", "-f", "%{file}\n"], text=True).splitlines())):
        path = Path(file)
        if path.parent != fonts:
            raise RuntimeError("Ambient font escaped native font isolation: " + file)
        active.append({"file": path.name, "sha256": sha256(path.read_bytes()).hexdigest()})
    active.sort(key=lambda item: (item["file"], item["sha256"]))
    if active != expected:
        raise RuntimeError("The complete isolated native font inventory differs.")
    report = ROOT / ".cache/ci-office/workbook/font-gates" / (gate_name + ".json")
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(json.dumps({"sources": inputs, "fallback": "Noto Color Emoji", "resolutions": resolutions,
                                  "activeFonts": active, "inheritedConfig": str(inherited),
                                  "inheritedConfigSha256": sha256(inherited.read_bytes()).hexdigest()}, indent=2) + "\n")
    return expected


def audit_schema(directory, *, capture_output=False, stdout=None, text=True):
    environment = {key: value for key, value in os.environ.items() if key not in {
        "PYTHONHOME", "PYTHONPATH", "URE_BOOTSTRAP", "UNO_PATH", "UNO_TYPES", "UNO_SERVICES"}}
    candidates = [ROOT / ".cache/xlsx-venv/bin/python", Path("/usr/bin/python3"), Path(sys.executable)]
    for python in dict.fromkeys(candidates):
        if not python.is_file():
            continue
        probe = subprocess.run([str(python), "-c", "from lxml import etree"], env=environment,
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=15)
        if probe.returncode == 0:
            return subprocess.run([str(python), str(ROOT / "scripts/check-xlsx-output.py"), str(directory)],
                                  cwd=ROOT, env=environment, check=True, capture_output=capture_output, stdout=stdout, text=text)
    raise RuntimeError("Official XLSX schema validation requires a separate Python with lxml.")
