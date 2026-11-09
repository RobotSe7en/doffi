"""Independent XML audit interpreter, separate from the archive's UNO runtime."""
from pathlib import Path
import os
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent


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
