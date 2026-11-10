"""
Recalculate an .xlsx with LibreOffice (headless) so formula results get stored in the file,
then report formula errors as JSON. Used by tests/test_deals.py and the GitHub Actions workflow.

Usage: python tests/recalc.py model.xlsx [timeout_seconds]
Needs LibreOffice:  sudo apt-get install -y libreoffice-calc
"""
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

ERRORS = ("#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!")

# Tell LibreOffice to always recalculate Excel files when it opens them
XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>"""


def recalc(path, timeout=120):
    path = Path(path).resolve()
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return {"error": "LibreOffice (soffice) not found"}
    with tempfile.TemporaryDirectory() as tmp:
        profile = Path(tmp) / "profile"
        (profile / "user").mkdir(parents=True)
        (profile / "user" / "registrymodifications.xcu").write_text(XCU)
        out = Path(tmp) / "out"
        cmd = [soffice, f"-env:UserInstallation=file://{profile}", "--headless", "--norestore",
               "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out), str(path)]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        produced = out / path.name
        if not produced.exists():
            return {"error": "conversion failed", "stderr": r.stderr[-500:]}
        shutil.copy(produced, path)

    formulas = load_workbook(path)
    values = load_workbook(path, data_only=True)
    total, errors, empty = 0, [], 0
    for ws in formulas:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    total += 1
                    v = values[ws.title][c.coordinate].value
                    if isinstance(v, str) and v.startswith(ERRORS):
                        errors.append(f"{ws.title}!{c.coordinate} {v}")
                    if v is None:
                        empty += 1
    return {"status": "success" if not errors else "errors_found", "total_formulas": total,
            "total_errors": len(errors), "errors": errors[:50], "formulas_without_value": empty}


if __name__ == "__main__":
    print(json.dumps(recalc(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 120), indent=2))
