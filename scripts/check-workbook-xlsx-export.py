"""Check SDK XLSX output against Transitional XSD and pinned actual Calc edits."""
import importlib.util
from pathlib import Path
import subprocess
from tempfile import TemporaryDirectory

from workbook_acceptance import audit_schema

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("workbook_axis", ROOT / "scripts/check-workbook-axis.py")
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)


def edit_table(document, _property):
    sheet = document.getSheets().getByIndex(0)
    assert sheet.getCellRangeByName("B2").getValue() == 6
    assert sheet.getCellRangeByName("H1").getValue() == 12
    sheet.getCellRangeByName("B2").setValue(7)
    sheet.getCellRangeByName("H1").setFormula("=B2*4")
    ranges = document.DatabaseRanges
    table = ranges.getByName("SalesTable")
    address = table.getDataArea()
    address.EndRow += 1
    table.setDataArea(address)


def edit_chart(document, _property):
    sheet = document.getSheets().getByIndex(0)
    assert sheet.getCellRangeByName("B2").getValue() == 39
    charts = sheet.getCharts()
    assert len(charts.getElementNames()) == 3
    chart = charts.getByName(charts.getElementNames()[0]).getEmbeddedObject()
    chart.HasMainTitle = True
    title = chart.getTitleObject()
    first = title.getText()[0]
    first.setString("Calc title edit")
    title.setText((first,))


def edit_note(document, _property):
    sheet = document.getSheets().getByIndex(0)
    note = sheet.getCellRangeByName("B2").getAnnotation()
    assert sheet.getAnnotations().getCount() == 1
    assert note.getAuthor() == "独立备注作者"
    note.setString("Calc note edit 禰󠄀")


def main():
    directory = ROOT / ".cache/ci-office/workbook/xlsx-export"
    directory.mkdir(parents=True, exist_ok=True)
    command = ["node", str(ROOT / "scripts/check-xlsx-export.mjs")]
    with TemporaryDirectory(prefix="doffi-xlsx-export-") as temporary:
        output = Path(temporary) / "editable"
        native = Path(temporary) / "calc"
        native.mkdir()
        subprocess.run(command + ["export", str(output)], cwd=ROOT, check=True, timeout=300)
        with (directory / "export-schemas.json").open("w") as report:
            audit_schema(output, stdout=report)
        edits = [("05-native-table.xlsx", edit_table), ("07-three-charts.xlsx", edit_chart), ("09-classic-notes.xlsx", edit_note)]
        for name, edit in edits:
            axis.native(output / name, native / name, {}, mutate=edit)
        with (directory / "calc-schemas.json").open("w") as report:
            audit_schema(native, stdout=report)
        subprocess.run(command + ["reopen", str(native)], cwd=ROOT, check=True, timeout=90)



if __name__ == "__main__":
    main()
