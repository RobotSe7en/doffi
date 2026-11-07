"""Record independent native function observations with pinned Calc, never SDK output."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import argparse
import importlib.util
import json
import os

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("calculation_scalar_source", ROOT / "scripts/check-xlsx-scalar-semantics.py")
probe = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probe)
FIXTURE = ROOT / "packages/office/tests/fixtures/workbook-calculation.xlsx"
REPORT = ROOT / "docs/evidence/workbook-calculation-functions.json"
FORMULAS = [
    '_xlfn.IFS(FALSE,1,TRUE,2)', '_xlfn.IFS(FALSE,1)', 'LEFT("中文ABC",2)', 'RIGHT("中文ABC",2)',
    'MID("中文ABC",2,3)', 'RIGHT("abc",0)', 'LEFT("abc",-1)', 'TRIM("  a   b  ")', 'UPPER("aBc中国")', 'LOWER("aBc中国")',
    '_xlfn.CONCAT("a",1,TRUE)', 'CONCATENATE("a","b",3)', '_xlfn.TEXTJOIN("-",TRUE,"a","","b")', '_xlfn.TEXTJOIN("-",FALSE,"a","","b")',
    '_xlfn.TEXTJOIN("-",TRUE,A1:A5)', '_xlfn.CONCAT(A1:A5)', 'SUBSTITUTE("a-a-a","a","x",2)', 'SUBSTITUTE("a-a-a","a","x")',
    'REPLACE("abcdef",2,3,"X")', 'FIND("B","aBcB",3)', 'SEARCH("b?","aBcB")', 'SEARCH("~*","a*b")', 'EXACT("ABC","abc")',
    'TEXT(1234.5,"0.00")', 'TEXT("123","0.00")', 'TEXT("abc","@")', 'VALUE("1234.50")', 'VALUE("15%")', 'VALUE("2024-02-29")',
    'DATE(2024,2,29)', 'DATE(1900,2,29)', 'DATE(1900,3,0)', 'YEAR(60)', 'MONTH(60)', 'DAY(60)', 'DAY(0)',
    'TIME(12,30,0)', 'HOUR(0.5)', 'MINUTE(0.5)', 'SECOND(0.5)', 'HOUR("12:30:05")', 'MINUTE("12:30:05")', 'SECOND("12:30:05")',
    '_xlfn.DAYS(DATE(2024,3,1),DATE(2024,2,28))', '_xlfn.DAYS("2024-03-01","2024-02-28")', 'EDATE(DATE(2024,1,31),1)', 'EOMONTH(DATE(2024,1,1),1)',
    'WEEKDAY(DATE(2024,1,1),2)', 'WEEKDAY(1)', 'DATE(-1,1,1)',
    'SUMIF(C1:C2,">2")', 'SUMIF(C1:C2,">2",C1)', 'COUNTIF(C1:C2,">2")', 'AVERAGEIF(C1:C2,">2")',
    'SUMIFS(C1:C2,C1:C2,">2",C1:C2,"<5")', 'COUNTIFS(C1:C2,">2",C1:C2,"<5")', 'AVERAGEIFS(C1:C2,C1:C2,">2",C1:C2,"<5")',
    'COUNTIF(A1:A8,"")', 'COUNTIF(A1:A8,0)', 'COUNTIF(A1:A8,"=0")', 'COUNTIF(A1:A8,"<>0")', 'COUNTIF(A1:A8,"*")',
    'COUNTIF(A1:A8,"2")', 'COUNTIFS(A1:A8,"")', 'COUNTIFS(A1:A8,0)', 'SUMIF(A1:A8,0,C1:C8)',
    'MATCH(3,C1:C2,0)', 'MATCH(4,C1:C2,1)', 'MATCH(4,C1:C2,0)', 'VLOOKUP(3,C1:C2,1,FALSE)',
    'HLOOKUP(3,C1:C2,2,FALSE)', 'INDEX(C1:C2,2)', 'SUM(INDEX(C1:C2,0))', 'ROW(INDEX(C1:C2,2))', 'SUBTOTAL(9,C1:C2)', 'SUBTOTAL(109,C1:C2)',
    'COUNTIF(A6,"#N/A")', 'COUNTIFS(A6,"#N/A")', 'COUNTIF(A6,"<>#N/A")', 'SUMIF(A6,"#N/A",C1)', 'AVERAGEIF(A6,"#N/A",C1)',
]
# Excel compatibility is an existing user decision. Calc's Gregorian results
# remain recorded observations and are never rewritten into Excel's calendar.
EXCEL_CALENDAR = {
    'DATE(1900,2,29)': {"type": "number", "value": 60}, 'DATE(1900,3,0)': {"type": "number", "value": 60},
    'YEAR(60)': {"type": "number", "value": 1900}, 'MONTH(60)': {"type": "number", "value": 2}, 'DAY(60)': {"type": "number", "value": 29},
    'DAY(0)': {"type": "number", "value": 0}, 'WEEKDAY(1)': {"type": "number", "value": 1},
}
DECLARED_DIFFERENCES = {
    '_xlfn.CONCAT("a",1,TRUE)': ({"type": "text", "value": "a1TRUE"}, "Existing declared logical-to-text conversion is TRUE/FALSE, shared with concatenation; see prior xlsx-scalar-semantics evidence"),
    '_xlfn.TEXTJOIN("-",TRUE,A1:A5)': ({"type": "text", "value": "2-2-TRUE"}, "Existing declared logical-to-text conversion is TRUE/FALSE, shared with concatenation"),
    '_xlfn.CONCAT(A1:A5)': ({"type": "text", "value": "22TRUE"}, "Existing declared logical-to-text conversion is TRUE/FALSE, shared with concatenation"),
    'DATE(-1,1,1)': ({"type": "error", "value": "#NUM!"}, "Microsoft DATE documentation: negative year returns #NUM!"),
    'COUNTIF(A1:A8,"<>0")': ({"type": "number", "value": 6}, "Microsoft Excel.Functions.countIf range describes nonblank cells; actual blank is not a nonzero stored value; Excel/WPS runtime unverified"),
    'COUNTIF(A1:A8,"*")': ({"type": "number", "value": 3}, "Microsoft COUNTIF wildcard example selects text-containing cells; numeric, logical, error and actual blank cells remain distinct"),
    'COUNTIFS(A1:A8,0)': ({"type": "number", "value": 2}, "ECMA-376 Part 1 18.17.7.56 explicitly treats an empty cell in any argument as zero; Calc observation retained; Excel/WPS runtime unverified"),
}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--verify-native", action="store_true")
    args = parser.parse_args()
    probe.FORMULAS = FORMULAS
    probe.STANDARD_DIFFERENCES = {formula: (value, "User-selected Excel 1900 compatibility calendar; Calc observation retained separately") for formula, value in EXCEL_CALENDAR.items()}
    probe.STANDARD_DIFFERENCES.update(DECLARED_DIFFERENCES)
    data = probe.generate(cached=True)
    if not args.record and not args.verify_native:
        assert FIXTURE.read_bytes() == data, "Independent calculation source changed"
        report = json.loads(REPORT.read_text())
        assert report["sourceSha256"] == sha256(data).hexdigest()
        assert [case["expression"] for case in report["cases"]] == ["=" + formula for formula in FORMULAS]
        print(f"Verified {len(FORMULAS)} independent calculation source cases")
        return
    native = probe.native
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    assert native.shared.run(["libreoffice", "--version"]) == native.BUILD
    directory = ROOT / ".cache/workbook-calculation-native"
    directory.mkdir(exist_ok=True)
    reference = probe.generate(cached=False)
    with TemporaryDirectory(dir=directory) as temporary:
        work = Path(temporary)
        output = work / "out"
        output.mkdir()
        source = work / "functions.xlsx"
        source.write_bytes(reference)
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(output), str(source)])
        cases = probe.observe(output / "functions.xlsx")
    report = {"scope": "Additional selected calculation families from independent ZIP/XML sources and pinned Calc. Source caches are deliberately wrong. Explicit host time is tested separately; Excel/WPS runtime remains unverified.",
              "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(),
              "referenceInputSha256": sha256(reference).hexdigest(), "cases": cases}
    if args.record:
        FIXTURE.write_bytes(data)
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert FIXTURE.read_bytes() == data, "Independent calculation source changed"
        assert json.loads(REPORT.read_text()) == report, "Native calculation observations changed"
    print(f"Verified {len(cases)} independent native calculation observations")


if __name__ == "__main__":
    main()
