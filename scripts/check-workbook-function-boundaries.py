"""Independent decimal, UTF-16 and two-calendar references; no SDK output."""
from decimal import Decimal, localcontext, ROUND_HALF_UP, ROUND_UP, ROUND_DOWN
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from io import BytesIO
import argparse
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("function_boundary_native", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/workbook-function-boundaries.json"
FIXTURES = ROOT / "packages/office/tests/fixtures"
DECIMAL_BASIS = "Independent Python Decimal quantization of literal decimals, with nearest ties away from zero / away from zero / toward zero."
UNICODE_BASIS = "Declared legacy Version 1 workbook profile without a compatibility-version extension, using Microsoft's documented Version 1/2 distinction and UTF-16 positions. Calc's scalar-character observations are retained separately; this is not an Excel runtime observation."
DATE_BASIS = "Microsoft's two date systems and the user-selected Excel 1900 compatibility calendar; Gregorian Calc differences are retained."


def cases(system):
    result = []
    def add(family, expression, value, basis="Independent declared argument/type/error expectation; native producer observation retained."):
        kind = "boolean" if isinstance(value, bool) else "number" if isinstance(value, (int, float)) else "error" if value in ["#NUM!", "#VALUE!", "#REF!", "#N/A", "#DIV/0!"] else "text"
        result.append({"family": family, "expression": "=" + expression, "expected": {"type": kind, "value": value}, "basis": basis})
    for name, direction in [("ROUND", ROUND_HALF_UP), ("ROUNDUP", ROUND_UP), ("ROUNDDOWN", ROUND_DOWN)]:
        for number in ["1.005", "-1.005", "2.675", "-2.675", "123.45", "-123.45"]:
            for digits in [2, 0, -1]:
                with localcontext() as context:
                    context.prec = 50
                    value = float(Decimal(number).quantize(Decimal(1).scaleb(-digits), rounding=direction))
                add("rounding", f"{name}({number},{digits})", 0 if value == 0 else value, DECIMAL_BASIS)
    for expression, value in [
        ('ROUND(TRUE,2)', 1), ('ROUND("1.25",1)', 1.3), ('ROUND(1.25,"1")', 1.3), ('ROUND("abc",2)', '#VALUE!'),
        ('ROUND(1/0,2)', '#DIV/0!'), ('ROUND(1.25,TRUE)', 1.3), ('ROUND(1.255,2.9)', 1.26), ('ROUND(125,-1.9)', 130),
        ('ROUND(1)', '#VALUE!'), ('ROUND(1,2,3)', '#VALUE!'), ('ROUND()', '#VALUE!'), ('ABS(1,2)', '#VALUE!'),
        ('LEN()', '#VALUE!'), ('MID("abc",1)', '#VALUE!'), ('DATE(2024,1)', '#VALUE!'), ('IFERROR(1)', '#VALUE!'),
        ('NOT(1,2)', '#VALUE!'), ('LEFT("abc",999)', 'abc'), ('RIGHT("abc",-1)', '#VALUE!'),
        ('MID("abc",99,1)', ''), ('MID("abc",0,1)', '#VALUE!'), ('SUBSTITUTE("abc","a","x",0)', '#VALUE!'),
        ('REPLACE("abc",5,1,"x")', 'abcx'), ('FIND("","abc",4)', 4), ('SEARCH("","abc",4)', 4), ('FIND("z","abc")', '#VALUE!'),
        ('MOD(5,-3)', -1), ('MOD(-5,3)', 1), ('MOD(-5,-3)', -2), ('MOD(5,0)', '#DIV/0!'), ('SQRT(-1)', '#NUM!'),
        ('SUM(A1:A4)', 1.25), ('COUNT(A1:A4)', 1), ('COUNTA(A1:A4)', 4), ('COUNTBLANK(A1:A4)', 1),
        ('SUM(B1)', '#N/A'), ('IFERROR(B1,7)', 7), ('_xlfn.IFNA(B1,8)', 8), ('ISBLANK(A4)', False), ('ISBLANK(B2)', True),
        ('INDEX(A1:B2,3,1)', '#REF!'), ('VLOOKUP(1.25,A1:B2,3,FALSE)', '#REF!'),
    ]: add("argument-type-error", expression, value)
    for name, count, token, expected in [('SUM',255,'1',255),('COUNT',255,'1',255),('AND',255,'TRUE',True),('OR',255,'FALSE',False),('CONCATENATE',255,'"x"','x'*255),('_xlfn.CONCAT',253,'"x"','x'*253)]:
        add("maximum-arity", name + '(' + ','.join([token]*count) + ')', expected)
    add("maximum-arity", '_xlfn.TEXTJOIN(",",TRUE,' + ','.join(['"x"']*252) + ')', ','.join(['x']*252))
    add("maximum-arity", '_xlfn.IFS(' + ','.join(['FALSE,0']*126 + ['TRUE,7']) + ')', 7)
    add("maximum-arity", '_xlfn.CONCAT(' + ','.join(['"x"']*254) + ')', '#VALUE!')
    add("maximum-arity", '_xlfn.TEXTJOIN(",",TRUE,' + ','.join(['"x"']*253) + ')', '#VALUE!')
    text = 'A禰󠄀😀B'
    for expression, value in [(f'LEN("{text}")',7), (f'FIND("😀","{text}")',5), (f'SEARCH("😀","{text}")',5),
        (f'MID("{text}",5,2)','😀'), (f'LEFT("{text}",4)','A禰󠄀'), (f'RIGHT("{text}",3)','😀B'),
        (f'REPLACE("{text}",5,2,"尾")','A禰󠄀尾B'), (f'SUBSTITUTE("{text}","😀","X")','A禰󠄀XB'), (f'EXACT("{text}","{text}")',True)]:
        add("unicode-version-1", expression, value, UNICODE_BASIS)
    date2024 = 45351 if system == '1900' else 43889
    for expression, value in [('DATE(1904,1,1)',1462 if system=='1900' else 0), ('DATE(1904,2,29)',1521 if system=='1900' else 59),
        ('DATE(1904,1,0)',1461 if system=='1900' else '#NUM!'), ('YEAR(0)',1900 if system=='1900' else 1904),
        ('MONTH(0)',1), ('DAY(0)',0 if system=='1900' else 1), ('DAY(59)',28 if system=='1900' else 29), ('DAY(60)',29 if system=='1900' else 1),
        ('DATE(2024,2,29)',date2024), ('EDATE(DATE(2024,1,31),1)',date2024), ('EOMONTH(DATE(2024,1,1),1)',date2024),
        ('_xlfn.DAYS(DATE(2024,3,1),DATE(2024,2,28))',2), ('DATE(-1,1,1)','#NUM!'), ('DATE(10000,1,1)','#NUM!'),
        ('TIME(25,0,0)',1/24), ('TIME(0,1500,0)',1/24), ('TIME(-1,0,0)','#NUM!'), ('TIME("abc",0,0)','#VALUE!'), ('HOUR(1.5)',12)]:
        add("date-system", expression, value, DATE_BASIS)
    return result


def generate(system, cached):
    book = native.fixture.Book(date1904=system == '1904'); sheet = book.sheet('Boundaries')
    for address, value in [('A1',1.25),('A2','2.5'),('A3',True)]: sheet.add(address,value)
    sheet.add('A4','' if cached else None,kind='str',formula='""',formula_attrs={'ca':'1'}); sheet.add('B1','#N/A',kind='e')
    for index, case in enumerate(cases(system),1): sheet.add('E'+str(index),999 if cached else None,formula=case['expression'][1:],formula_attrs={'ca':'1'})
    parts=book.finish(); parts['xl/workbook.xml']=parts['xl/workbook.xml'].replace('fullCalcOnLoad="0"','fullCalcOnLoad="1" forceFullCalc="1"')
    return native.fixture.zipped(parts)


def parts(data):
    with ZipFile(BytesIO(data)) as archive: return {name:archive.read(name) for name in archive.namelist()}


def observe(path, system):
    with ZipFile(path) as archive:
        sheet=ET.fromstring(archive.read('xl/worksheets/sheet1.xml')); shared=[]
        if 'xl/sharedStrings.xml' in archive.namelist(): shared=[''.join(node.itertext()) for node in ET.fromstring(archive.read('xl/sharedStrings.xml'))]
    cells={cell.get('r'):cell for cell in sheet.findall('{*}sheetData/{*}row/{*}c')}; output=[]
    for row, case in enumerate(cases(system),1):
        cell=cells['E'+str(row)]; kind=cell.get('t','n'); raw=cell.findtext('{*}v')
        if kind=='b': value={'type':'boolean','value':raw=='1'}
        elif kind=='e': value={'type':'error','value':raw}
        elif kind in ['s','str','inlineStr']: value={'type':'text','value':shared[int(raw)] if kind=='s' else ''.join(cell.find('{*}is').itertext()) if kind=='inlineStr' else raw or ''}
        else: value={'type':'number','value':float(raw)}
        output.append({**case,'address':'E'+str(row),'observed':value,'producerExpression':cell.findtext('{*}f'),'matchesProducer':value==case['expected']})
    return output


def main():
    parser=argparse.ArgumentParser(description=__doc__); parser.add_argument('--record',action='store_true'); parser.add_argument('--verify-native',action='store_true'); args=parser.parse_args()
    old=None if args.record else json.loads(REPORT.read_text()); records=[]; directory=ROOT/'.cache/workbook-function-boundaries-native'; directory.mkdir(exist_ok=True)
    if args.record or args.verify_native:
        assert sha256((ROOT/'.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest()==native.ARCHIVE_HASH
        os.environ.update({'PATH':str(ROOT/'.cache/acceptance-office/bin')+os.pathsep+os.environ['PATH'],'LD_LIBRARY_PATH':str(ROOT/'.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu'),'LC_ALL':'C.UTF-8','LANG':'C.UTF-8','TZ':'UTC'})
        assert native.shared.run(['libreoffice','--version'])==native.BUILD
    for system in ['1900','1904']:
        paths={name:FIXTURES/f'workbook-function-boundaries-{system}{suffix}.xlsx' for name,suffix in [('source',''),('reference','-reference'),('producer','-calc')]}
        for name,cached in [('source',True),('reference',False)]:
            data=generate(system,cached)
            if args.record: paths[name].write_bytes(data)
            assert parts(paths[name].read_bytes())==parts(data), 'Independent source parts changed'
            shutil.copyfile(paths[name],directory/paths[name].name)
        if args.record or args.verify_native:
            with TemporaryDirectory(dir=directory) as temporary:
                work=Path(temporary); output=work/'out'; output.mkdir(); source=work/f'{system}.xlsx'; source.write_bytes(paths['reference'].read_bytes())
                native.shared.run(['libreoffice','-env:UserInstallation='+(work/'profile').as_uri(),'--headless','--convert-to','xlsx:Calc MS Excel 2007 XML','--outdir',str(output),str(source)])
                produced=output/source.name
                if args.record: shutil.copyfile(produced,paths['producer'])
                assert observe(produced,system)==observe(paths['producer'],system), 'Producer observations changed'
        shutil.copyfile(paths['producer'],directory/paths['producer'].name)
        records.append({'dateSystem':system,'files':{name:{'path':str(path.relative_to(ROOT)),'sha256':sha256(path.read_bytes()).hexdigest()} for name,path in paths.items()},'cases':observe(paths['producer'],system)})
    if args.record or args.verify_native:
        python=ROOT/'.cache/xlsx-venv/bin/python'; subprocess.run([str(python) if python.is_file() else sys.executable,str(ROOT/'scripts/check-xlsx-output.py'),str(directory)],check=True)
    report={'scope':'Independent decimal / argument / legacy UTF-16 / two-calendar inputs with intentionally wrong formula caches and unmodified pinned Calc resaves. Excel/WPS runtime remains unverified; Version 2 extensions are not silently interpreted as Version 1.',
        'office':native.BUILD,'archiveSha256':native.ARCHIVE_HASH,'basis':['https://support.microsoft.com/en-gb/excel/compatibility-versions','https://support.microsoft.com/en-gb/excel/functions/left-function','https://support.microsoft.com/en-gb/excel/functions/right-function','https://support.microsoft.com/en-gb/excel/functions/time-function','https://support.microsoft.com/en-us/excel/date-systems-in-excel','https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/wrongly-assumes-1900-is-leap-year'],'workbooks':records}
    if args.record: REPORT.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    else: assert old==report, 'Pinned immutable function evidence changed'
    print(f'Verified {sum(len(item["cases"]) for item in records)} independent function boundary observations')


if __name__=='__main__': main()
