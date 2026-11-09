"""Independent displayed-value filters and actual pinned Calc operations."""
from workbook_acceptance import audit_schema
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import importlib.util
import json
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('display_filter_axis', ROOT / 'scripts/check-workbook-axis.py')
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)
source = axis.source
FIXTURE = ROOT / 'packages/office/tests/fixtures/workbook-filter-display.xlsx'
OUTPUT = FIXTURE.with_name('workbook-filter-display-calc.xlsx')
REPORT = ROOT / 'docs/evidence/workbook-filter-display.json'
CASES = [('decimal', '0.00', 1, 2, '1.00'), ('rounded', '0.00', 1, 1.004, '1.00'), ('percent', '0%', .5, .6, '50%'),
         ('date', 'yyyy-mm-dd', 45351, 45352, '2024-02-29'), ('currency', '"$"#,##0.00', 12.5, 13.5, '$12.50'), ('text', '"ID-"@', '001', '002', 'ID-001')]


def generate():
    book = source.Book()
    ET.register_namespace('', source.S)
    styles = ET.fromstring(book.parts['xl/styles.xml'])
    formats, cells = styles.find('{*}numFmts'), styles.find('{*}cellXfs')
    records = []
    for table in [False, True]:
        for index, (name, code, first, second, selected) in enumerate(CASES):
            identity = ('Table' if table else 'Range') + name
            sheet = book.sheet(identity)
            style = len(cells)
            ET.SubElement(formats, '{' + source.S + '}numFmt', {'numFmtId': str(200 + style), 'formatCode': code})
            ET.SubElement(cells, '{' + source.S + '}xf', {'numFmtId': str(200 + style), 'fontId': '0', 'fillId': '0', 'borderId': '0', 'xfId': '0', 'applyNumberFormat': '1'})
            for column, label in zip('ABC', ['Key', 'Amount', 'Tag']): sheet.add(column + '1', label)
            for row, value in [(2, first), (3, second)]:
                sheet.add('A' + str(row), value, style=style)
                sheet.add('B' + str(row), 10 * (row - 1))
                sheet.add('C' + str(row), 'keep' if row == 2 else 'skip')
            sheet.add('D1', 999, formula='SUBTOTAL(9,B2:B3)')
            filters = '<autoFilter ref="A1:C3"><filterColumn colId="0"><filters><filter' + source.attrs({'val': selected}) + '/></filters></filterColumn>'
            if table: filters += '<filterColumn colId="2"><filters><filter val="keep"/></filters></filterColumn>'
            filters += '</autoFilter>'
            if table:
                path = f'xl/tables/table{index + 1}.xml'
                book.part(path, source.XML + '<table xmlns="' + source.S + '"' + source.attrs({'id': index + 1, 'name': 'Formatted_' + str(index + 1), 'displayName': 'Formatted_' + str(index + 1), 'ref': 'A1:C3', 'totalsRowShown': '0'}) + '>' + filters + '<tableColumns count="3">' + ''.join('<tableColumn' + source.attrs({'id': column + 1, 'name': label}) + '/>' for column, label in enumerate(['Key', 'Amount', 'Tag'])) + '</tableColumns></table>', 'table')
                book.relate(sheet.path, 'table', source.R + '/table', path)
                sheet.tail = '<tableParts count="1"><tablePart r:id="table"/></tableParts>'
            else: sheet.after = filters
            records.append({'sheet': identity, 'format': code, 'input': [first, second], 'selected': selected, 'table': table, 'expected': 30 if name == 'rounded' and not table else 10})
    formats.set('count', str(len(formats))); cells.set('count', str(len(cells)))
    book.parts['xl/styles.xml'] = source.XML + ET.tostring(styles, encoding='unicode')
    return source.zipped(book.finish()), records


def observe(path):
    with ZipFile(path) as archive:
        return [{'subtotal': sheet.findtext('.//{*}c[@r="D1"]/{*}v'), 'type': sheet.find('.//{*}c[@r="D1"]').get('t', 'n'),
                 'hiddenRows': [int(row.get('r')) - 1 for row in sheet.findall('{*}sheetData/{*}row') if row.get('hidden') in ['true', '1']]}
                for sheet in [ET.fromstring(archive.read(f'xl/worksheets/sheet{index + 1}.xml')) for index in range(12)]]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--record', action='store_true'); parser.add_argument('--verify-native', action='store_true'); args = parser.parse_args()
    data, cases = generate()
    if args.record: FIXTURE.write_bytes(data)
    assert FIXTURE.read_bytes() == data
    directory = ROOT / '.cache/workbook-filter-display-native'; directory.mkdir(exist_ok=True)
    if args.record or args.verify_native:
        assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == axis.ARCHIVE_HASH
        def edit(document, prop):
            import uno
            from com.sun.star.sheet.FilterOperator import EQUAL
            from com.sun.star.sheet.FilterConnection import AND
            for index, case in enumerate(cases):
                sheet = document.getSheets().getByIndex(index); target = sheet.getCellRangeByName('A1:C3'); descriptor = target.createFilterDescriptor(True)
                descriptor.ContainsHeader = True; descriptor.UseRegularExpressions = False
                field = uno.createUnoStruct('com.sun.star.sheet.TableFilterField'); field.Field = 0; field.Operator = EQUAL; field.IsNumeric = False; field.StringValue = case['selected']; fields = [field]
                if case['table']:
                    second = uno.createUnoStruct('com.sun.star.sheet.TableFilterField'); second.Field = 2; second.Connection = AND; second.Operator = EQUAL; second.IsNumeric = False; second.StringValue = 'keep'; fields.append(second)
                descriptor.setFilterFields(tuple(fields)); target.filter(descriptor)
            document.calculateAll()
        produced = directory / 'filtered.xlsx'; axis.native(FIXTURE, produced, {}, mutate=edit)
        if args.record: shutil.copyfile(produced, OUTPUT)
        assert observe(produced) == observe(OUTPUT)
        shutil.copyfile(FIXTURE, directory / FIXTURE.name); shutil.copyfile(OUTPUT, directory / OUTPUT.name)
        audit_schema(directory)
    report = {'scope': 'Independent intrinsic display labels for ordinary and two-column native-table filtering. SDK accepts exact native scalar spellings or intrinsic source-format text. Calc numeric coercion/format normalization is retained separately; this is not Excel/WPS runtime acceptance.',
        'office': axis.BUILD, 'archiveSha256': axis.ARCHIVE_HASH, 'sourceSha256': sha256(data).hexdigest(), 'outputSha256': sha256(OUTPUT.read_bytes()).hexdigest(),
        'basis': ['https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofilter', 'https://support.microsoft.com/en-us/excel/quick-start-filter-data-by-using-an-autofilter'], 'cases': cases, 'producer': observe(OUTPUT)}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    else: assert json.loads(REPORT.read_text()) == report
    print('Verified twelve independent ordinary/table displayed-value filters and native observations')


if __name__ == '__main__': main()
