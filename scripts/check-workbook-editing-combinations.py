"""Independent mixed-feature workbook and pinned Calc edit; never SDK output."""
from workbook_acceptance import audit_schema
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
import argparse
import importlib.util
import json
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('combination_axis', ROOT / 'scripts/check-workbook-axis.py')
axis = importlib.util.module_from_spec(spec)
spec.loader.exec_module(axis)
source = axis.source
FIXTURE = ROOT / 'packages/office/tests/fixtures/workbook-editing-combinations.xlsx'
OUTPUT = FIXTURE.with_name('workbook-editing-combinations-calc.xlsx')
REPORT = ROOT / 'docs/evidence/workbook-editing-combinations.json'


def generate():
    inputs = {name: parts for name, parts, _ in source.cases()}
    book = source.Book()
    sheet = book.sheet('Sales')
    for column, label in zip('ABCD', ['产品', '数量', '单价', '金额']):
        sheet.add(column + '1', label, style=1)
    for row, name, quantity, price in [(2, '甲', 2, 10), (3, '乙', 3, 20), (4, '丙', 4, 30), (5, '丁', 1, 40)]:
        if row == 2:
            sheet.add('A2', kind='inlineStr', rich='<r><rPr><b/><rFont val="Noto Sans CJK SC"/></rPr><t>甲禰󠄀</t></r><r><rPr><i/><rFont val="Liberation Sans"/></rPr><t xml:space="preserve"> Latin😀</t></r>')
        else:
            sheet.add('A' + str(row), name)
        sheet.add('B' + str(row), quantity)
        sheet.add('C' + str(row), price, style=2)
        sheet.add('D' + str(row), 999, style=2, formula='SalesTable[[#This Row],[数量]]*SalesTable[[#This Row],[单价]]')
    sheet.add('A6', '合计', style=1)
    sheet.add('D6', 999, style=2, formula='SUBTOTAL(109,SalesTable[金额])')
    sheet.after = ('<conditionalFormatting sqref="D2:D5"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>50</formula></cfRule></conditionalFormatting>'
        '<dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="0" showErrorMessage="1" errorStyle="stop" sqref="B2:B5"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>')
    book.names.append(({'name': 'Quantities'}, 'Sales!$B$2:$B$5'))
    table = inputs['05-native-table.xlsx']['xl/tables/table1.xml']
    book.part('xl/tables/table1.xml', table, 'table')
    book.relate(sheet.path, 'table', source.R + '/table', 'xl/tables/table1.xml')
    # Existing independently authored drawing XML, with a fourth chart object.
    drawing = inputs['08-images-shapes.xlsx']['xl/drawings/drawing1.xml']
    content = drawing.split('>', 2)[2].rsplit('</xdr:wsDr>', 1)[0].replace('editAs="twoCell"', 'editAs="oneCell"')
    content += source.one_anchor(source.graphic_frame(4, 'chart'), 9, 1, 300, 190)
    path = source.add_drawing(book, sheet, content)
    book.part('xl/media/blue.png', source.png(), 'image/png')
    book.relate(path, 'image', source.R + '/image', 'xl/media/blue.png')
    chart = inputs['07-three-charts.xlsx']['xl/charts/chart1.xml'].replace('Data!', 'Sales!')
    book.part('xl/charts/chart1.xml', chart, 'chart')
    book.relate(path, 'chart', source.R + '/chart', 'xl/charts/chart1.xml')
    source.add_classic(book, sheet)
    # Keep the classic note distinct from the modern B2 root and C3 thread.
    book.parts['xl/comments/comment1.xml'] = book.parts['xl/comments/comment1.xml'].replace('ref="B2"', 'ref="D4"')
    book.parts['xl/drawings/comments.vml'] = book.parts['xl/drawings/comments.vml'].replace('<x:Row>1</x:Row><x:Column>1</x:Column>', '<x:Row>3</x:Row><x:Column>3</x:Column>')
    source.add_modern(book, sheet)
    root = '{33333333-3333-4333-8333-333333333333}'
    second = '{55555555-5555-4555-8555-555555555555}'
    book.parts['xl/comments/comment1.xml'] = book.parts['xl/comments/comment1.xml'].replace('</authors>', f'<author>tc={root}</author><author>tc={second}</author></authors>').replace('</commentList>',
        '<comment ref="B2" authorId="1"><text><t>兼容缓存</t></text></comment><comment ref="C3" authorId="2"><text><t>兼容缓存</t></text></comment></commentList>')
    sheet.tail += '<tableParts count="1"><tablePart r:id="table"/></tableParts>'
    observer = book.sheet('Observer')
    observer.add('A1', 999, formula='SUM(SalesTable[金额])')
    observer.add('A2', 999, formula='SUM(Quantities)')
    return source.zipped(book.finish())


def observe(path):
    with ZipFile(path) as archive:
        result = {'sheets': [], 'tables': [], 'charts': [], 'classicNotes': [], 'modernThreads': [], 'duplicateVmlIds': []}
        for name in sorted(archive.namelist()):
            if name.endswith('.vml'):
                root = ET.fromstring(archive.read(name))
                seen = set()
                for node in root:
                    identity = node.get('id')
                    if identity in seen:
                        result['duplicateVmlIds'].append(identity)
                    if identity:
                        seen.add(identity)
            if not name.endswith('.xml'):
                continue
            root = ET.fromstring(archive.read(name))
            if re.fullmatch(r'xl/worksheets/sheet\d+.xml', name):
                result['sheets'].append({'part': name, 'cells': [{'address': cell.get('r'), 'type': cell.get('t', 'n'), 'formula': cell.findtext('{*}f'), 'value': cell.findtext('{*}v')} for cell in root.findall('{*}sheetData/{*}row/{*}c')],
                    'hiddenRows': [int(row.get('r')) - 1 for row in root.findall('{*}sheetData/{*}row') if row.get('hidden') in ['1', 'true']]})
            elif name.startswith('xl/tables/'):
                result['tables'].append({'name': root.get('name'), 'range': root.get('ref')})
            elif name.startswith('xl/charts/chart'):
                result['charts'].append([node.text for node in root.findall('.//{*}f')])
            elif root.tag.endswith('}comments'):
                result['classicNotes'].extend([node.get('ref') for node in root.findall('{*}commentList/{*}comment')])
            elif root.tag.endswith('}ThreadedComments'):
                result['modernThreads'].extend([node.get('ref') for node in root])
        return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--record', action='store_true')
    parser.add_argument('--verify-native', action='store_true')
    args = parser.parse_args()
    data = generate()
    if args.record:
        FIXTURE.write_bytes(data)
    assert FIXTURE.read_bytes() == data
    directory = ROOT / '.cache/workbook-editing-combinations-native'
    directory.mkdir(exist_ok=True)
    if args.record or args.verify_native:
        assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == axis.ARCHIVE_HASH
        def edit(document, prop):
            sheet = document.getSheets().getByIndex(0)
            sheet.getCellRangeByName('B2').setValue(6)
            document.calculateAll()
        produced = directory / 'edited.xlsx'
        axis.native(FIXTURE, produced, {}, mutate=edit)
        if args.record:
            shutil.copyfile(produced, OUTPUT)
        assert observe(produced) == observe(OUTPUT), 'Native mixed observations changed'
        shutil.copyfile(FIXTURE, directory / FIXTURE.name)
        shutil.copyfile(OUTPUT, directory / OUTPUT.name)
        audit_schema(directory)
    report = {'scope': 'Independent mixed native source: table, rich text, name, rules, three drawing anchor kinds, chart, classic note and modern threads. SDK current expectations are independent of Calc caches. Unmodified actual Calc edit/resave and differences retained; Excel/WPS runtime unverified.',
        'office': axis.BUILD, 'archiveSha256': axis.ARCHIVE_HASH,
        'sourceSha256': sha256(data).hexdigest(), 'outputSha256': sha256(OUTPUT.read_bytes()).hexdigest(),
        'expected': {'initialTotal': 240, 'editedTotal': 280, 'initialQuantitySum': 10, 'editedQuantitySum': 14, 'filteredEditedTotal': 120, 'drawingCount': 4, 'chartCount': 1, 'classicCount': 1, 'modernRoots': 2, 'modernReplies': 1},
        'producerImport': {'outcome': 'rejected', 'code': 'InvalidDocument', 'reason': 'Pinned Calc emits repeated VML type/shape identities for the three classic compatibility records. Raw bytes are preserved; XML/XSD/OPC success does not prove valid VML identities.'},
        'producer': observe(OUTPUT)}
    if args.record:
        REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    else:
        assert json.loads(REPORT.read_text()) == report
    print('Verified independent mixed editing source and unmodified native producer')


if __name__ == '__main__':
    main()
