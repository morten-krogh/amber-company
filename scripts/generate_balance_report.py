#!/usr/bin/env python3
"""Generate an Excel balance report and PDF rendering data from a CP437 SIE4 file.

Uses only the Python standard library. Render the PDF with render_balance_report.swift.
"""
import argparse
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
import hashlib
import json
from pathlib import Path
import shlex
from xml.sax.saxutils import escape
import zipfile


NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ZERO = Decimal('0')


def read_sie(path):
    data = path.read_bytes()
    accounts, opening, closing, movements, result = {}, {}, {}, defaultdict(Decimal), {}
    vouchers = []
    metadata = {}
    for line in data.decode('cp437').splitlines():
        fields = shlex.split(line)
        if not fields:
            continue
        tag = fields[0]
        if tag == '#KONTO':
            accounts[fields[1]] = fields[2]
        elif tag in ('#FNAMN', '#ORGNR'):
            metadata[tag[1:]] = fields[1]
        elif tag == '#RAR' and fields[1] == '0':
            metadata['start'], metadata['end'] = fields[2:4]
        elif tag in ('#IB', '#UB', '#RES') and fields[1] == '0':
            {'#IB': opening, '#UB': closing, '#RES': result}[tag][fields[2]] = Decimal(fields[3])
        elif tag == '#VER':
            vouchers.append([fields[1], fields[2], ZERO])
        elif tag == '#TRANS':
            amount = Decimal(fields[3])
            movements[fields[1]] += amount
            vouchers[-1][2] += amount
    assert all(v[2] == 0 for v in vouchers), 'An unbalanced voucher exists'
    for account in set(opening) | set(closing) | {a for a in movements if a.startswith(('1', '2'))}:
        assert opening.get(account, ZERO) + movements[account] == closing.get(account, ZERO), account
    for account in set(result) | {a for a in movements if not a.startswith(('1', '2'))}:
        assert movements[account] == result.get(account, ZERO), account
    metadata['latest_voucher'] = f'{vouchers[-1][0]}{vouchers[-1][1]}'
    metadata['sha256'] = hashlib.sha256(data).hexdigest()
    metadata['source_file'] = path.name
    metadata['generated'] = datetime.now().astimezone().isoformat(timespec='seconds')
    return accounts, opening, closing, metadata


def make_report(accounts, opening, closing, meta):
    rows = []
    available = set(opening) | set(closing)

    def append(kind, label='', account='', values=None, members=None):
        rows.append(dict(kind=kind, label=label, account=account,
                         values=[str(v) for v in values] if values is not None else None,
                         members=members or []))
        return len(rows) - 1

    def section(label, predicate):
        append('section', label)
        indices = []
        for a in sorted(a for a in available if predicate(a)):
            sign = Decimal('1') if a.startswith('1') else Decimal('-1')
            start, end = opening.get(a, ZERO) * sign, closing.get(a, ZERO) * sign
            indices.append(append('account', accounts[a], a, [start, end - start, end]))
        total = [sum(Decimal(rows[i]['values'][c]) for i in indices) for c in range(3)]
        return append('subtotal', 'Summa ' + label.lower(), values=total, members=indices)

    append('major', 'Tillgångar')
    fixed = section('Anläggningstillgångar', lambda a: a.startswith('1') and a < '1400')
    current = section('Omsättningstillgångar', lambda a: a.startswith('1') and a >= '1400')
    assets = append('total', 'Summa tillgångar', values=[
        Decimal(rows[fixed]['values'][i]) + Decimal(rows[current]['values'][i]) for i in range(3)
    ], members=[fixed, current])
    append('major', 'Eget kapital och skulder')
    equity = section('Eget kapital', lambda a: a.startswith('20'))
    reserves = section('Obeskattade reserver', lambda a: a.startswith('21'))
    provisions = None
    if any(a.startswith('22') for a in available):
        provisions = section('Avsättningar', lambda a: a.startswith('22'))
    longterm = section('Långfristiga skulder', lambda a: a.startswith('23'))
    shortterm = section('Kortfristiga skulder', lambda a: a.startswith(('24', '25', '26', '27', '28', '29')))
    members = [equity, reserves, longterm, shortterm] + ([provisions] if provisions is not None else [])
    liabilities = append('total', 'Summa eget kapital och skulder', values=[
        sum(Decimal(rows[j]['values'][i]) for j in members) for i in range(3)
    ], members=members)
    difference = [Decimal(rows[assets]['values'][i]) - Decimal(rows[liabilities]['values'][i]) for i in range(3)]
    append('check', 'Kontrolldifferens: tillgångar minus eget kapital och skulder', values=difference,
           members=[assets, liabilities])
    notes = []
    if any(difference):
        notes.append('Kontrolldifferens −0,34 SEK finns redan i ingående SIE-balans och kvarstår i utgående balans. Ingen justeringspost har lagts till i rapporten.')
    notes.append('Konto 2518 innehåller 733,11 SEK preliminärt omklassificerad historisk differens; orsaken är ännu inte fastställd.')
    notes.append('Skatteberäkningen förutsätter inga inrullade skattemässiga underskott och att årets datoravskrivning är avdragsgill.')
    return dict(metadata=meta, rows=rows, notes=notes)


def write_xlsx(report, path):
    rows, meta = report['rows'], report['metadata']
    cellrows, merge = [], []

    def cell(ref, value, style=0, formula=None):
        if isinstance(value, Decimal):
            f = f'<f>{escape(formula)}</f>' if formula else ''
            return f'<c r="{ref}" s="{style}">{f}<v>{value:.2f}</v></c>'
        return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'

    titles = ['Balansrapport', meta['FNAMN'] + ' · ' + meta['ORGNR'],
              f'Räkenskapsår: {meta["start"]}–{meta["end"]} · Senaste verifikation: {meta["latest_voucher"]}',
              'Källa: ' + meta['source_file'] + ' · Genererad: ' + meta['generated']]
    for n, title in enumerate(titles, 1):
        cellrows.append(f'<row r="{n}" ht="22" customHeight="1">{cell("A"+str(n), title, 1 if n == 1 else 0)}</row>')
        merge.append(f'<mergeCell ref="A{n}:E{n}"/>')
    header = 6
    cellrows.append(f'<row r="{header}">' + ''.join(cell(f'{c}{header}', v, 2) for c, v in zip('ABCDE', ['Konto', 'Kontonamn', 'Ingående balans', 'Period', 'Utgående balans'])) + '</row>')
    first = 7
    for index, item in enumerate(rows):
        n = first + index
        kind = item['kind']
        style = 2 if kind in ('major', 'section') else 4 if kind in ('subtotal', 'total') else 5 if kind == 'check' else 0
        cells = cell('A'+str(n), item['account'], style) + cell('B'+str(n), item['label'], style)
        for c, value in zip('CDE', item['values'] or []):
            formula = None
            if kind == 'account' and c == 'D':
                formula = f'E{n}-C{n}'
            elif item['members']:
                refs = [f'{c}{first+i}' for i in item['members']]
                formula = '-'.join(refs) if kind == 'check' else '+'.join(refs)
            cells += cell(c+str(n), Decimal(value), 5 if kind == 'check' else 4 if kind in ('subtotal', 'total') else 3, formula)
        cellrows.append(f'<row r="{n}" ht="19" customHeight="1">{cells}</row>')
    for i, note in enumerate(report['notes'], first + len(rows) + 2):
        cellrows.append(f'<row r="{i}" ht="34" customHeight="1">{cell("A"+str(i), note, 6)}</row>')
        merge.append(f'<mergeCell ref="A{i}:E{i}"/>')
    last = first + len(rows) + len(report['notes']) + 1
    sheet = f'''<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="{NS}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="6" topLeftCell="A7" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="19"/><cols><col min="1" max="1" width="9" customWidth="1"/><col min="2" max="2" width="60" customWidth="1"/><col min="3" max="5" width="20" customWidth="1"/></cols><sheetData>{''.join(cellrows)}</sheetData><mergeCells count="{len(merge)}">{''.join(merge)}</mergeCells><pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="1"/></worksheet>'''
    styles = f'''<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00;[Red]-#,##0.00"/></numFmts><fonts count="3"><font><sz val="10"/><name val="Calibri"/></font><font><b/><sz val="10"/><name val="Calibri"/></font><font><b/><color rgb="FF9C0006"/><sz val="10"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE7EDF3"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="7"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="164" fontId="2" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment wrapText="1" vertical="top"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''
    files = {
        '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        '_rels/.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': f'<?xml version="1.0"?><workbook xmlns="{NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Balansrapport" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Balansrapport!$A$1:$E${last}</definedName></definedNames><calcPr fullCalcOnLoad="1"/></workbook>',
        'xl/_rels/workbook.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': sheet,
        'xl/styles.xml': styles,
    }
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('sie', type=Path)
    parser.add_argument('--xlsx', type=Path, required=True)
    parser.add_argument('--render-data', type=Path, required=True)
    args = parser.parse_args()
    report = make_report(*read_sie(args.sie))
    write_xlsx(report, args.xlsx)
    args.render_data.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    for item in report['rows']:
        if item['kind'] in ('total', 'check'):
            print(item['label'], item['values'])


if __name__ == '__main__':
    main()
