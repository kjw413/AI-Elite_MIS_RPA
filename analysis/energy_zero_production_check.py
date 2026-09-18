"""Read-only XML audit of production daily rows and energy reconciliation."""
from pathlib import Path
from collections import defaultdict
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from datetime import datetime
import ast
import json
import sys
import openpyxl
from energy_zero_20260918 import energy, DB, ROOT

sys.stdout.reconfigure(encoding='utf-8')
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
M = '{' + NS['m'] + '}'

def daily_rows(path):
    with ZipFile(path) as z:
        book = ET.fromstring(z.read('xl/workbook.xml'))
        ids = {x.attrib['Id']: x.attrib['Target'] for x in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        sheet = next(x for x in book.find(M+'sheets') if x.attrib['name']=='daily')
        rid = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        target = ids[rid]
        target = target.lstrip('/') if target.startswith('/') else 'xl/'+target
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            strings = [''.join(x.itertext()) for x in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        def value(c):
            if c.attrib.get('t') == 'inlineStr': return ''.join(c.find(M+'is').itertext())
            v = c.findtext(M+'v')
            if c.attrib.get('t') == 's': return strings[int(v)]
            if v is None: return None
            try: return float(v)
            except ValueError: return v
        with z.open(target) as stream:
            header = None
            for _, row in ET.iterparse(stream, events=('end',)):
                if row.tag != M+'row': continue
                if header is None:
                    header = {''.join(filter(str.isalpha,c.attrib['r'])):value(c) for c in row}
                    print('DAILY_HEADER', header, flush=True)
                else:
                    cells = {header.get(''.join(filter(str.isalpha,c.attrib['r']))): c for c in row}
                    day = value(cells['date'])
                    if isinstance(day, (int,float)): day = openpyxl.utils.datetime.from_excel(day).strftime('%Y-%m-%d')
                    day = str(day)[:10]
                    if '2026-09-01' <= day <= '2026-09-17':
                        yield day, value(cells['factory']), value(cells['item_code']), value(cells['actual_qty'])
                row.clear()

rows = list(daily_rows(DB/'DB_생산실적.xlsx'))
print('SEPTEMBER_ROWS',len(rows),flush=True)
# Use the exact pure aggregation helpers, without importing modules with runtime side effects.
source = ast.parse((ROOT/'energy_builder.py').read_text(encoding='utf-8-sig'))
names = {'_item_code','_mix_equivalent_qty','_aggregate_production_rows','_date_key'}
functions = ast.Module(body=[x for x in source.body if isinstance(x,ast.FunctionDef) and x.name in names],type_ignores=[])
from datetime import date
import logging
scope = {'datetime':datetime,'date':date,'GWANGJU_FACTORY_CODE':'F30','GWANGJU_WIP_MIX_CONVERSION':{'260014':10.91954,'260016':1.0,'260039':1.0,'260042':4.0,'260047':1.0,'260351':1.0,'260352':1.0},'log':logging.getLogger(__name__)}
scope['GWANGJU_MIX_CONVERSION'] = dict(scope['GWANGJU_WIP_MIX_CONVERSION'], **{'129998':10.91954,'129999':1.0})
exec(compile(functions,'energy_builder.py pure helpers','exec'),scope)
base = scope['_aggregate_production_rows'](rows,include_gwangju_wip=False)
print('DB_DAILY_DAYS',len(base),flush=True)
w = openpyxl.load_workbook(DB/'DB_재공품.xlsx',read_only=True,data_only=True)
ws = w['광주']
it = ws.values
headers = next(it)
print('WIP_HEADER',headers,flush=True)
wip = {}
for r in it:
    day = scope['_date_key'](r[0]) if r else None
    if not day or not '26-09-01' <= day <= '26-09-17': continue
    wip[day] = sum(float(r[i] or 0)*scope['GWANGJU_WIP_MIX_CONVERSION'][str(h)] for i,h in enumerate(headers) if str(h) in scope['GWANGJU_WIP_MIX_CONVERSION'])
w.close()
report = {'production_daily_september_rows':len(rows),'production_by_factory_date':[{'factory':f,'date':'20'+d,'kg':v} for (f,d),v in sorted(base.items())]}
current=energy(DB/'DB_에너지.xlsx')
mapping={'남양주1':'F10A','남양주2':'F10B','김해':'F20','광주':'F30','논산':'F40','경산':'F50'}
# Mapping verified against factories.py FACTORY_MASTER.
report['factory_mapping']=mapping
report['comparison']=[]
for (s,d),r in current.items():
    if not '2026-09-01'<=d<='2026-09-17':continue
    expected=base.get((mapping[s],d[2:]),0) + (wip.get(d[2:],0) if s=='광주' else 0)
    report['comparison'].append({'sheet':s,'date':d,'row':r['row'],'actual':r['values'][12],'production_daily':expected,'difference':expected-r['values'][12]})
for sheet in mapping:
    rs=[r for r in report['comparison'] if r['sheet']==sheet]
    print('RECONCILE',sheet,'mismatch_days',sum(abs(r['difference'])>0.01 for r in rs),'current',sum(r['actual'] for r in rs),'expected',sum(r['production_daily'] for r in rs),flush=True)
report['wip_by_date']=wip
path=ROOT/'analysis'/'energy_zero_production_check.json'
path.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print('SAVED',path,flush=True)

