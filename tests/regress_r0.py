import sys, datetime, json
sys.path.insert(0, '.')
from core.parsers import parse_jadwal, parse_hrd
from core import engine
from core.inf_writer import render_inf
A = '/mnt/user-data/attachments/1a0376e9-f1cd-5af9-96bf-99b03536ce1c/'
jad = parse_jadwal(A + 'dddeacc0-JADWAL_KERJA_SK.xlsx'); hrd = parse_hrd(A + 'ef819b7b-25._SHUTTLE_KILN_1.xlsx')
print(jad['ym'], len(jad['people']), len(hrd['rows']))
for a in engine.detect(jad, hrd['rows']): print(a['kind'], a['nik'], a['nama'], a['msg'][:90])
dec = {'409': '2G0301'}
plan, snap = engine.compute(jad, hrd['rows'], None, dec)
info = render_inf(plan, 0, datetime.date(2026, 10, 16), 'assets/template_inf.xlsx', '/tmp/r0.xlsx')
print(info)
