import sys, datetime
sys.path.insert(0, '.')
from core.parsers import parse_jadwal, parse_hrd
from core import engine
from core.depts import DEPTS
from core.qc_writer import render_qc
A = '/mnt/user-data/attachments/1a0376e9-f1cd-5af9-96bf-99b03536ce1c/'
cfg = DEPTS['QC']
jad = parse_jadwal(A + '5fbf0280-JADWAL_KERJA_OKTOBER_QC.xlsx'); hrd = parse_hrd(A + '1c09f081-23._QC.xlsx')
plan, snap = engine.compute(jad, hrd['rows'], None, {'962': '05D002'}, None, cfg)
info = render_qc(plan, 0, datetime.date(2026, 10, 9), 'assets/template_qc.xlsx', '/tmp/qc_r0.xlsx', cfg)
print(info)
