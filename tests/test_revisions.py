import sys, copy, datetime, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import openpyxl
from core.parsers import parse_jadwal, parse_hrd
from core import engine
from core.inf_writer import render_inf
from core.store import Store
A = '/mnt/user-data/attachments/1a0376e9-f1cd-5af9-96bf-99b03536ce1c/'
TPL = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'assets/template_inf.xlsx')
jad = parse_jadwal(A + 'dddeacc0-JADWAL_KERJA_SK.xlsx'); hrd = parse_hrd(A + 'ef819b7b-25._SHUTTLE_KILN_1.xlsx')


def apply(base, ents):
    d = list(base)
    for a, b, ds in ents:
        for x in ds:
            assert d[x - 1] == a, (x, d[x - 1], a); d[x - 1] = b
    return d


def entries_xlsx(path):
    wb = openpyxl.load_workbook(path)
    n = {}
    for ws in wb:
        n[ws.title] = sum(1 for r in range(12, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, int))
    return n


def test_r0_reconstructs():
    plan, snap = engine.compute(jad, hrd['rows'], None, {'409': '2G0301'})
    pat = engine.patterns(hrd['rows'])
    assert len(plan['people']) == 28 and not plan['skipped']
    for p, sp in zip(plan['people'], snap['people']):
        s = hrd['rows'][p['nik']]
        if p['regu']:
            E = p['regu']['tgl']; base = pat[p['pws_lama']][:E - 1] + pat[p['pws']][E - 1:]
        else:
            base = pat[p['pws_lama'] if p['exempt'] else p['pws']] if not p['exempt'] else pat['05D004']
        assert apply(base, p['ents']) == sp['days'], p['nama']


def mutate(j, changes):
    j2 = copy.deepcopy(j)
    for p in j2['people']:
        for nik, (day, code) in changes.items():
            if p['nik'] == nik: p['days'][day - 1] = code
    return j2


def test_r1_r2_chain(tmp_path):
    plan0, snap0 = engine.compute(jad, hrd['rows'], None, {'409': '2G0301'})
    sulton = next(p for p in jad['people'] if p['nama'].startswith('SULTON') or p['nik'] == '1047')['nik']
    qcnik = next(p['nik'] for p in jad['people'] if p['seksi'] == 'QCSK')
    j1 = mutate(jad, {sulton: (3, 'OFF' if jad['people'][0]['days'][2] != 'OFF' else 'DS04')})
    # ubah satu orang QC & satu orang shift
    shiftp = next(p for p in jad['people'] if p['pws'] in ('2G0301',) and p['seksi'] != 'QCSK')
    cur = shiftp['days'][4]
    j1 = mutate(j1, {shiftp['nik']: (5, 'S005' if cur != 'S005' else 'P004')})
    plan1, snap1 = engine.compute(j1, hrd['rows'], snap0, {})
    changed = [p for p in plan1['people'] if p['ents']]
    assert 1 <= len(changed) <= 3, [(p['nama'], p['ents']) for p in changed]
    for p in plan1['people']:
        base = next(x for x in snap0['people'] if x['nik'] == p['nik'])['days']
        assert apply(base, p['ents']) == next(x for x in snap1['people'] if x['nik'] == p['nik'])['days']
    out1 = str(tmp_path / 'r1.xlsx')
    info = render_inf(plan1, 1, datetime.date(2026, 10, 20), TPL, out1)
    assert info['sheets'] == ['SATRIO R1', 'SULTON R1', 'PERUBAHAN REGU R1']
    n = entries_xlsx(out1)
    assert n['SULTON R1'] + n['SATRIO R1'] == sum(len(p['ents']) for p in plan1['people'])
    # R2 dari R1: perubahan lagi pada orang yang sama hanya relatif R1
    j2 = mutate(j1, {shiftp['nik']: (5, cur)})          # kembali ke semula
    plan2, snap2 = engine.compute(j2, hrd['rows'], snap1, {})
    ch2 = [p for p in plan2['people'] if p['ents']]
    assert len(ch2) == 1 and ch2[0]['nik'] == shiftp['nik']
    out2 = str(tmp_path / 'r2.xlsx')
    assert render_inf(plan2, 2, datetime.date(2026, 10, 25), TPL, out2)['sheets'][1] == 'SULTON R2'
    # tidak ada perubahan sama sekali -> file tetap valid
    plan3, _ = engine.compute(j2, hrd['rows'], snap2, {})
    assert not any(p['ents'] or p['regu'] for p in plan3['people'])
    render_inf(plan3, 3, datetime.date(2026, 10, 26), TPL, str(tmp_path / 'r3.xlsx'))
    openpyxl.load_workbook(str(tmp_path / 'r3.xlsx'))


def test_regu_change_in_r1(tmp_path):
    plan0, snap0 = engine.compute(jad, hrd['rows'], None, {'409': '2G0301'})
    j1 = copy.deepcopy(jad)
    p = next(x for x in j1['people'] if x['pws'] == '05D004')
    p['pws'] = '05D002'
    plan1, _ = engine.compute(j1, hrd['rows'], snap0, {})
    r = next(x for x in plan1['people'] if x['nik'] == p['nik'])
    assert r['regu'] and r['regu']['lama'] == '05D004' and r['regu']['baru'] == '05D002'
    render_inf(plan1, 1, datetime.date(2026, 10, 20), TPL, str(tmp_path / 'x.xlsx'))


def test_ledger_carry(tmp_path):
    st = Store(str(tmp_path / 'a.db'))
    plan0, snap0 = engine.compute(jad, hrd['rows'], None, {'409': '2G0301'})
    snap0['people'][0]['kurang'] = 4
    st.save_ledger('SK', '2026-10', snap0)
    nov = copy.deepcopy(snap0); nov['people'][0]['kurang'] = -1
    st.save_ledger('SK', '2026-11', nov)
    row = next(r for r in st.get_ledger('SK', '2026-11') if r['nik'] == snap0['people'][0]['nik'])
    assert row['saldo_awal'] == 4 and row['saldo_akhir'] == 3


def test_store_revs_and_backup(tmp_path):
    st = Store(str(tmp_path / 'b.db'))
    assert st.next_rev('SK', '2026-10') == 0
    plan0, snap0 = engine.compute(jad, hrd['rows'], None, {'409': '2G0301'})
    st.save_rev('SK', '2026-10', 0, plan0, snap0, 'HRD'); st.save_rev('SK', '2026-10', 1, plan0, snap0, 'R0')
    assert st.next_rev('SK', '2026-10') == 2
    st.save_rev('SK', '2026-10', 0, plan0, snap0, 'HRD')       # timpa R0 -> R1 gugur
    assert st.revs('SK', '2026-10') == [0]
    data = st.export_all(); st2 = Store(':memory:'); st2.import_all(data)
    assert st2.get_rev('SK', '2026-10', 0)['plan']['ym'] == '2026-10'


def test_delete_from_recomputes_ledger(tmp_path):
    st = Store(str(tmp_path / 'd.db'))
    plan0, snap0 = engine.compute(jad, hrd['rows'], None, {'409': '2G0301'})
    st.save_rev('SK', '2026-10', 0, plan0, snap0, 'HRD', 'j.xlsx', 'h.xlsx')
    st.save_ledger('SK', '2026-10', snap0)
    st.set_setting('SK', 'holidays:2026-10', [17])
    plan1, snap1 = engine.compute(jad, hrd['rows'], snap0, {'409': '2G0301'})
    st.save_rev('SK', '2026-10', 1, plan1, snap1, 'R0', 'j.xlsx', '')
    st.delete_from('SK', '2026-10', 1)
    assert st.revs('SK', '2026-10') == [0]
    st.delete_from('SK', '2026-10', 0)
    assert st.revs('SK', '2026-10') == [] and not st.ledger_months('SK')
    assert st.get_setting('SK', 'holidays:2026-10', None) is None


def test_holiday_suppresses_change():
    d0 = {'409': '2G0301'}
    p0, _ = engine.compute(jad, hrd['rows'], None, d0)
    cand = [(p['nik'], x) for p in p0['people'] for a, b, ds in p['ents'] if b == 'OFF' for x in ds]
    assert cand
    nik, day = cand[0]
    p1, s1 = engine.compute(jad, hrd['rows'], None, d0, None, None, [day])
    q = next(p for p in p1['people'] if p['nik'] == nik)
    assert all(day not in ds for a, b, ds in q['ents'])
    assert p1['holidays'] == [day]
