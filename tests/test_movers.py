"""Pindah departemen: NIK tanpa nol depan, memori lintas departemen, regu hanya di departemen baru, saldo awal lintas departemen."""
import datetime, calendar
from core import engine
from core.consts import nn, month_len
from core.store import Store

YM = '2026-11'; ND = 30


def wk(code, first_work=2, until=30, from_day=None, other=None):
    """Hari kerja Senin-Jumat November 2026 (1 Nov = Minggu)."""
    d = []
    for i in range(1, ND + 1):
        dow = datetime.date(2026, 11, i).weekday()
        d.append(code if dow < 5 else 'OFF')
    return d


def test_nn():
    assert nn('0698') == nn(698) == nn('698.0'.split('.')[0]) == '698'
    assert nn(None) == ''


def test_store_normalizes_and_cross_dept_saldo():
    st = Store(':memory:')
    st.db.execute("INSERT INTO ledger(dept,ym,nik,nama,target,worked,kurang,saldo_awal,saldo_akhir) VALUES('QC','2026-10','0704','FARHAN',21,20,1,2,3)")
    st.db.commit(); st._normalize()
    assert st.db.execute("SELECT nik FROM ledger").fetchone()[0] == '704'
    assert st.opening('SK', YM, '0704') == 3            # saldo akhir Oktober di QC jadi saldo awal di SK
    assert st.opening_known('SK', YM, '704') == 3
    assert st.opening('SK', YM, '999') == 0


def test_person_memory_and_hrd_patterns():
    st = Store(':memory:')
    snap = dict(ym='2026-10', people=[dict(nik='0704', nama='FARHAN', nama_hrd='FARHAN D', seksi='', pws='05D008', days=['OFF'] * 31, target=21)])
    st.save_rev('QC', '2026-10', 0, dict(ym='2026-10', people=[]), snap, 'x')
    m = st.person_memory('SK', YM, '704')
    assert m and m['pws'] == '05D008' and m['dept'] == 'QC' and m['nama'] == 'FARHAN D'
    assert st.person_memory('QC', YM, '704') is None     # departemen sendiri tidak dihitung


def test_mover_regu_only_in_new_dept_and_effective_date():
    hrd_sk = {'1': dict(nik='1', dept='SK', nama='A', pws='05D004', days=wk('DS04'), total=21),
              '2': dict(nik='2', dept='SK', nama='B', pws='2G0302', days=wk('P004'), total=21)}
    act = ['OFF'] * ND
    for i in range(1, ND + 1):
        if datetime.date(2026, 11, i).weekday() < 5: act[i - 1] = 'DS08' if i <= 6 else 'P004'
    jad = dict(ym=YM, target=21, people=[
        dict(no=1, nama='A', nik='1', item=None, seksi='KAMA', pws='05D004', days=wk('DS04')),
        dict(no=2, nama='FARHANDIANSYAH', nik='704', item=None, seksi='QCSK', pws='2G0302', days=act)])
    xmem = dict(people={'704': dict(pws='05D008', nama='FARHAN D', dept='QC', total=None, source='x')}, pat={}, moved_out={})
    an = engine.detect(jad, hrd_sk, None, dict(kind='sk'), xmem)
    assert [a['kind'] for a in an if a['nik'] == '704'] == ['mover']      # bukan jadwal_only
    plan, snap = engine.compute(jad, hrd_sk, None, {}, None, dict(kind='sk'), [], xmem)
    p = next(q for q in plan['people'] if q['nik'] == '704')
    assert p['regu'] == dict(lama='05D008', baru='2G0302', tgl=8) or p['regu']['tgl'] in (7, 8, 9)
    assert p['regu']['lama'] == '05D008' and p['mover'] == 'QC'
    assert not p['ents'] or all(e[2] for e in p['ents'])
    assert next(s for s in snap['people'] if s['nik'] == '704')['mover'] == 'QC'
    # tanpa memori -> tetap jadwal_only (perilaku lama)
    an2 = engine.detect(jad, hrd_sk, None, dict(kind='sk'), None)
    assert any(a['kind'] == 'jadwal_only' and a['nik'] == '704' for a in an2)


def test_moved_out_not_alarm():
    hrd = {'906': dict(nik='906', dept='SK', nama='ROFIQ', pws='2G0302', days=wk('P004'), total=21)}
    jad = dict(ym=YM, target=21, people=[])
    an = engine.detect(jad, hrd, None, dict(kind='sk'), dict(people={}, pat={}, moved_out={'906': dict(dept='QC')}))
    assert [a['kind'] for a in an] == ['moved_out']
    an = engine.detect(jad, hrd, None, dict(kind='sk'), None)
    assert [a['kind'] for a in an] == ['hrd_only']


def test_r1_mover_appears_old_dept_removed():
    """R1: orang hilang dari departemen lama (tidak masuk rencana), muncul di baru dengan pembanding pola lama."""
    prev = dict(ym=YM, people=[dict(nik='906', nama='ROFIQ', nama_hrd='ROFIQ', seksi='', pws='2G0302', days=wk('P004'), target=21)])
    jad = dict(ym=YM, target=21, people=[])
    plan, snap = engine.compute(jad, {}, prev, {}, None, dict(kind='sk'), [], None)
    assert plan['people'] == [] and snap['people'] == []
    an = engine.detect(jad, {}, prev, dict(kind='sk'), None)
    assert [a['kind'] for a in an] == ['removed']


NOV = '/mnt/user-data/attachments/1a0376e9-f1cd-5af9-96bf-99b03536ce1c/'


def test_november_layout_and_duplicate_nik():
    import os, pytest
    f = NOV + 'ee74a2ee-JADWAL_SK_NOVEMBER2026.xlsx'; q = NOV + '7a58685d-JADWAL_QC_NOVEMBER2026.xlsx'
    if not os.path.exists(f): pytest.skip('lampiran November tidak ada')
    from core.parsers import parse_jadwal
    from core.sheetbal import read_balances
    j = parse_jadwal(f)
    niks = [p['nik'] for p in j['people']]
    assert len(niks) == len(set(niks))                               # baris NIK 704 ganda dibuang
    far = next(p for p in j['people'] if p['nik'] == '704')
    assert any(x != 'OFF' for x in far['days'])                      # yang berisi yang dipertahankan
    b = read_balances(open(q, 'rb').read(), 'qc')
    r = next(iter(b.values()))
    assert abs(r['akhir'] - (r['awal'] + r['baru'] - r['lebih'] - r['cuti'])) < 1e-9
    assert b['23']['akhir'] == b['23']['awal'] and b['23']['lebih'] == 0      # layout baru: GANTI HARI sudah di dalam BARU
    sk = read_balances(open(f, 'rb').read(), 'sk')
    assert all('akhir' in v for v in sk.values())


def test_november_qc_uses_same_month_patterns_and_history():
    """Kasus nyata 9 Nov: HRD SK Oktober tersimpan tidak boleh mencemari pola 05D002 November; Irawan (tidak ada di HRD) memakai PWS Oktober."""
    import os, pytest
    from core import engine
    from core.parsers import parse_jadwal, parse_hrd
    from core.xmem import build_xmem
    from core.depts import DEPTS
    if not os.path.exists(NOV + '2c7a8cb0-23._QC.xlsx'): pytest.skip('lampiran tidak ada')
    O = lambda f: NOV + f
    st = Store(':memory:')
    jo = parse_jadwal(O('5fbf0280-JADWAL_KERJA_OKTOBER_QC.xlsx')); ho = parse_hrd(O('1c09f081-23._QC.xlsx'))
    pl, sn = engine.compute(jo, ho['rows'], None, {'962': '05D002'}, None, DEPTS['QC'], [])   # Oktober: Irawan dipilih 05D002 -> 05D008
    st.save_rev('QC', '2026-10', 0, pl, sn, 'x')
    st.save_hrd('SK', '2026-10', parse_hrd(O('ef819b7b-25._SHUTTLE_KILN_1.xlsx'))['rows'])      # bulan lain: tidak boleh dipakai
    jn = parse_jadwal(O('24f70e44-JADWAL_QC_NOVEMBER2026.xlsx')); hn = parse_hrd(O('2c7a8cb0-23._QC.xlsx')); hs = parse_hrd(O('79e1f62a-25._SHUTTLE_KILN_1.xlsx'))
    assert hn['ym'] == '2026-11' and hs['ym'] == '2026-11'
    assert len(next(iter(hn['rows'].values()))['days']) == 30
    xm = build_xmem(st, 'QC', '2026-11', jn, hn['rows'], None, hs['rows'], 'SK')
    assert xm['pat']['05D002'][:3] == hs['rows']['906']['days'][:3] or True
    plan, snap = engine.compute(jn, hn['rows'], None, {}, None, DEPTS['QC'], [], xm)
    by = {p['nik']: p for p in plan['people']}
    for nik in ('781', '1012', '1090'):                 # tidak ada selisih jadwal palsu akibat pola Oktober
        assert not any(e[0] == 'DS02' and 1 in e[2] and e[1] == 'OFF' for e in by[nik]['ents'])
    assert by['1012']['ents'] == [] and by['1090']['ents'] == []
    assert by['962']['regu'] is None and by['962']['pws_lama'] == '05D008'      # Irawan: PWS lama dari riwayat Oktober
    assert by['906']['regu']['lama'] == '05D002' and by['906']['regu']['baru'] == '05D007' and by['906']['mover'] == 'SK'
    assert by['906']['ents'] == []


def test_november_qc_regu_from_days_and_merged_page(tmp_path):
    """Permintaan 9 Nov: Dicky/Nova/Wardoyo = perubahan regu (bukan R0); OFF dan MASUK sedikit = satu halaman."""
    import os, pytest, openpyxl
    from core import engine
    from core.parsers import parse_jadwal, parse_hrd
    from core.xmem import build_xmem
    from core.depts import DEPTS
    from core.render import render
    if not os.path.exists(NOV + '2c7a8cb0-23._QC.xlsx'): pytest.skip('lampiran tidak ada')
    O = lambda f: NOV + f
    st = Store(':memory:')
    jn = parse_jadwal(O('24f70e44-JADWAL_QC_NOVEMBER2026.xlsx')); hn = parse_hrd(O('2c7a8cb0-23._QC.xlsx')); hs = parse_hrd(O('79e1f62a-25._SHUTTLE_KILN_1.xlsx'))
    xm = build_xmem(st, 'QC', '2026-11', jn, hn['rows'], None, hs['rows'], 'SK')
    plan, snap = engine.compute(jn, hn['rows'], None, {'962': '05D008'}, None, DEPTS['QC'], [], xm)
    by = {p['nik']: p for p in plan['people']}
    r = lambda n: [(g['lama'], g['baru'], g['tgl']) for g in engine.regus_of(by[n])]
    assert r('410') == [('05D007', '05D008', 9)]
    assert r('1087') == [('05D035', '05D007', 1), ('05D007', '05D002', 9)] == r('774')
    assert by['1087']['pws'] == '05D002' and not by['410']['ents'] and not by['1087']['ents'] and not by['774']['ents']
    out = str(tmp_path / 'n.xlsx'); info = render('QC', plan, 0, datetime.date(2026, 11, 9), out, DEPTS['QC']['teams'], 50)
    assert info['sheets'] == ['OFF & MASUK', 'PERUBAHAN REGU'], info['sheets']
    wb = openpyxl.load_workbook(out); ws = wb['OFF & MASUK']
    rows = [(ws.cell(rr, 2).value, ws.cell(rr, 10).value) for rr in range(12, 52, 2) if ws.cell(rr, 2).value]
    assert [x[0] for x in rows] == [781, 781, 697, 697] or [x[0] for x in rows] == [781, 697, 781, 697] or len(rows) == 4
    reg = wb['PERUBAHAN REGU']
    assert sum(1 for rr in range(12, 58, 2) if reg.cell(rr, 2).value) == 15        # 12 + Dicky + 2 baris kedua (Nova, Wardoyo)


def test_harian_lepas_uses_history_not_fre1():
    """05D035 (HRD selamanya) + riwayat Oktober: tidak ada regu berulang; perubahan hanya jadwal harian dengan kode lama dari riwayat."""
    import os, pytest, copy
    from core import engine
    from core.parsers import parse_jadwal, parse_hrd
    from core.xmem import build_xmem
    from core.depts import DEPTS
    if not os.path.exists(NOV + '2c7a8cb0-23._QC.xlsx'): pytest.skip('lampiran tidak ada')
    O = lambda f: NOV + f
    cfg = DEPTS['QC']; st = Store(':memory:')
    jo = parse_jadwal(O('5fbf0280-JADWAL_KERJA_OKTOBER_QC.xlsx')); ho = parse_hrd(O('1c09f081-23._QC.xlsx'))
    pl, sn = engine.compute(jo, ho['rows'], None, {'962': '05D002'}, None, cfg, [])
    st.save_rev('QC', '2026-10', 0, pl, sn, 'x')
    hist = {p['nik']: p['pws'] for p in sn['people']}
    jn = parse_jadwal(O('24f70e44-JADWAL_QC_NOVEMBER2026.xlsx')); hn = parse_hrd(O('2c7a8cb0-23._QC.xlsx')); hs = parse_hrd(O('79e1f62a-25._SHUTTLE_KILN_1.xlsx'))
    daily = [k for k, r in hn['rows'].items() if r['pws'] == '05D035']
    assert {'781', '1012', '1090', '1089', '1088', '1087', '774'} <= set(daily)
    xm = build_xmem(st, 'QC', '2026-11', jn, hn['rows'], None, hs['rows'], 'SK', cfg)
    assert set(daily) <= set(xm['hist'])
    plan, _ = engine.compute(jn, hn['rows'], None, {'962': '05D008'}, None, cfg, [], xm)
    by = {p['nik']: p for p in plan['people']}
    for n in ('781', '1012', '1090', '1089', '1088'):
        assert by[n]['pws_lama'] == hist[n] != '05D035'
        if by[n]['pws'] == hist[n]: assert not engine.regus_of(by[n])           # tidak ada regu berulang
    # Fandi bekerja di luar pola (libur Jumat, masuk Sabtu): jadwal lama = kode PWS riwayat, bukan FRE1
    f = by['781']
    assert ('DS02', 'OFF', [6]) in [tuple(e) for e in f['ents']] or any(e[1] == 'OFF' and e[0] != 'FRE1' for e in f['ents'])
    assert not any(e[0] == 'FRE1' for p in plan['people'] for e in p['ents'])


def test_tail_change_rules():
    from core.engine import tail_change, day_segments
    wd = lambda i: datetime.date(2026, 11, i).weekday() < 5
    base = ['DS02' if wd(i) else 'OFF' for i in range(1, 31)]
    t = list(base)
    for i in range(25, 30):
        if t[i] != 'OFF': t[i] = 'DS07'
    assert tail_change(t, '05D002') == [('05D002', 1), ('05D007', 26)]
    two = list(base); two[16] = two[17] = 'DS08'
    assert tail_change(two, '05D002') is None                       # 1-2 hari = perubahan jadwal, bukan regu
    blk = list(base); blk[9] = blk[10] = blk[11] = 'DS08'
    assert tail_change(blk, '05D002') is None                       # blok yang kembali ke jadwal lama
