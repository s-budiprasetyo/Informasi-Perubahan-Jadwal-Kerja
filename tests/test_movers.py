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
