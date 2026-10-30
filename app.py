"""Aplikasi Streamlit - Informasi Perubahan Jadwal Jam Kerja (Shuttle Kiln).
Jalankan:  streamlit run app.py"""
import os, io, tempfile, datetime
import pandas as pd
import streamlit as st

from core.consts import BULAN_ID, ym_label, prev_ym, sheet_suffix
from core.parsers import parse_jadwal, parse_hrd
from core import engine, backup
from core.inf_writer import render_inf
from core.store import Store
from core.depts import DEPTS
from core import qc_layout
from core.render import render

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, 'assets', 'template_inf.xlsx')
DB_PATH = os.environ.get('JADWAL_DB', os.path.join(HERE, 'data', 'jadwal.db'))

st.set_page_config(page_title='Jadwal Kerja SK', page_icon='🗓️', layout='wide')


# ---------------- login sederhana (satu pengguna)
def _secret(key, default):
    try: return str(st.secrets.get(key, default))
    except Exception: return default


def login_gate():
    import hmac, time
    if st.session_state.get('auth'): return
    st.title('🔒 Masuk')
    with st.form('login'):
        u = st.text_input('User'); pw = st.text_input('Password', type='password')
        ok = st.form_submit_button('Masuk')
    if ok:
        if time.time() < st.session_state.get('lock_until', 0):
            st.error('Terlalu banyak percobaan. Tunggu beberapa detik.')
        elif hmac.compare_digest(u, _secret('app_user', 'admin')) and hmac.compare_digest(pw, _secret('app_password', '123a')):
            st.session_state['auth'] = True; st.session_state['fails'] = 0; st.rerun()
        else:
            n = st.session_state.get('fails', 0) + 1; st.session_state['fails'] = n
            if n >= 3: st.session_state['lock_until'] = time.time() + 15 * n
            st.error('User atau password salah.')
    st.stop()


login_gate()


@st.cache_resource
def get_store():
    s = Store(DB_PATH)
    if s.is_empty() and backup.configured(st.secrets):
        try: backup.pull(s, st.secrets, list(DEPTS))
        except Exception: pass
    return s


store = get_store()


def secrets_obj():
    try: return st.secrets
    except Exception: return {}


@st.cache_data(show_spinner=False)
def _pj(b): return parse_jadwal(io.BytesIO(b))


@st.cache_data(show_spinner=False)
def _ph(b): return parse_hrd(io.BytesIO(b))


def teams_of(dept):
    return store.get_setting(dept, 'teams', DEPTS[dept].get('teams'))


def pct_of(dept):
    return store.get_setting(dept, 'bulk_pct', DEPTS[dept].get('bulk_pct', 50))


def render_bytes(dept, plan, rev, gen_date):
    with tempfile.TemporaryDirectory() as d:
        p = os.path.join(d, 'x.xlsx')
        info = render(dept, plan, rev, gen_date, p, teams_of(dept), pct_of(dept))
        return open(p, 'rb').read(), info


def file_name(dept, ym, rev):
    y, m = map(int, ym.split('-'))
    return f'{DEPTS[dept]["filename"]}_{BULAN_ID[m - 1].upper()}_{y}{"" if rev == 0 else f"_R{rev}"}.xlsx'


def fmt_ents(p):
    return '; '.join(f"{a}→{b}: {', '.join(map(str, ds))}" for a, b, ds in p['ents'])


def plan_tables(plan):
    ch, rg = [], []
    for p in plan['people']:
        if p['ents']:
            ch.append({'NIK': p['nik'], 'Nama': p['nama_hrd'], 'Seksi': p['seksi'], 'PWS': p['pws'],
                       'Jumlah hari': sum(len(e[2]) for e in p['ents']), 'Perubahan (lama→baru: tanggal)': fmt_ents(p)})
        if p['regu']:
            rg.append({'NIK': p['nik'], 'Nama': p['nama_hrd'], 'Seksi': p['seksi'], 'Dari': p['regu']['lama'],
                       'Ke': p['regu']['baru'], 'Tgl efektif': p['regu']['tgl']})
    return pd.DataFrame(ch), pd.DataFrame(rg)


dept = st.sidebar.selectbox('Departemen', list(DEPTS), format_func=lambda k: f"{k} - {DEPTS[k]['label']}", key='dept_sel')
cfg = DEPTS[dept]
st.sidebar.caption(cfg['note'])
page = st.sidebar.radio('Menu', ['Buat Jadwal', 'Riwayat', 'Saldo Hari', 'Pengaturan'])
st.sidebar.caption('PT. Surya Pertiwi Nusantara - Gresik')
if st.sidebar.button('Keluar'):
    st.session_state.clear(); st.rerun()

# =====================================================================
if page == 'Buat Jadwal':
    st.title(f"Informasi Perubahan Jadwal Jam Kerja - {cfg['label']}")
    c1, c2, c3 = st.columns(3)
    today = datetime.date.today()
    yy = c1.number_input('Tahun', 2024, 2100, today.year, step=1)
    mm = c2.selectbox('Bulan', range(1, 13), index=today.month - 1, format_func=lambda i: BULAN_ID[i - 1])
    ym = f'{int(yy)}-{mm:02d}'
    auto_rev = store.next_rev(dept, ym)
    rev = c3.number_input('Revisi', 0, max(auto_rev, 0), auto_rev, step=1,
                          help='Otomatis dari riwayat. Mengisi angka lama = menimpa revisi itu (revisi sesudahnya gugur).')
    rev = int(rev)
    have = store.revs(dept, ym)
    st.caption(f'Riwayat {ym_label(ym)}: ' + (', '.join(f'R{r}' for r in have) if have else 'belum ada') +
               f'  |  Yang akan dibuat: **R{rev}**' + ('  ⚠️ menimpa revisi yang sudah ada' if rev in have else ''))

    stored_hrd = store.get_hrd(dept, ym)
    need_hrd = rev == 0 or stored_hrd is None
    b1, b2 = st.columns(2)
    f_j = b1.file_uploader(f"① Jadwal Kerja {dept} (jadwal terapan lapangan)", type=['xlsx'], key=f'fj{dept}{ym}{rev}')
    if need_hrd:
        f_h = b2.file_uploader(f"② HRD - {cfg['hrd_hint']} (jadwal dasar)", type=['xlsx'], key=f'fh{dept}{ym}{rev}')
    else:
        f_h = None
        b2.info('② File HRD bulan ini sudah tersimpan - tidak perlu diunggah. Pembanding = revisi sebelumnya R%d.' % (rev - 1))
    if not f_j or (need_hrd and not f_h):
        st.stop()

    try:
        jad = _pj(f_j.getvalue())
        hrd_rows = _ph(f_h.getvalue())['rows'] if f_h else stored_hrd
        if f_h:
            hm = _ph(f_h.getvalue())['ym']
            if hm and hm != ym: st.error(f'File HRD berisi bulan {hm}, sedangkan yang dipilih {ym}.'); st.stop()
    except Exception as e:
        st.error(f'File tidak terbaca: {e}'); st.stop()
    if jad['ym'] != ym:
        st.error(f"Jadwal Kerja berisi bulan {jad['ym']}, sedangkan yang dipilih {ym}. Ubah pilihan bulan."); st.stop()
    prev = None
    if rev > 0:
        r = store.get_rev(dept, ym, rev - 1)
        if not r:
            st.error(f'Revisi R{rev - 1} belum tersimpan, jadi R{rev} tidak punya pembanding.'); st.stop()
        prev = r['snapshot']

    anoms = engine.detect(jad, hrd_rows, prev, cfg)
    saved = store.get_decisions(dept, ym)
    ACT = ('pws_mismatch', 'hrd_only', 'jadwal_only')
    n_act = sum(1 for a in anoms if a['kind'] in ACT)
    top = st.columns([6, 1])
    with top[1].popover(f'🔔 {len(anoms)}', use_container_width=True):
        if not anoms: st.write('Tidak ada catatan.')
        for a in anoms:
            st.markdown(f"**{a['nama']}** ({a['nik']})  \n{a['msg']}")
            if a['kind'] == 'pws_mismatch':
                d = saved.get(a['nik'], a['suggest']); opts = a['options']
                st.selectbox('Pakai PWS', opts, index=opts.index(d) if d in opts else 0, key=f"dec_{a['nik']}")
            elif a['kind'] == 'jadwal_only':
                d = saved.get(a['nik'], '(pilih)'); opts = a['options']
                st.selectbox('PWS lama (dasar)', opts, index=opts.index(d) if d in opts else 0, key=f"dec_{a['nik']}")
            elif a['kind'] == 'hrd_only':
                d = saved.get(a['nik'], 'abaikan')
                st.selectbox('Keputusan', a['options'], index=a['options'].index(d) if d in a['options'] else 0,
                             key=f"dec_{a['nik']}", help='Masukkan ke jadwal (jadwal = jadwal HRD) atau abaikan.')
            st.divider()
    if n_act: top[0].warning(f'{n_act} catatan menunggu keputusan Anda (🔔).')
    decisions = {k: v for k, v in saved.items()}
    for a in anoms:
        if a['kind'] in ACT:
            decisions[a['nik']] = st.session_state.get(f"dec_{a['nik']}", saved.get(a['nik'], a.get('suggest', '(pilih)' if a['kind'] == 'jadwal_only' else 'abaikan')))
    plan0, _ = engine.compute(jad, hrd_rows, prev, decisions, None, cfg)
    tgl = {}
    with st.expander('Tanggal efektif perubahan regu (bisa diubah)', expanded=0 < sum(1 for p in plan0['people'] if p['regu']) <= 6):
        for p in plan0['people']:
            if p['regu']:
                tgl[p['nik']] = st.number_input(f"{p['nama_hrd']}  {p['regu']['lama']} → {p['regu']['baru']}", 1, 31,
                                                p['regu']['tgl'], key=f"tgl_{dept}_{ym}_{rev}_{p['nik']}")
    plan, snap = engine.compute(jad, hrd_rows, prev, decisions, tgl, cfg)

    basis = 'HRD (jadwal dasar)' if prev is None else f'revisi sebelumnya R{rev - 1}'
    st.subheader(f'Pratinjau R{rev}  -  dibandingkan dengan {basis}')
    ch, rg = plan_tables(plan)
    m = st.columns(4)
    m[0].metric('Orang dalam jadwal', len(plan['people'])); m[1].metric('Orang berubah', len(ch))
    m[2].metric('Baris perubahan', sum(len(p['ents']) for p in plan['people'])); m[3].metric('Perubahan regu', len(rg))
    if plan['skipped']: st.warning('Dilewati (tanpa pembanding): ' + ', '.join(plan['skipped']))
    if plan.get('pending'): st.error('Ada orang yang belum diputuskan di 🔔 (PWS lama). Pilih dulu sebelum membuat Excel.')
    if cfg['kind'] == 'qc':
        if rev == 0:
            bl, rest = qc_layout.split_bulk(plan, pct_of(dept))
            lines = [f"{'OFF' if b['kind'] == 'off' else 'MASUK'} tgl {b['date']}: {len(b['niks'])} orang" for b in bl]
            lines.append(f"R0 (tidak teratur): {sum(len(v) for v in rest.values())} baris")
            st.info('Sheet yang akan dibuat → ' + ' | '.join(lines) + ' | PERUBAHAN REGU')
        else:
            st.info(f'Sheet yang akan dibuat → R{rev} (semua perubahan dari R{rev - 1}) | PERUBAHAN REGU (bila ada)')
    t1, t2, t3 = st.tabs(['Perubahan jadwal', 'Perubahan regu', 'Kurang hari (kolom AN)'])
    t1.dataframe(ch, use_container_width=True, hide_index=True) if len(ch) else t1.info('Tidak ada perubahan jadwal.')
    t2.dataframe(rg, use_container_width=True, hide_index=True) if len(rg) else t2.info('Tidak ada perubahan regu.')
    led = pd.DataFrame([{'NIK': p['nik'], 'Nama': p['nama'], 'Target': p['target'], 'Hari kerja': p['worked'],
                         'Kurang hari': p['kurang'], 'Saldo awal': store.opening(dept, ym, p['nik'])} for p in snap['people']])
    led['Saldo akhir'] = led['Saldo awal'] + led['Kurang hari']
    t3.dataframe(led, use_container_width=True, hide_index=True)
    t3.caption('Kurang hari = target (file HRD) - hari kerja di Jadwal. Positif = masih kurang, negatif = lebih. Lebih hari bulan ini mengurangi kekurangan bulan lalu.')

    if st.button(f'💾 Simpan R{rev} & buat Excel', type='primary', disabled=bool(plan.get('pending'))):
        gen = datetime.date.today()
        data, info = render_bytes(dept, plan, rev, gen)
        store.save_rev(dept, ym, rev, plan, snap, basis, f_j.name, f_h.name if f_h else '')
        if f_h: store.save_hrd(dept, ym, hrd_rows)
        store.save_decisions(dept, ym, decisions)
        store.save_ledger(dept, ym, snap)
        st.session_state['out'] = (file_name(dept, ym, rev), data, info, rev)
        if backup.configured(secrets_obj()):
            try: backup.push(store, secrets_obj(), list(DEPTS)); st.session_state['bk'] = 'Backup Google Sheets berhasil.'
            except Exception as e: st.session_state['bk'] = f'Backup Google Sheets gagal: {e}'
    if 'out' in st.session_state:
        name, data, info, r_ = st.session_state['out']
        st.success(f'R{r_} tersimpan. Sheet: {", ".join(info["sheets"])}.  ' + st.session_state.get('bk', ''))
        st.download_button('⬇️ Unduh ' + name, data, name,
                           'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# =====================================================================
elif page == 'Riwayat':
    st.title('Riwayat revisi')
    rows = store.list_revs(dept)
    if not rows: st.info(f'Belum ada riwayat untuk {dept}.'); st.stop()
    df = pd.DataFrame(rows).rename(columns={'dept': 'Seksi', 'ym': 'Bulan', 'rev': 'Rev', 'created': 'Dibuat', 'basis': 'Dibandingkan dengan',
                                           'jadwal_file': 'File Jadwal', 'hrd_file': 'File HRD'})
    st.dataframe(df, use_container_width=True, hide_index=True)
    pick = st.selectbox('Lihat / unduh ulang', [f"{r['dept']} | {r['ym']} | R{r['rev']}" for r in rows])
    d_, y_, r_ = [x.strip() for x in pick.split('|')]; r_ = int(r_[1:])
    rec = store.get_rev(d_, y_, r_)
    ch, rg = plan_tables(rec['plan'])
    st.write(f"Dibandingkan dengan: **{rec['basis']}**")
    st.dataframe(ch, use_container_width=True, hide_index=True) if len(ch) else st.info('Tidak ada perubahan jadwal.')
    if len(rg): st.dataframe(rg, use_container_width=True, hide_index=True)
    data, info = render_bytes(d_, rec['plan'], r_, datetime.date.today())
    st.download_button('⬇️ Unduh ulang (tanggal header = hari ini)', data, file_name(d_, y_, r_),
                       'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# =====================================================================
elif page == 'Saldo Hari':
    st.title('Saldo kurang / lebih hari')
    months = store.ledger_months(dept)
    if not months: st.info('Belum ada data. Saldo terisi saat Anda menyimpan revisi.'); st.stop()
    ym = st.selectbox('Bulan', months, format_func=ym_label)
    df = pd.DataFrame(store.get_ledger(dept, ym))[['nik', 'nama', 'target', 'worked', 'kurang', 'saldo_awal', 'saldo_akhir']]
    df.columns = ['NIK', 'Nama', 'Target', 'Hari kerja', 'Kurang hari (AN)', 'Saldo awal', 'Saldo akhir']
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.caption('Saldo akhir = saldo awal + kurang hari bulan ini. Contoh: Okt kurang 4, Nov lebih 1 → sisa 3.')
    with st.expander('Koreksi saldo awal manual'):
        nik = st.selectbox('Karyawan', df['NIK'], format_func=lambda n: f"{n} - {df.set_index('NIK').loc[n, 'Nama']}")
        v = st.number_input('Saldo awal', value=float(df.set_index('NIK').loc[nik, 'Saldo awal']), step=1.0)
        if st.button('Simpan saldo awal'):
            store.set_opening(dept, ym, nik, v)
            snap = {'people': [dict(nik=r['nik'], nama=r['nama'], target=r['target'], worked=r['worked'], kurang=r['kurang'])
                               for r in store.get_ledger(dept, ym)]}
            store.save_ledger(dept, ym, snap); st.rerun()

# =====================================================================
else:
    st.title('Pengaturan & backup')
    st.write('Database lokal: `%s`' % DB_PATH)
    if cfg['kind'] == 'qc':
        st.subheader('Susunan halaman (tim) - QC')
        st.caption('Tiap tim = satu halaman pada sheet OFF / MASUK. Urutan kolom "No" menentukan urutan dalam halaman '
                   '(foreman tim di nomor 1). NIK yang belum terdaftar ikut tim orang di atasnya di Jadwal Kerja.')
        teams = teams_of(dept) or []
        names = {}
        for r in store.list_revs(dept)[:1]:
            for p in store.get_rev(r['dept'], r['ym'], r['rev'])['snapshot']['people']: names[p['nik']] = p['nama']
        rows = [dict(NIK=str(n), Nama=names.get(str(n), ''), Tim=ti + 1, No=oi + 1) for ti, t in enumerate(teams) for oi, n in enumerate(t)]
        ed = st.data_editor(pd.DataFrame(rows), num_rows='dynamic', use_container_width=True, hide_index=True, key='teams_ed')
        pct = st.slider('Batas perubahan serentak (% orang) untuk sheet OFF / MASUK', 20, 100, int(pct_of(dept)))
        if st.button('Simpan susunan tim'):
            ed = ed.dropna(subset=['NIK', 'Tim']); ed['Tim'] = ed['Tim'].astype(int)
            nt = [[str(r.NIK) for r in g.sort_values('No').itertuples()] for _, g in sorted(ed.groupby('Tim'))]
            store.set_setting(dept, 'teams', nt); store.set_setting(dept, 'bulk_pct', pct); st.success('Tersimpan.')
        st.divider()
    st.subheader('Backup Google Sheets')
    if backup.configured(secrets_obj()):
        st.success('Google Sheets terhubung. Lembar per departemen: data_backup_SK / _QC, saldo_hari_SK / _QC, riwayat_SK / _QC.')
        c1, c2 = st.columns(2)
        if c1.button('Backup sekarang (semua departemen)'):
            try: st.success(f'Selesai ({backup.push(store, secrets_obj(), list(DEPTS))} bagian).')
            except Exception as e: st.error(e)
        if c2.button('Pulihkan dari Google Sheets (menimpa data lokal)'):
            try: st.success('Dipulihkan: ' + ', '.join(backup.pull(store, secrets_obj(), list(DEPTS))))
            except Exception as e: st.error(e)
    else:
        st.info('Google Sheets belum dikonfigurasi. Lihat README.md bagian "Backup Google Sheets".')
    st.download_button('Unduh salinan database (.db)', open(DB_PATH, 'rb').read() if os.path.exists(DB_PATH) else b'', 'jadwal.db')
