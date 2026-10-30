"""Penyimpanan riwayat (SQLite) + ekspor/impor JSON untuk backup."""
import sqlite3, json, os, datetime, threading
from .consts import prev_ym

SCHEMA = """
CREATE TABLE IF NOT EXISTS revisions(dept TEXT, ym TEXT, rev INTEGER, created TEXT, basis TEXT,
  plan TEXT, snapshot TEXT, jadwal_file TEXT, hrd_file TEXT, PRIMARY KEY(dept,ym,rev));
CREATE TABLE IF NOT EXISTS hrd(dept TEXT, ym TEXT, data TEXT, PRIMARY KEY(dept,ym));
CREATE TABLE IF NOT EXISTS decisions(dept TEXT, ym TEXT, nik TEXT, value TEXT, PRIMARY KEY(dept,ym,nik));
CREATE TABLE IF NOT EXISTS ledger(dept TEXT, ym TEXT, nik TEXT, nama TEXT, target REAL, worked REAL, kurang REAL,
  saldo_awal REAL, saldo_akhir REAL, PRIMARY KEY(dept,ym,nik));
CREATE TABLE IF NOT EXISTS files(dept TEXT, ym TEXT, rev INTEGER, name TEXT, blob BLOB, PRIMARY KEY(dept,ym,rev));
CREATE TABLE IF NOT EXISTS settings(dept TEXT, key TEXT, value TEXT, PRIMARY KEY(dept,key));
CREATE TABLE IF NOT EXISTS cuti(id INTEGER PRIMARY KEY AUTOINCREMENT, dept TEXT, nik TEXT, nama TEXT, tgl TEXT, hari REAL, ket TEXT, ym TEXT);
CREATE TABLE IF NOT EXISTS opening(dept TEXT, ym TEXT, nik TEXT, saldo REAL, PRIMARY KEY(dept,ym,nik));
"""
TABLES = ['revisions', 'hrd', 'decisions', 'ledger', 'opening', 'settings', 'cuti']


class Store:
    def __init__(self, path):
        self.path = path
        if path != ':memory:': os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        self.db = sqlite3.connect(path, check_same_thread=False)
        self.db.row_factory = sqlite3.Row
        self.db.executescript(SCHEMA); self.lock = threading.Lock()
        try: self.db.execute('ALTER TABLE ledger ADD COLUMN cuti REAL DEFAULT 0'); self.db.commit()
        except sqlite3.OperationalError: pass

    # ---- revisi
    def revs(self, dept, ym):
        return [r[0] for r in self.db.execute('SELECT rev FROM revisions WHERE dept=? AND ym=? ORDER BY rev', (dept, ym))]

    def next_rev(self, dept, ym):
        r = self.revs(dept, ym)
        return (r[-1] + 1) if r else 0

    def get_rev(self, dept, ym, rev):
        r = self.db.execute('SELECT * FROM revisions WHERE dept=? AND ym=? AND rev=?', (dept, ym, rev)).fetchone()
        if not r: return None
        return dict(r, plan=json.loads(r['plan']), snapshot=json.loads(r['snapshot']))

    def save_rev(self, dept, ym, rev, plan, snapshot, basis, jadwal_file='', hrd_file=''):
        with self.lock:
            # menimpa revisi n berarti revisi sesudahnya tidak berlaku lagi
            self.db.execute('DELETE FROM revisions WHERE dept=? AND ym=? AND rev>?', (dept, ym, rev))
            self.db.execute('INSERT OR REPLACE INTO revisions VALUES(?,?,?,?,?,?,?,?,?)',
                            (dept, ym, rev, datetime.datetime.now().isoformat(timespec='seconds'), basis,
                             json.dumps(plan), json.dumps(snapshot), jadwal_file, hrd_file))
            self.db.commit()

    def delete_from(self, dept, ym, rev):
        """Hapus revisi `rev` dan semua sesudahnya. Saldo hari bulan itu dihitung ulang dari revisi terakhir yang tersisa."""
        with self.lock:
            self.db.execute('DELETE FROM revisions WHERE dept=? AND ym=? AND rev>=?', (dept, ym, rev))
            self.db.execute('DELETE FROM files WHERE dept=? AND ym=? AND rev>=?', (dept, ym, rev))
            self.db.commit()
        left = self.revs(dept, ym)
        if left: self.save_ledger(dept, ym, self.get_rev(dept, ym, left[-1])['snapshot'])
        else:
            with self.lock:
                for t in ('ledger', 'hrd', 'decisions', 'opening'):
                    self.db.execute(f'DELETE FROM {t} WHERE dept=? AND ym=?', (dept, ym))
                self.db.execute('DELETE FROM settings WHERE dept=? AND key=?', (dept, f'holidays:{ym}'))
                self.db.commit()

    def list_revs(self, dept=None):
        q = 'SELECT dept,ym,rev,created,basis,jadwal_file,hrd_file FROM revisions' + (' WHERE dept=?' if dept else '') + ' ORDER BY ym DESC,dept,rev DESC'
        return [dict(r) for r in self.db.execute(q, (dept,) if dept else ())]

    # ---- HRD
    def save_hrd(self, dept, ym, hrd_rows):
        self.db.execute('INSERT OR REPLACE INTO hrd VALUES(?,?,?)', (dept, ym, json.dumps(hrd_rows))); self.db.commit()

    def get_hrd(self, dept, ym):
        r = self.db.execute('SELECT data FROM hrd WHERE dept=? AND ym=?', (dept, ym)).fetchone()
        return json.loads(r[0]) if r else None

    # ---- keputusan lonceng
    def save_decisions(self, dept, ym, dec):
        with self.lock:
            for nik, v in dec.items():
                self.db.execute('INSERT OR REPLACE INTO decisions VALUES(?,?,?,?)', (dept, ym, nik, str(v)))
            self.db.commit()

    def get_decisions(self, dept, ym):
        return {r['nik']: r['value'] for r in self.db.execute('SELECT nik,value FROM decisions WHERE dept=? AND ym=?', (dept, ym))}

    # ---- saldo hari (kurang/lebih)
    def set_opening(self, dept, ym, nik, saldo):
        self.db.execute('INSERT OR REPLACE INTO opening VALUES(?,?,?,?)', (dept, ym, nik, saldo)); self.db.commit()

    def opening(self, dept, ym, nik):
        r = self.db.execute('SELECT saldo FROM opening WHERE dept=? AND ym=? AND nik=?', (dept, ym, nik)).fetchone()
        if r: return r[0]
        r = self.db.execute('SELECT saldo_akhir FROM ledger WHERE dept=? AND ym=? AND nik=?', (dept, prev_ym(ym), nik)).fetchone()
        return r[0] if r else 0

    def save_ledger(self, dept, ym, snapshot):
        """Kurang Hari (kolom AN) = target - hari kerja. saldo_akhir = saldo_awal + kurang - potong cuti
        (positif = masih berhutang hari, negatif = lebih hari)."""
        with self.lock:
            self.db.execute('DELETE FROM ledger WHERE dept=? AND ym=?', (dept, ym))
            for p in snapshot['people']:
                aw = self.opening(dept, ym, p['nik']); ct = self.cuti_sum(dept, ym, p['nik'])
                self.db.execute('INSERT INTO ledger(dept,ym,nik,nama,target,worked,kurang,saldo_awal,saldo_akhir,cuti) VALUES(?,?,?,?,?,?,?,?,?,?)',
                                (dept, ym, p['nik'], p['nama'], p['target'], p['worked'], p['kurang'], aw, aw + p['kurang'] - ct, ct))
            self.db.commit()

    def recalc_ledgers(self, dept, from_ym):
        """Hitung ulang saldo semua bulan >= from_ym (urut), mis. setelah potong cuti berubah."""
        for ym in sorted(m for m in self.ledger_months(dept) if m >= from_ym):
            rows = self.get_ledger(dept, ym)
            self.save_ledger(dept, ym, {'people': [dict(nik=r['nik'], nama=r['nama'], target=r['target'], worked=r['worked'],
                                                         kurang=r['kurang']) for r in rows]})

    # ---- potong cuti
    def add_cuti(self, dept, ym, nik, nama, tgl, hari, ket=''):
        with self.lock:
            self.db.execute('INSERT INTO cuti(dept,nik,nama,tgl,hari,ket,ym) VALUES(?,?,?,?,?,?,?)', (dept, nik, nama, str(tgl), float(hari), ket, ym))
            self.db.commit()
        self.recalc_ledgers(dept, ym)

    def delete_cuti(self, dept, cid):
        r = self.db.execute('SELECT ym FROM cuti WHERE dept=? AND id=?', (dept, cid)).fetchone()
        if not r: return
        with self.lock:
            self.db.execute('DELETE FROM cuti WHERE dept=? AND id=?', (dept, cid)); self.db.commit()
        self.recalc_ledgers(dept, r[0])

    def cuti_sum(self, dept, ym, nik):
        r = self.db.execute('SELECT COALESCE(SUM(hari),0) FROM cuti WHERE dept=? AND ym=? AND nik=?', (dept, ym, nik)).fetchone()
        return r[0] or 0

    def cuti_log(self, dept, ym=None):
        """Daftar potong cuti + sisa kurang hari SETELAH tiap pengajuan (urut tanggal pengajuan per karyawan per bulan)."""
        q = 'SELECT * FROM cuti WHERE dept=?' + (' AND ym=?' if ym else '') + ' ORDER BY ym, nik, tgl, id'
        out = []; run = {}
        for r in self.db.execute(q, (dept, ym) if ym else (dept,)):
            r = dict(r); key = (r['ym'], r['nik'])
            if key not in run:
                lg = self.db.execute('SELECT saldo_awal,kurang FROM ledger WHERE dept=? AND ym=? AND nik=?', (dept, r['ym'], r['nik'])).fetchone()
                run[key] = (lg['saldo_awal'] + lg['kurang']) if lg else self.opening(dept, r['ym'], r['nik'])
            run[key] -= r['hari']; r['sisa'] = run[key]; out.append(r)
        return sorted(out, key=lambda r: (r['ym'], r['tgl'], r['id']), reverse=True)

    def get_ledger(self, dept, ym):
        return [dict(r) for r in self.db.execute('SELECT * FROM ledger WHERE dept=? AND ym=? ORDER BY nama', (dept, ym))]

    def ledger_months(self, dept):
        return [r[0] for r in self.db.execute('SELECT DISTINCT ym FROM ledger WHERE dept=? ORDER BY ym DESC', (dept,))]

    # ---- backup
    def is_empty(self):
        return self.db.execute('SELECT COUNT(*) FROM revisions').fetchone()[0] == 0

    def export_all(self, dept=None):
        out = {}
        for t in TABLES:
            q, args = (f'SELECT * FROM {t} WHERE dept=?', (dept,)) if dept else (f'SELECT * FROM {t}', ())
            out[t] = [dict(r) for r in self.db.execute(q, args)]
        return out

    def import_all(self, data, dept=None):
        with self.lock:
            for t in TABLES:
                if dept: self.db.execute(f'DELETE FROM {t} WHERE dept=?', (dept,))
                else: self.db.execute(f'DELETE FROM {t}')
                for row in data.get(t, []):
                    cols = list(row)
                    self.db.execute(f'INSERT OR REPLACE INTO {t}({",".join(cols)}) VALUES({",".join("?" * len(cols))})',
                                    [row[c] for c in cols])
            self.db.commit()

    # ---- berkas Jadwal Kerja yang diunggah (lokal saja; terlalu besar untuk backup Sheets)
    def save_file(self, dept, ym, rev, name, blob):
        self.db.execute('INSERT OR REPLACE INTO files VALUES(?,?,?,?,?)', (dept, ym, rev, name, blob)); self.db.commit()

    def get_file(self, dept, ym, rev):
        r = self.db.execute('SELECT name, blob FROM files WHERE dept=? AND ym=? AND rev=?', (dept, ym, rev)).fetchone()
        return (r[0], bytes(r[1])) if r else None

    # ---- pengaturan per departemen (mis. susunan tim QC)
    def set_setting(self, dept, key, value):
        self.db.execute('INSERT OR REPLACE INTO settings VALUES(?,?,?)', (dept, key, json.dumps(value))); self.db.commit()

    def get_setting(self, dept, key, default=None):
        r = self.db.execute('SELECT value FROM settings WHERE dept=? AND key=?', (dept, key)).fetchone()
        return json.loads(r[0]) if r else default

    def depts_with_data(self):
        return [r[0] for r in self.db.execute('SELECT DISTINCT dept FROM revisions')]
