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
CREATE TABLE IF NOT EXISTS settings(dept TEXT, key TEXT, value TEXT, PRIMARY KEY(dept,key));
CREATE TABLE IF NOT EXISTS opening(dept TEXT, ym TEXT, nik TEXT, saldo REAL, PRIMARY KEY(dept,ym,nik));
"""
TABLES = ['revisions', 'hrd', 'decisions', 'ledger', 'opening', 'settings']


class Store:
    def __init__(self, path):
        self.path = path
        if path != ':memory:': os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        self.db = sqlite3.connect(path, check_same_thread=False)
        self.db.row_factory = sqlite3.Row
        self.db.executescript(SCHEMA); self.lock = threading.Lock()

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
        """Kurang Hari (kolom AN) = target - hari kerja. saldo_akhir = saldo_awal + kurang
        (positif = masih berhutang hari, negatif = lebih hari)."""
        with self.lock:
            self.db.execute('DELETE FROM ledger WHERE dept=? AND ym=?', (dept, ym))
            for p in snapshot['people']:
                aw = self.opening(dept, ym, p['nik'])
                self.db.execute('INSERT INTO ledger VALUES(?,?,?,?,?,?,?,?,?)',
                                (dept, ym, p['nik'], p['nama'], p['target'], p['worked'], p['kurang'], aw, aw + p['kurang']))
            self.db.commit()

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

    # ---- pengaturan per departemen (mis. susunan tim QC)
    def set_setting(self, dept, key, value):
        self.db.execute('INSERT OR REPLACE INTO settings VALUES(?,?,?)', (dept, key, json.dumps(value))); self.db.commit()

    def get_setting(self, dept, key, default=None):
        r = self.db.execute('SELECT value FROM settings WHERE dept=? AND key=?', (dept, key)).fetchone()
        return json.loads(r[0]) if r else default

    def depts_with_data(self):
        return [r[0] for r in self.db.execute('SELECT DISTINCT dept FROM revisions')]
