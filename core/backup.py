"""Backup/restore ke Google Sheets (opsional). Sumber kebenaran tetap SQLite lokal.
Tiap departemen punya lembar sendiri dengan akhiran kode departemen:
  data_backup_SK / data_backup_QC   (JSON, untuk pemulihan)
  saldo_hari_SK  / saldo_hari_QC    (terbaca manusia)
  riwayat_SK     / riwayat_QC       (terbaca manusia)"""
import json, os

CHUNK = 40000


def _get(secrets, key):
    try: return secrets.get(key)
    except Exception: return None


def _creds_from(secrets):
    info = _get(secrets, 'gcp_service_account')
    if not info:
        raw = _get(secrets, 'gcp_json')          # seluruh isi file JSON ditempel apa adanya
        if raw:
            try: info = json.loads(raw)
            except Exception: info = None
    sheet_id = _get(secrets, 'backup_sheet_id')
    if info and sheet_id: return dict(info), sheet_id
    path = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON', 'service_account.json')
    sid = os.environ.get('BACKUP_SHEET_ID')
    if os.path.exists(path) and sid:
        return json.load(open(path)), sid
    return None, None


def configured(secrets):
    return _creds_from(secrets)[0] is not None


def _open(secrets):
    import gspread
    info, sid = _creds_from(secrets)
    if not info: raise RuntimeError('Google Sheets belum dikonfigurasi.')
    return gspread.service_account_from_dict(info).open_by_key(sid)


def _ws(sh, title, rows=200, cols=12):
    try: return sh.worksheet(title)
    except Exception: return sh.add_worksheet(title, rows, cols)


def push(store, secrets, depts):
    """Kirim data tiap departemen ke lembar akhiran masing-masing. Return jumlah bagian JSON."""
    sh = _open(secrets); total = 0
    for d in depts:
        ex = store.export_all(d)
        data = json.dumps(ex, ensure_ascii=False)
        ws = _ws(sh, f'data_backup_{d}'); ws.clear()
        parts = [data[i:i + CHUNK] for i in range(0, len(data), CHUNK)] or ['']
        ws.update(values=[[i, p] for i, p in enumerate(parts)], range_name='A1'); total += len(parts)
        led = _ws(sh, f'saldo_hari_{d}'); led.clear()
        rows = [['bulan', 'nik', 'nama', 'target', 'hari_kerja', 'kurang_hari', 'saldo_awal', 'saldo_akhir']]
        rows += [[r['ym'], r['nik'], r['nama'], r['target'], r['worked'], r['kurang'], r['saldo_awal'], r['saldo_akhir']]
                 for r in ex['ledger']]
        led.update(values=rows, range_name='A1')
        hist = _ws(sh, f'riwayat_{d}'); hist.clear()
        hrows = [['bulan', 'revisi', 'dibuat', 'dasar', 'file_jadwal', 'file_hrd']]
        hrows += [[r['ym'], r['rev'], r['created'], r['basis'], r['jadwal_file'], r['hrd_file']] for r in store.list_revs(d)]
        hist.update(values=hrows, range_name='A1')
    return total


def _read_chunks(ws):
    vals = ws.get_all_values()
    return ''.join(v[1] for v in sorted(vals, key=lambda r: int(r[0])) if len(v) > 1)


def pull(store, secrets, depts):
    """Pulihkan tiap departemen dari lembar akhirannya. Return daftar departemen yang dipulihkan."""
    sh = _open(secrets); done = []
    for d in depts:
        try: raw = _read_chunks(sh.worksheet(f'data_backup_{d}'))
        except Exception:
            if d != 'SK': continue
            try: raw = _read_chunks(sh.worksheet('data_backup'))      # nama lama (sebelum ada QC)
            except Exception: continue
        if raw:
            store.import_all(json.loads(raw), d); done.append(d)
    return done
