"""Backup/restore ke Google Sheets (opsional). Sumber kebenaran tetap SQLite lokal;
Google Sheets dipakai agar riwayat tidak hilang bila server/komputer diganti."""
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


def push(store, secrets):
    sh = _open(secrets)
    data = json.dumps(store.export_all(), ensure_ascii=False)
    ws = _ws(sh, 'data_backup'); ws.clear()
    parts = [data[i:i + CHUNK] for i in range(0, len(data), CHUNK)] or ['']
    ws.update(values=[[i, p] for i, p in enumerate(parts)], range_name='A1')
    # lembar terbaca manusia
    led = _ws(sh, 'saldo_hari'); led.clear()
    rows = [['dept', 'bulan', 'nik', 'nama', 'target', 'hari_kerja', 'kurang_hari', 'saldo_awal', 'saldo_akhir']]
    rows += [[r['dept'], r['ym'], r['nik'], r['nama'], r['target'], r['worked'], r['kurang'], r['saldo_awal'], r['saldo_akhir']]
             for r in store.export_all()['ledger']]
    led.update(values=rows, range_name='A1')
    hist = _ws(sh, 'riwayat'); hist.clear()
    hrows = [['dept', 'bulan', 'revisi', 'dibuat', 'dasar', 'file_jadwal', 'file_hrd']]
    hrows += [[r['dept'], r['ym'], r['rev'], r['created'], r['basis'], r['jadwal_file'], r['hrd_file']] for r in store.list_revs()]
    hist.update(values=hrows, range_name='A1')
    return len(parts)


def pull(store, secrets):
    ws = _open(secrets).worksheet('data_backup')
    vals = ws.get_all_values()
    data = ''.join(v[1] for v in sorted(vals, key=lambda r: int(r[0])) if len(v) > 1)
    store.import_all(json.loads(data))
