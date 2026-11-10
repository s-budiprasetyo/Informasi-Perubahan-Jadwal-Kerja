"""Pembaca file: Jadwal Kerja SK (lapangan) dan HRD '25. Shuttle Kiln 1'."""
import io, re, warnings
import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
from .consts import norm, nn, parse_month_text, month_len

warnings.filterwarnings('ignore')


def _wb(f):
    return openpyxl.load_workbook(f, data_only=False)


def _is_blue(cell):
    """Sel berwarna biru muda = hari pengganti untuk menutup kurang hari."""
    f = cell.fill
    if not f or not f.fill_type or f.fgColor is None or f.fgColor.type != 'rgb' or not isinstance(f.fgColor.rgb, str): return False
    h = f.fgColor.rgb[-6:]
    try: r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except ValueError: return False
    return b >= 0xC0 and r <= 0xA8 and b > r


def _is_formula(v):
    return isinstance(v, ArrayFormula) or (isinstance(v, str) and v.startswith('='))


class _Vals:
    """Pembaca nilai sel: memakai HASIL rumus yang tersimpan (bukan teks rumus / objek ArrayFormula).
    Sel rumus tanpa hasil tersimpan dicatat di `bad` (file perlu dibuka & disimpan ulang di Excel)."""
    def __init__(self, f):
        data = f.read() if hasattr(f, 'read') else open(f, 'rb').read()
        self.wf = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
        self.wv = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        self.bad = []

    def sheet(self, title):
        return self.wf[title], self.wv[title]

    def val(self, wsf, wsv, r, c):
        v = wsv.cell(r, c).value
        raw = wsf.cell(r, c).value
        if v is None and _is_formula(raw):
            self.bad.append(wsf.cell(r, c).coordinate); return None
        return v if not _is_formula(v) else None


def _find_header(ws):
    """Cari baris header (berisi 'Nama' & 'PRN SAP') -> (row, {label: col})."""
    for r in range(1, 15):
        labels = {}
        for c in range(1, 60):
            v = ws.cell(r, c).value
            if isinstance(v, str): labels.setdefault(v.strip().upper(), c)
        if 'NAMA' in labels and any(k.startswith('PRN') for k in labels):
            return r, labels
    raise ValueError('Baris header (Nama / PRN SAP) tidak ditemukan di Jadwal Kerja.')


def _target(ws):
    for r in range(1, 12):
        for c in range(1, 12):
            if str(ws.cell(r, c).value or '').strip().upper().startswith('HARI KERJA'):
                t = ws.cell(r, c + 1).value
                m = re.match(r'\s*(\d+)', str(t or ''))
                if m: return int(m.group(1))
    return 22


def parse_jadwal(f):
    """-> dict(ym, target, people=[dict(no,nama,nik,item,seksi,pws,days)]).
    Kolom dicari dari baris header sehingga tata letak SK dan QC sama-sama terbaca."""
    V = _Vals(f); wb = V.wf
    ws = next((s for s in wb if any('JADWAL KERJA' in str(s.cell(r, c).value or '').upper()
                                      for r in (1, 2, 3) for c in (1, 2))), wb.active)
    wsv = V.wv[ws.title]; cv = lambda r, c: V.val(ws, wsv, r, c)
    ym = None
    for r in (1, 2, 3, 4):
        for c in (1, 2):
            ym = ym or parse_month_text(ws.cell(r, c).value)
    if ym is None:
        raise ValueError('Bulan tidak terbaca ("JADWAL KERJA KARYAWAN BULAN ...").')
    nd = month_len(ym)
    hr, L = _find_header(ws)
    col = lambda *names: next((L[n] for n in names if n in L), None)
    c_no = col('NO.', 'NO'); c_nama = col('NAMA'); c_nik = next(v for k, v in L.items() if k.startswith('PRN'))
    c_item = col('ITEM KERJA'); c_seksi = col('SEKSI'); c_pws = col('PWS')
    c_day = next((c for c in range(1, 60) if ws.cell(hr + 1, c).value == 1 or wsv.cell(hr + 1, c).value == 1), 8)
    target = _target(wsv)
    people = []
    for row in range(hr + 2, ws.max_row + 1):
        nama = cv(row, c_nama)
        if str(nama or '').strip().upper().startswith('KETERANGAN'):
            break
        raw = str(cv(row, c_nik) or '').strip()
        if raw.endswith('.0'): raw = raw[:-2]
        nik = nn(raw)
        if not nama or not raw.isdigit() or not nik:
            continue
        days = [norm(cv(row, c)) for c in range(c_day, c_day + nd)]
        rec = dict(no=cv(row, c_no) if c_no else None, nama=str(nama).strip(), nik=nik,
                   item=cv(row, c_item) if c_item else None,
                   seksi=str(cv(row, c_seksi) or '').strip().upper() if c_seksi else '',
                   pws=str(cv(row, c_pws) or '').strip(), days=days,
                   repl=[i + 1 for i, c in enumerate(range(c_day, c_day + nd)) if _is_blue(ws.cell(row, c)) and days[i] != 'OFF'])
        dup = next((i for i, q in enumerate(people) if q['nik'] == nik), None)
        if dup is None: people.append(rec)
        elif sum(x != 'OFF' for x in days) > sum(x != 'OFF' for x in people[dup]['days']): people[dup] = rec   # NIK ganda: simpan yang berisi
    if V.bad:
        raise ValueError(f'Jadwal Kerja memuat {len(V.bad)} sel rumus tanpa hasil tersimpan (mis. {", ".join(V.bad[:5])}). '
                         'Buka file di Excel lalu simpan ulang (Ctrl+S) agar hasil rumus tersimpan, kemudian unggah lagi.')
    return dict(ym=ym, target=target, people=people)


def parse_hrd(f):
    """-> dict(ym, rows={nik: dict(nik,dept,nama,pws,days,total)})"""
    V = _Vals(f); ws = V.wf.active; wsv = V.wv[ws.title]; cv = lambda r, c: V.val(ws, wsv, r, c)
    ym = None
    for r in range(1, 4):                       # judul bulan bisa di kolom mana saja (A/B/C ...)
        for c in range(1, 12):
            ym = ym or parse_month_text(ws.cell(r, c).value)
    hdr = 4
    tot_col = None
    for c in range(1, ws.max_column + 1):
        if 'TOTAL' in str(ws.cell(hdr, c).value or '').upper():
            tot_col = c
    nd = month_len(ym) if ym else ((tot_col - 5) if tot_col else 31)    # kolom tanggal berakhir sebelum kolom Total
    rows = {}
    for r in range(hdr + 1, ws.max_row + 1):
        nik = cv(r, 1)
        if nik is None or str(nik).strip() == '':
            continue
        nik = nn(nik)
        if not nik: continue
        total = cv(r, tot_col) if tot_col else None
        rows[nik] = dict(nik=nik, dept=cv(r, 2), nama=str(cv(r, 3) or '').strip(),
                         pws=str(cv(r, 4) or '').strip(),
                         days=[norm(cv(r, c)) for c in range(5, 5 + nd)],
                         total=total if isinstance(total, (int, float)) else None)
    if V.bad:
        raise ValueError(f'File HRD memuat {len(V.bad)} sel rumus tanpa hasil tersimpan (mis. {", ".join(V.bad[:5])}). '
                         'Buka di Excel lalu simpan ulang, kemudian unggah lagi.')
    return dict(ym=ym, rows=rows)
