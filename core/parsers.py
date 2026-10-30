"""Pembaca file: Jadwal Kerja SK (lapangan) dan HRD '25. Shuttle Kiln 1'."""
import re, warnings
import openpyxl
from .consts import norm, parse_month_text, month_len

warnings.filterwarnings('ignore')


def _wb(f):
    return openpyxl.load_workbook(f, data_only=False)


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
    wb = _wb(f)
    ws = next((s for s in wb if any('JADWAL KERJA' in str(s.cell(r, c).value or '').upper()
                                      for r in (1, 2, 3) for c in (1, 2))), wb.active)
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
    c_day = next((c for c in range(1, 60) if ws.cell(hr + 1, c).value == 1), 8)
    target = _target(ws)
    people = []
    for row in range(hr + 2, ws.max_row + 1):
        nama = ws.cell(row, c_nama).value
        if str(nama or '').strip().upper().startswith('KETERANGAN'):
            break
        nik = str(ws.cell(row, c_nik).value or '').strip()
        if not nama or not nik.isdigit():
            continue
        days = [norm(ws.cell(row, c).value) for c in range(c_day, c_day + nd)]
        people.append(dict(no=ws.cell(row, c_no).value if c_no else None, nama=str(nama).strip(), nik=nik,
                           item=ws.cell(row, c_item).value if c_item else None,
                           seksi=str(ws.cell(row, c_seksi).value or '').strip().upper() if c_seksi else '',
                           pws=str(ws.cell(row, c_pws).value or '').strip(), days=days))
    return dict(ym=ym, target=target, people=people)


def parse_hrd(f):
    """-> dict(ym, rows={nik: dict(nik,dept,nama,pws,days,total)})"""
    wb = _wb(f)
    ws = wb.active
    ym = parse_month_text(ws['B2'].value) or parse_month_text(ws['A2'].value)
    nd = month_len(ym) if ym else 31
    hdr = 4
    tot_col = None
    for c in range(1, ws.max_column + 1):
        if 'TOTAL' in str(ws.cell(hdr, c).value or '').upper():
            tot_col = c
    rows = {}
    for r in range(hdr + 1, ws.max_row + 1):
        nik = ws.cell(r, 1).value
        if nik is None or str(nik).strip() == '':
            continue
        nik = str(nik).strip()
        total = ws.cell(r, tot_col).value if tot_col else None
        rows[nik] = dict(nik=nik, dept=ws.cell(r, 2).value, nama=str(ws.cell(r, 3).value or '').strip(),
                         pws=str(ws.cell(r, 4).value or '').strip(),
                         days=[norm(ws.cell(r, c).value) for c in range(5, 5 + nd)],
                         total=total if isinstance(total, (int, float)) else None)
    return dict(ym=ym, rows=rows)
