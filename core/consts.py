"""Konstanta & fungsi kecil bersama."""
import calendar, re

BULAN_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus',
            'September', 'Oktober', 'November', 'Desember']
BULAN_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des']

# jam kerja per kode (teks persis seperti di form)
HOURS = {
    'DS04': '(08:00 s/d 17:00)', 'DS02': '(07:00 s/d 16:00)',
    'P004': '(07:00 s/d 16:00)', 'S005': '(16:00 s/d 00:30)',
    'DS08': '(06:00 s/d 15:00)', 'DS07': '(09:00 s/d 18:00)',
    'M003': '(00:00 s/d 08:00)', 'P005': '(08:00 s/d 16:00)', 'S004': '(16:00 s/d 24:00)',
}
# teks sheet PERUBAHAN REGU (baris ke-2)
REGU_TXT = {'05D004': 'DS04(08:00 s/d 17:00)', '05D002': 'DS02(07:00 s/d 16:00)',
            '05D035': 'DS04(08:00 s/d 17:00)',
            '05D008': 'DS08(06:00 s/d 15:00)', '05D007': 'DS07(09:00 s/d 18:00)',
            '2G0301': 'S005 & P004', '2G0302': 'P004 & S005'}
DAY_PWS = ('05D004', '05D002', '05D035')          # regu "jam 8" -> halaman pertama
# urutan baris dalam form (sesuai contoh September)
ORDER = [('DS04', 'OFF'), ('DS02', 'OFF'), ('P004', 'OFF'), ('S005', 'OFF'),
         ('OFF', 'DS04'), ('OFF', 'DS02'), ('OFF', 'P004'), ('OFF', 'S005'),
         ('P004', 'S005'), ('S005', 'P004'), ('P004', 'DS04'), ('S005', 'DS04'),
         ('S005', 'DS02'), ('P004', 'DS02'), ('DS02', 'DS04'), ('DS04', 'DS02')]


def norm(v):
    """Normalisasi kode jam kerja: kosong/IOFF -> OFF, DSO4 -> DS04."""
    if v is None or str(v).strip() == '':
        return 'OFF'
    s = str(v).strip().upper()
    if s.startswith('DS'):
        s = s.replace('O', '0')
    return 'OFF' if s in ('IOFF', 'OFF', '-', '–') else s


def month_len(ym):
    y, m = map(int, ym.split('-'))
    return calendar.monthrange(y, m)[1]


def prev_ym(ym):
    y, m = map(int, ym.split('-'))
    return f'{y - 1}-12' if m == 1 else f'{y}-{m - 1:02d}'


def ym_label(ym):
    y, m = map(int, ym.split('-'))
    return f'{BULAN_ID[m - 1]} {y}'


def parse_month_text(text):
    """'... BULAN OKTOBER 2026' -> '2026-10' (atau None)."""
    if not text:
        return None
    for i, b in enumerate(BULAN_ID):
        m = re.search(rf'{b}\s+(\d{{4}})', str(text), re.I)
        if m:
            return f'{m.group(1)}-{i + 1:02d}'
    return None


def sheet_suffix(rev):
    return '' if rev == 0 else f' R{rev}'
