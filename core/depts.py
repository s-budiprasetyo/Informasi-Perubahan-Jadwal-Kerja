"""Konfigurasi per departemen. Departemen baru = tambah satu entri di sini."""

DEPTS = {
    'SK': dict(
        label='Shuttle Kiln', kind='sk', template='template_inf.xlsx',
        exempt_pws=('05D035',),        # FRE1 diperlakukan 05D004, tanpa form regu
        regu_name='hrd', filename='INF_PERB_JADWAL_JAM_KERJA_SK',
        hrd_hint='25. Shuttle Kiln 1',
        note='Satu file berisi dua seksi: KAMA (foreman Sulton/Fajar) dan QC SK (foreman Satrio), dibedakan kolom SEKSI di Jadwal Kerja.',
    ),
    'QC': dict(
        label='Quality Control', kind='qc', template='template_qc.xlsx',
        exempt_pws=(),
        regu_name='jadwal', filename='INF_PERB_JADWAL_JAM_KERJA_QC',
        hrd_hint='23. QC',
        bulk_pct=50,                   # >= 50% orang berubah serentak di satu tanggal -> sheet OFF / MASUK
        # urutan halaman (tim) seperti contoh Oktober 2026; foreman tim di urutan pertama
        teams=[
            ['23', '93', '416', '701', '692', '699', '781', '1012', '1090', '962', '1109', '1110', '1111', '717', '718', '418'],
            ['136', '710', '931', '421', '800', '704', '695', '693', '697', '690', '971', '708', '716', '689'],
            ['901', '709', '419', '408', '691', '755', '688', '700', '698', '410', '706', '1028', '949', '714', '1089', '1087', '774', '1088'],
        ],
        note='Tanpa seksi. Perubahan serentak (OFF/MASUK) dipisah per tanggal; sisanya masuk sheet R0 (atau R1, R2, ... pada revisi).',
    ),
}


def cfg(dept):
    return DEPTS[dept]
