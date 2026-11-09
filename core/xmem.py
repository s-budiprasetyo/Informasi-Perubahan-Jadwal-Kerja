"""Memori karyawan untuk mesin perbandingan: pindahan antar departemen dan orang yang belum ada di HRD bulan ini."""


def build_xmem(store, dept, ym, jad, hrd_rows, prev, x_rows=None, other_dept=None, cfg=None):
    """-> dict(people={nik: info}, pat={pws: hari}, moved_out={nik: info}).
    Pola PWS dari departemen lain HANYA dari bulan yang sama (unggahan didahulukan, lalu HRD tersimpan bulan itu).
    Orang yang tidak ada di HRD & revisi sebelumnya: jejak terbaru antara departemen lain (pindahan) dan riwayat departemen sendiri."""
    xmem = dict(people={}, pat={}, moved_out={})
    if x_rows:
        for s_ in x_rows.values(): xmem['pat'].setdefault(s_['pws'], s_['days'])
    for k_, v_ in store.other_patterns(dept, ym).items(): xmem['pat'].setdefault(k_, v_)
    xmem['hist'] = {}
    daily = (cfg or {}).get('daily_pws') or ()
    if hrd_rows and daily and not prev:        # harian lepas: jejak PWS terakhir dari riwayat departemen sendiri
        for n_, s_ in hrd_rows.items():
            if s_['pws'] in daily:
                m_ = store.own_memory(dept, ym, n_)
                if m_: xmem['hist'][n_] = m_
    prev_by = {p['nik'] for p in prev['people']} if prev else set()
    for p_ in jad['people']:
        n_ = p_['nik']
        if n_ in prev_by or (hrd_rows and n_ in hrd_rows): continue
        if x_rows and n_ in x_rows:
            s_ = x_rows[n_]
            xmem['people'][n_] = dict(pws=s_['pws'], nama=s_['nama'], dept=other_dept, total=None, ym=ym,
                                      source='HRD departemen lain (unggahan)')
            continue
        m_o = store.person_memory(dept, ym, n_); m_s = store.own_memory(dept, ym, n_)
        m_ = m_o if (m_o and (not m_s or m_o['ym'] >= m_s['ym'])) else m_s
        if m_: xmem['people'][n_] = m_
    if hrd_rows and not prev:
        jn = {p_['nik'] for p_ in jad['people']}
        for n_ in hrd_rows:
            if n_ not in jn:
                m_ = store.person_memory(dept, ym, n_)
                if m_ and m_['source'] != 'HRD' and m_['cur']: xmem['moved_out'][n_] = m_
    return xmem
