"""Mesin perbandingan jadwal.

R0  : Jadwal Kerja  vs  pola PWS dari HRD
R>=1: Jadwal Kerja  vs  snapshot revisi sebelumnya
"""
from collections import defaultdict
import re
from .consts import ORDER, DAY_PWS, HOURS, month_len


def patterns(hrd_rows):
    pat = {}
    for r in hrd_rows.values():
        pat.setdefault(r['pws'], r['days'])
    if '05D004' in pat:
        pat.setdefault('05D035', pat['05D004'])
    return pat


def ensure_pat(pat, pws):
    """Pola PWS yang belum ada di HRD (mis. 05D008) diturunkan dari pola hari-kerja 05D0nn:
    hari masuk sama (Senin-Jumat), kode jam = 'DS' + dua digit terakhir."""
    if not pws or pws in pat: return pat.get(pws)
    m = re.fullmatch(r'05D0(\d\d)', pws)
    if not m: return None
    tpl = next((v for k, v in pat.items() if re.fullmatch(r'05D0\d\d', k) and k != '05D035'), None)
    if tpl is None: return None
    code = 'DS' + m.group(1)
    pat[pws] = [code if x != 'OFF' else 'OFF' for x in tpl]
    return pat[pws]


def diffs(base, act):
    d = defaultdict(list)
    for i, (a, b) in enumerate(zip(base, act)):
        if a != b:
            d[(a, b)].append(i + 1)
    return d


def order_entries(d):
    d = dict(d)
    ents = []
    for k in ORDER:
        if k in d:
            ents.append([k[0], k[1], d.pop(k)])
    for k in sorted(d):
        ents.append([k[0], k[1], d.pop(k)])
    return ents


def effective_date(act, newp):
    """Tanggal efektif perubahan regu: hari pertama memakai kode regu baru;
    1 bila sudah bekerja sesuai regu baru sejak awal / tidak ada hari kerja lain sebelumnya."""
    newc = {x for x in newp if x != 'OFF'}
    F = next((i for i in range(len(act)) if act[i] != 'OFF' and act[i] in newc), None)
    if F is not None and any(act[i] != 'OFF' for i in range(F)):
        return F + 1
    return 1


def blend_base(old, new, act, pat, exempt_pws=('05D035',)):
    """(baseline, regu_dict|None, exempt) untuk perbandingan R0."""
    exempt = old in exempt_pws
    if exempt: old = '05D004'
    ensure_pat(pat, new); ensure_pat(pat, old)
    newp, oldp = pat[new], pat[old]
    if exempt: return pat['05D004'], None, True
    if old == new: return newp, None, False
    E = effective_date(act, newp)
    return oldp[:E - 1] + newp[E - 1:], dict(lama=old, baru=new, tgl=E), False


def _same_pws(a, b):
    n = lambda x: 'DAY' if x in ('05D004', '05D035') else x
    return n(a) == n(b)


def detect(jad, hrd_rows, prev=None, cfg=None):
    """Daftar anomali (independen dari keputusan pengguna) untuk panel lonceng."""
    nd = month_len(jad['ym'])
    cfg = cfg or {}
    pat = patterns(hrd_rows) if hrd_rows else {}
    for p in jad['people']: ensure_pat(pat, p['pws'])
    out = []
    niks = {p['nik'] for p in jad['people']}
    prev_by = {p['nik']: p for p in prev['people']} if prev else {}
    for p in jad['people']:
        nik = p['nik']
        s = hrd_rows.get(nik) if hrd_rows else None
        if nik not in prev_by and not s:
            out.append(dict(kind='jadwal_only', nik=nik, nama=p['nama'], pws=p['pws'], default=cfg.get('new_person_pws', '(pilih)'),
                            options=['(pilih)', 'abaikan'] + sorted(k for k in pat if not k.startswith('4G')),
                            msg=f"Ada di Jadwal Kerja (PWS {p['pws']}) tetapi tidak ada di file HRD maupun revisi sebelumnya. "
                                'Pilih PWS lama (dasar HRD) untuk membandingkan, atau abaikan.'))
            continue
        if prev:
            if nik not in prev_by:
                out.append(dict(kind='new_person', nik=nik, nama=p['nama'],
                                msg='Orang baru di revisi ini; pembanding = pola HRD.'))
            elif not _same_pws(prev_by[nik]['pws'], p['pws']):
                out.append(dict(kind='pws_changed', nik=nik, nama=p['nama'],
                                msg=f"PWS berubah {prev_by[nik]['pws']} -> {p['pws']} (masuk form Perubahan Regu)."))
            continue
        if p['pws'] not in pat:
            out.append(dict(kind='unknown_pws', nik=nik, nama=p['nama'],
                            msg=f"PWS {p['pws']} tidak punya pola di file HRD - dilewati."))
            continue
        hp = s['pws'] if s and s['pws'] in pat else p['pws']
        act = p['days'][:nd]
        E0 = effective_date(act, pat[p['pws']]) if p['pws'] in pat else 1
        worked = [c for c in act[E0 - 1:] if c != 'OFF']
        codes = lambda w: {x for x in pat[w] if x != 'OFF'}
        known = codes(p['pws']) if p['pws'] in pat else set(worked)
        bad = sum(1 for c in worked if c not in known)
        if worked and bad >= len(worked) * 0.4:
            alt = s['pws'] if s and s['pws'] in pat else None
            altbad = sum(1 for c in worked if c not in codes(alt)) if alt else 99
            sug = alt if alt and altbad < bad else p['pws']
            out.append(dict(kind='pws_mismatch', nik=nik, nama=p['nama'], declared=p['pws'], suggest=sug,
                            options=sorted({p['pws'], sug}),
                            msg=f"Kode PWS di Jadwal ({p['pws']}) tidak cocok dengan pola harinya "
                                f"({bad} dari {len(worked)} hari kerja memakai kode di luar pola PWS). Disarankan: {sug}."))
    if not prev and hrd_rows is not None:
        for p in jad['people']:
            s0 = hrd_rows.get(p['nik'])
            hp = s0['pws'] if s0 else cfg.get('new_person_pws')
            if not hp or p['pws'] not in pat or ensure_pat(pat, hp) is None: continue
            act = p['days'][:nd]
            k0 = next((i for i, x in enumerate(act) if x != 'OFF'), None)
            if not k0: continue
            base, _, _ = blend_base(hp, p['pws'], act, pat, tuple(cfg.get('exempt_pws', ('05D035',))))
            exp = sum(1 for i in range(k0) if base[i] != 'OFF')
            if exp >= 1:
                out.append(dict(kind='late_start', nik=p['nik'], nama=p['nama'], start=k0 + 1, expected=exp,
                                default='ya' if exp >= 3 else 'tidak', options=['ya', 'tidak'],
                                msg=f"Libur dari tanggal 1 sampai {k0}, padahal dasar jadwal masuk {exp} hari di rentang itu. "
                                    f"Karyawan baru yang mulai masuk tanggal {k0 + 1}? 'ya' = libur sebelum tanggal itu tidak dibuatkan perubahan."))
    for p in jad['people']:
        for c in set(p['days']):
            if c != 'OFF' and c not in HOURS:
                out.append(dict(kind='unknown_code', nik=p['nik'], nama=p['nama'],
                                msg=f'Kode jam kerja {c} belum punya teks jam - kolom jam dikosongkan.'))
                break
    if hrd_rows and not prev:
        for nik, s in hrd_rows.items():
            if nik not in niks:
                out.append(dict(kind='hrd_only', nik=nik, nama=s['nama'], pws=s['pws'],
                                options=['abaikan', 'SK', 'QCSK'] if cfg.get('kind', 'sk') == 'sk' else ['abaikan', 'masukkan'],
                                msg=f"Ada di HRD (PWS {s['pws']}) tetapi tidak ada di Jadwal Kerja."))
    if prev:
        for nik, q in prev_by.items():
            if nik not in niks:
                out.append(dict(kind='removed', nik=nik, nama=q['nama'],
                                msg='Ada di revisi sebelumnya tetapi tidak ada di file baru.'))
    return out


def compute(jad, hrd_rows, prev=None, decisions=None, tgl_over=None, cfg=None, holidays=None):
    """-> (plan, snapshot).  decisions: {nik: value, 'start:nik': 'ya'|'tidak'}; tgl_over: {nik: int};
    holidays: tanggal hari nasional (libur serentak; hanya berlaku di R0)."""
    decisions = decisions or {}
    tgl_over = tgl_over or {}
    cfg = cfg or {}
    holidays = sorted({int(h) for h in (holidays or [])})
    exempt_pws = tuple(cfg.get('exempt_pws', ('05D035',)))
    ym = jad['ym']; nd = month_len(ym)
    pat = patterns(hrd_rows) if hrd_rows else {}
    for p in jad['people']: ensure_pat(pat, p['pws'])
    prev_by = {p['nik']: p for p in prev['people']} if prev else {}
    people_in = list(jad['people'])
    niks = {p['nik'] for p in people_in}
    # HRD-only yang diputuskan dimasukkan
    if hrd_rows and not prev:
        for nik, s in hrd_rows.items():
            d = decisions.get(nik)
            if nik not in niks and d in ('SK', 'QCSK', 'masukkan'):
                people_in.append(dict(no=None, nama=s['nama'], nik=nik, item=None, seksi='' if d == 'masukkan' else d, pws=s['pws'],
                                      days=list(s['days'][:nd])))
    plan_people, snap_people, skipped, starts = [], [], [], {}
    for j in people_in:
        nik = j['nik']; act = j['days'][:nd]
        s = hrd_rows.get(nik) if hrd_rows else None
        d = decisions.get(nik)
        synthetic = False
        if not s and not (prev and nik in prev_by) and d in pat:
            s = dict(pws=d, nama=j['nama'], total=None, dept=None); synthetic = True   # PWS lama dipilih pengguna
        new = d if (d in pat and not synthetic and not (prev and nik in prev_by)) else j['pws']
        regu = None; exempt = False; in_prev = bool(prev and nik in prev_by); hol_cut = 0
        if prev and nik in prev_by:
            q = prev_by[nik]
            base = q['days'][:nd]; old = q['pws']
            exempt = q.get('exempt', False)
            if not _same_pws(old, new) and new in pat:
                regu = dict(lama=old, baru=new, tgl=effective_date(act, pat[new]))
        else:
            if not s:
                skipped.append(nik); continue
            old = s['pws']
            if old in exempt_pws: old = '05D004'
            if new not in pat or old not in pat:
                skipped.append(nik); continue
            base, regu, exempt = blend_base(s['pws'], new, act, pat, exempt_pws)
            base = list(base)
            # karyawan baru di tengah bulan: libur sebelum tanggal masuk bukan perubahan jadwal
            k0 = next((i for i, x in enumerate(act) if x != 'OFF'), None)
            if decisions.get('start:' + nik) == 'ya' and k0:
                for i in range(k0): base[i] = act[i]
                starts[nik] = k0 + 1
                if regu: regu['tgl'] = max(regu['tgl'], k0 + 1)
            # hari nasional: jadwal libur serentak -> tidak dibuatkan perubahan (X -> OFF diabaikan)
            for h in holidays:
                if 1 <= h <= nd:
                    if base[h - 1] != 'OFF': hol_cut += 1
                    if act[h - 1] == 'OFF': base[h - 1] = 'OFF'
        if regu and nik in tgl_over:
            regu['tgl'] = int(tgl_over[nik])
        ents = order_entries(diffs(base, act))
        nama_hrd = s['nama'] if s else (prev_by[nik]['nama_hrd'] if nik in prev_by else j['nama'])
        plan_people.append(dict(nik=nik, nama=j['nama'], nama_hrd=nama_hrd, seksi=j['seksi'], pws=new,
                                pws_lama=old, ents=ents, regu=regu, exempt=exempt))
        if in_prev:
            target = prev_by[nik].get('target') or jad.get('target') or 22     # target bulan ini dibawa dari revisi sebelumnya
        else:
            target = ((s or {}).get('total') or jad.get('target') or 22) - hol_cut
        worked = sum(1 for x in act if x != 'OFF')
        snap_people.append(dict(nik=nik, nama=j['nama'], nama_hrd=nama_hrd, seksi=j['seksi'], pws=new,
                                days=act, exempt=exempt, target=target, worked=worked, kurang=target - worked))
    pending = [p['nik'] for p in jad['people'] if decisions.get(p['nik']) == '(pilih)']
    plan = dict(ym=ym, people=plan_people, skipped=skipped, pending=pending, holidays=holidays, starts=starts)
    snap = dict(ym=ym, people=snap_people)
    return plan, snap
