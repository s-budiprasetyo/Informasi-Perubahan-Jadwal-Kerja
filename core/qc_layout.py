"""Tata letak form departemen QC: tim (halaman) dan pemisahan perubahan serentak."""
import math
from collections import defaultdict


def order_people(people, teams):
    """people: plan['people'] (urutan Jadwal Kerja). teams: [[nik,...],...].
    Mengembalikan list tim; tiap tim = list orang. NIK tanpa tim ikut tim orang di atasnya
    (blok berurutan di Jadwal Kerja)."""
    where = {}
    for ti, t in enumerate(teams):
        for oi, nik in enumerate(t): where[str(nik)] = (ti, oi)
    cur = 0; groups = defaultdict(list)
    for idx, p in enumerate(people):
        if p['nik'] in where:
            ti, oi = where[p['nik']]; cur = ti; key = (oi, idx)
        else:
            ti = cur; key = (10 ** 6, idx)
        groups[ti].append((key, p))
    return [[p for _, p in sorted(groups[t], key=lambda x: x[0])] for t in sorted(groups)]


def split_bulk(plan, pct=50):
    """-> (bulks, rest)
    bulks: [dict(kind='off'|'on', date=int, niks=set)] tanggal di mana >= pct% orang
           sama-sama X->OFF (atau OFF->X); rest: {nik: [[a,b,[tgl]],...]} sisa tidak teratur."""
    n = len(plan['people']); need = max(5, math.ceil(n * pct / 100))
    cnt = defaultdict(set)
    for p in plan['people']:
        for a, b, ds in p['ents']:
            kind = 'off' if b == 'OFF' else 'on' if a == 'OFF' else None
            if kind:
                for d in ds: cnt[(kind, d)].add(p['nik'])
    keys = sorted((k for k, v in cnt.items() if len(v) >= need), key=lambda k: (k[0] != 'off', k[1]))
    bulks = [dict(kind=k, date=d, niks=cnt[(k, d)]) for k, d in keys]
    kd = {(b['kind'], b['date']) for b in bulks}
    rest = {}
    for p in plan['people']:
        out = []
        for a, b, ds in p['ents']:
            kind = 'off' if b == 'OFF' else 'on' if a == 'OFF' else None
            left = [d for d in ds if (kind, d) not in kd]
            if left: out.append([a, b, left])
        if out: rest[p['nik']] = out
    return bulks, rest
