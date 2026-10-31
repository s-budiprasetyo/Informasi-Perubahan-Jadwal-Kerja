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


def split_kinds(plan):
    """Pisahkan perubahan R0 menurut jenisnya (mengikuti form buatan manual):
    off  = X -> OFF   -> sheet OFF   (tiap orang satu baris per tanggal/kelompok tanggal, tanggal boleh beda-beda)
    on   = OFF -> X   -> sheet MASUK
    rest = X -> Y (tukar kode jam, keduanya masuk) -> sheet R0 (tidak teratur)
    Mengembalikan dict kind -> {nik: [[a,b,[tgl]],...]}."""
    out = {'off': {}, 'on': {}, 'rest': {}}
    for p in plan['people']:
        for a, b, ds in p['ents']:
            k = 'off' if b == 'OFF' else 'on' if a == 'OFF' else 'rest'
            out[k].setdefault(p['nik'], []).append([a, b, list(ds)])
    return out
