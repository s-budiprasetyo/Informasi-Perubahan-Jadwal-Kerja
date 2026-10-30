"""Penulis Excel departemen QC: sheet OFF / MASUK (serentak), R0 / R1.. (sisa), PERUBAHAN REGU."""
import copy, re, os, shutil, tempfile, zipfile
from xml.sax.saxutils import escape
from lxml import etree
from .consts import HOURS, REGU_TXT, BULAN_ID, BULAN_ABBR, ym_label
from .inf_writer import (Sheet, q, XDR, col_of, rownum, setcell, clear_row, fix_cached, fmt_dates, hrs)
from . import qc_layout

CAP = 19                    # baris per halaman (slot ke-20 dibiarkan kosong agar garis bawah halaman aman)
PAGE = 52


class QSheet(Sheet):
    def __init__(s, p1, p2):
        super().__init__(p1)
        t2 = etree.parse(p2).getroot()
        s.trows2 = {int(r.get('r')): r for r in t2.find(q('sheetData'))}
        s.tm2 = [m.get('ref') for m in t2.find(q('mergeCells'))]

    def _m(s, tm, a, b):
        return [r for r in tm if rownum(r.split(':')[0]) >= a and rownum(r.split(':')[1]) <= b]

    def blank(s, r0, last=False):
        s.drop_merges(r0, r0 + 1)
        t = 50 if last else 44
        for k in (0, 1):
            row = s.put_row(r0 + k, s.trows[t + k]); clear_row(row)
        for ref in s._m(s.tm, t, t + 1): s.add_merge(s.shifted(ref, r0 - t))

    def filled(s, r0, kind, vals, many):
        s.drop_merges(r0, r0 + 1)
        src, tm = (s.trows2, s.tm2) if kind == 'ox' else (s.trows, s.tm)
        for k in (0, 1):
            row = s.put_row(r0 + k, src[12 + k]); clear_row(row)
            if kind == 'xy':
                for c in row:
                    if col_of(c.get('r')) in ('S', 'T', 'U'):
                        c2 = next(x for x in s.trows2[12 + k] if col_of(x.get('r')) == col_of(c.get('r')))
                        c.set('s', c2.get('s'))
            for (col, rr), v in vals.items():
                if rr == k: setcell(row, col, v)
        for ref in s._m(tm, 12, 13): s.add_merge(s.shifted(ref, r0 - 12))
        if kind == 'xy':
            for k in (0, 1): s.add_merge(f'S{r0 + k}:U{r0 + k}')
        if many:
            for mm in list(s.mc):
                if mm.get('ref') in (f'J{r0}:O{r0}', f'J{r0 + 1}:O{r0 + 1}'): s.mc.remove(mm)
            s.add_merge(f'J{r0}:O{r0 + 1}')


def _drawing_pages(src, n):
    dt = etree.parse(src); dr = dt.getroot()
    anc = dr.findall('{%s}twoCellAnchor' % XDR)
    anc.sort(key=lambda a: int(a.find('{%s}from' % XDR).find('{%s}row' % XDR).text))
    for a in anc[n:]: dr.remove(a)
    k = len(anc)
    while k < n:
        nw = copy.deepcopy(anc[-1] if k <= len(anc) else anc[-1])
        for tag in ('from', 'to'):
            e = nw.find('{%s}%s' % (XDR, tag)).find('{%s}row' % XDR); e.text = str(int(e.text) + PAGE * (k - len(anc) + 1))
        nw.find('.//{%s}cNvPr' % XDR).set('id', str(20 + k)); dr.append(nw); k += 1
    return etree.tostring(dt, xml_declaration=True, encoding='UTF-8', standalone=True)


def _list_sheet(W, D, pages, date_txt, BULAN, alasan):
    """pages: list of list of rows (nik, nama, a, b, ds). Return (xml, drawing, last_row)."""
    s = QSheet(W('sheet1.xml'), W('sheet2.xml'))
    n = max(1, len(pages)); starts = [12 + PAGE * i for i in range(n)]
    for i in range(3, n): s.clone_block(105, 156, 105 + PAGE * (i - 2))
    if n < 3: s.truncate_after(PAGE * n)
    for pi in range(n):
        page = pages[pi] if pi < len(pages) else []
        for k in range(20):
            r0 = starts[pi] + 2 * k
            if k >= len(page) or k >= CAP:
                s.blank(r0, last=(k == 19)); continue
            nik, nm, a, b, ds = page[k]
            many = len(ds) > 2
            v = {('A', 0): k + 1, ('B', 0): int(nik), ('D', 0): nm, ('V', 0): alasan}
            if many: v[('J', 0)] = fmt_dates(ds, BULAN, False)
            else:
                v[('J', 0)] = f'{ds[0]} {BULAN}' if len(ds) == 1 else f'{ds[0]} & {ds[1]} {BULAN}'
                v[('J', 1)] = f' ( {len(ds)} Hari )'
            if b == 'OFF': kind = 'xo'; v.update({('P', 0): a, ('P', 1): hrs(a), ('S', 0): 'OFF'})
            elif a == 'OFF': kind = 'ox'; v.update({('P', 0): 'OFF', ('S', 0): b, ('S', 1): hrs(b)})
            else: kind = 'xy'; v.update({('P', 0): a, ('P', 1): hrs(a), ('S', 0): b, ('S', 1): hrs(b)})
            s.filled(r0, kind, v, many)
    setcell(s.rows[3], 'P', date_txt)
    s.finish(date_txt); s.set_breaks([PAGE * i for i in range(1, n)])
    return (etree.tostring(s.tree, xml_declaration=True, encoding='UTF-8', standalone=True),
            _drawing_pages(D('drawing1.xml'), n), PAGE * n)


def _regu_sheet(W, D, people, regu_name, BULAN, date_txt):
    rt = etree.parse(W('sheet3.xml')); root = rt.getroot()
    rows = {int(r.get('r')): r for r in root.find(q('sheetData'))}
    for r in list(range(12, 58)) + list(range(69, 115)):
        if r in rows: clear_row(rows[r])
    ps = [p for p in people if p['regu']]
    if len(ps) > 46: raise ValueError('Form PERUBAHAN REGU hanya muat 46 orang (2 halaman).')
    for i, p in enumerate(ps):
        r0 = (12 + 2 * i) if i < 23 else (69 + 2 * (i - 23)); g = p['regu']
        a, b = rows[r0], rows[r0 + 1]
        for col, v in [('A', i + 1), ('B', int(p['nik'])), ('C', p['nama_hrd'] if regu_name == 'hrd' else p['nama']),
                       ('I', 'QC'), ('J', g['lama']), ('L', g['baru']), ('N', f"{g['tgl']} {BULAN}"), ('R', 'Kebutuhan Produksi')]:
            setcell(a, col, v)
        for col, v in [('J', REGU_TXT.get(g['lama'], '')), ('L', REGU_TXT.get(g['baru'], '')), ('N', '( Seterusnya )')]:
            setcell(b, col, 'FRE1' if (col == 'J' and g['lama'] == '05D035') else v)
    head = next(x for x in root.find(q('sheetData')) if int(x.get('r')) == 3)
    setcell(head, 'O', date_txt); fix_cached(root, date_txt)
    sv = root.find(q('sheetViews'))[0]; sv.set('topLeftCell', 'A1'); sv.attrib.pop('tabSelected', None)
    for sel in sv.findall(q('selection')): sel.set('activeCell', 'A1'); sel.set('sqref', 'A1')
    return (etree.tostring(rt, xml_declaration=True, encoding='UTF-8', standalone=True),
            open(D('drawing3.xml'), 'rb').read(), 115 if len(ps) > 23 else 58, len(ps))


def _safe(name, used):
    n = re.sub(r'[\[\]:*?/\\]', '', name)[:31] or 'Sheet'
    base, i = n, 2
    while n.lower() in used:
        n = (base[:28] + f' {i}'); i += 1
    used.add(n.lower()); return n


def assemble(src, sheets, out_path):
    """sheets: [dict(name, xml, drawing, print_area_last, regu=bool)]"""
    tmp = tempfile.mkdtemp(); out = os.path.join(tmp, 'o'); shutil.copytree(src, out)
    for sub in ('xl/worksheets', 'xl/drawings'):
        shutil.rmtree(os.path.join(out, sub), ignore_errors=True)
    for f in ('xl/calcChain.xml',):
        if os.path.exists(os.path.join(out, f)): os.remove(os.path.join(out, f))
    os.makedirs(os.path.join(out, 'xl/worksheets/_rels')); os.makedirs(os.path.join(out, 'xl/drawings'))
    ct = open(os.path.join(out, '[Content_Types].xml'), encoding='utf-8').read()
    ct = re.sub(r'<Override PartName="/xl/(worksheets|drawings)/[^"]*"[^>]*/>', '', ct)
    ov = ''
    for i, sh in enumerate(sheets, 1):
        open(os.path.join(out, f'xl/worksheets/sheet{i}.xml'), 'wb').write(sh['xml'])
        open(os.path.join(out, f'xl/drawings/drawing{i}.xml'), 'wb').write(sh['drawing'])
        open(os.path.join(out, f'xl/worksheets/_rels/sheet{i}.xml.rels'), 'w', encoding='utf-8').write(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{i}.xml"/></Relationships>')
        ov += (f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
               f'<Override PartName="/xl/drawings/drawing{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')
    open(os.path.join(out, '[Content_Types].xml'), 'w', encoding='utf-8').write(ct.replace('</Types>', ov + '</Types>'))
    rp = os.path.join(out, 'xl/_rels/workbook.xml.rels'); r = open(rp, encoding='utf-8').read()
    r = re.sub(r'<Relationship Id="[^"]*" Type="[^"]*/worksheet"[^>]*/>', '', r)
    add = ''.join(f'<Relationship Id="rIdS{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
                  for i in range(1, len(sheets) + 1))
    open(rp, 'w', encoding='utf-8').write(r.replace('</Relationships>', add + '</Relationships>'))
    wp = os.path.join(out, 'xl/workbook.xml'); w = open(wp, encoding='utf-8').read()
    shx = ''.join(f'<sheet name="{escape(sh["name"])}" sheetId="{i}" r:id="rIdS{i}"/>' for i, sh in enumerate(sheets, 1))
    w = re.sub(r'<sheets>.*?</sheets>', f'<sheets>{shx}</sheets>', w, flags=re.S)
    dn = ''.join(f'<definedName name="_xlnm.Print_Area" localSheetId="{i - 1}">\'{escape(sh["name"])}\'!'
                 + (f'$A$2:$U${sh["last"]}' if sh['regu'] else f'$A$1:$W${sh["last"]}') + '</definedName>'
                 for i, sh in enumerate(sheets, 1))
    w = re.sub(r'<definedNames>.*?</definedNames>', f'<definedNames>{dn}</definedNames>', w, flags=re.S)
    w = re.sub(r'<calcPr[^>]*/>', '<calcPr calcId="144525" concurrentCalc="0" fullCalcOnLoad="1"/>', w)
    w = w.replace('activeTab="2"', 'activeTab="0"')
    open(wp, 'w', encoding='utf-8').write(w)
    N = len(sheets)
    names = ''.join(f'<vt:lpstr>{escape(sh["name"])}</vt:lpstr>' for sh in sheets)
    pas = ''.join(f"<vt:lpstr>'{escape(sh['name'])}'!Print_Area</vt:lpstr>" for sh in sheets)
    app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
           'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><HeadingPairs>'
           f'<vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{N}</vt:i4></vt:variant>'
           f'<vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>{N}</vt:i4></vt:variant></vt:vector></HeadingPairs>'
           f'<TitlesOfParts><vt:vector size="{2 * N}" baseType="lpstr">{names}{pas}</vt:vector></TitlesOfParts></Properties>')
    open(os.path.join(out, 'docProps/app.xml'), 'w', encoding='utf-8').write(app)
    if os.path.exists(out_path): os.remove(out_path)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(out, '[Content_Types].xml'), '[Content_Types].xml')
        for dp, dn_, fn in os.walk(out):
            for f in fn:
                full = os.path.join(dp, f); arc = os.path.relpath(full, out)
                if arc != '[Content_Types].xml': z.write(full, arc)
    shutil.rmtree(tmp, ignore_errors=True)


def _pack(rows, cap=CAP):
    pages = [[]]
    for r in rows:
        if len(pages[-1]) >= cap: pages.append([])
        pages[-1].append(r)
    return pages


def render_qc(plan, rev, gen_date, template, out_path, cfg, teams=None, bulk_pct=None):
    ym = plan['ym']; y, m = map(int, ym.split('-')); BULAN = ym_label(ym)
    date_txt = f'GRESIK, {gen_date.day} {BULAN_ID[gen_date.month - 1]} {gen_date.year}'
    teams = teams or cfg.get('teams', [])
    pct = cfg.get('bulk_pct', 50) if bulk_pct is None else bulk_pct
    alasan = 'Kebutuhan Produksi'
    tmp = tempfile.mkdtemp(); src = os.path.join(tmp, 'src')
    with zipfile.ZipFile(template) as z: z.extractall(src)
    W = lambda f: os.path.join(src, 'xl', 'worksheets', f); D = lambda f: os.path.join(src, 'xl', 'drawings', f)
    by = {p['nik']: p for p in plan['people']}
    tm = qc_layout.order_people(plan['people'], teams)
    sheets, used, info = [], set(), {}

    def add(name, xml, drw, last, regu=False):
        sheets.append(dict(name=_safe(name, used), xml=xml, drawing=drw, last=last, regu=regu))

    if rev == 0:
        bulks, rest = qc_layout.split_bulk(plan, pct)
        multi = {k: sum(1 for b in bulks if b['kind'] == k) > 1 for k in ('off', 'on')}
        for b in bulks:
            pages = []
            for team in tm:
                rows = [(p['nik'], p['nama'], *( (e[0], e[1]) ), [b['date']]) for p in team if p['nik'] in b['niks']
                        for e in p['ents'] if b['date'] in e[2] and ((e[1] == 'OFF') if b['kind'] == 'off' else (e[0] == 'OFF'))]
                if rows: pages.extend(_pack(rows))
            base = 'OFF' if b['kind'] == 'off' else 'MASUK'
            nm = f"{base} {b['date']} {BULAN_ABBR[m - 1]}" if multi[b['kind']] else base
            x, d, last = _list_sheet(W, D, pages, date_txt, BULAN, alasan); add(nm, x, d, last)
        rows = [(p['nik'], p['nama'], e[0], e[1], e[2]) for p in plan['people'] for e in rest.get(p['nik'], [])]
        if rows or not bulks:
            x, d, last = _list_sheet(W, D, _pack(rows), date_txt, BULAN, alasan); add('R0', x, d, last)
        info['bulk'] = [(b['kind'], b['date'], len(b['niks'])) for b in bulks]; info['rest'] = len(rows)
    else:
        rows = [(p['nik'], p['nama'], e[0], e[1], e[2]) for p in plan['people'] for e in p['ents']]
        x, d, last = _list_sheet(W, D, _pack(rows), date_txt, BULAN, alasan); add(f'R{rev}', x, d, last)
        info['rest'] = len(rows)
    x, d, last, nr = _regu_sheet(W, D, plan['people'], cfg.get('regu_name', 'jadwal'), BULAN, date_txt)
    info['regu'] = nr
    if nr or rev == 0 and False: add('PERUBAHAN REGU', x, d, last, regu=True)
    assemble(src, sheets, out_path)
    shutil.rmtree(tmp, ignore_errors=True)
    info['sheets'] = [s['name'] for s in sheets]
    return info
