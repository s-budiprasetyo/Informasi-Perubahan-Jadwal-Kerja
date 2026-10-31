"""Penulis Excel 'Informasi Perubahan Jadwal Jam Kerja' langsung di OOXML template
(agar gambar/tanda tangan & format asli tetap utuh)."""
import copy, re, os, shutil, tempfile, zipfile, datetime, math
from lxml import etree
from .consts import HOURS, REGU_TXT, DAY_PWS, BULAN_ID, BULAN_ABBR, ym_label, sheet_suffix

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
q = lambda t: '{%s}%s' % (NS, t)
col_of = lambda ref: re.match(r'[A-Z]+', ref).group(0)
rownum = lambda ref: int(re.sub('[A-Z]', '', ref))
hrs = lambda c: HOURS.get(c, '')


def setcell(row, col, val):
    for c in row:
        if col_of(c.get('r')) == col:
            for ch in list(c): c.remove(ch)
            c.attrib.pop('t', None)
            if val is None: return
            if isinstance(val, (int, float)):
                etree.SubElement(c, q('v')).text = str(val)
            else:
                c.set('t', 'inlineStr'); i = etree.SubElement(c, q('is')); t = etree.SubElement(i, q('t'))
                t.text = str(val); t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            return
    raise KeyError(col)


def clear_row(row):
    for c in row:
        for ch in list(c): c.remove(ch)
        c.attrib.pop('t', None)


def fmt_dates(ds, bln, qc):
    n = len(ds)
    if n == 1: return f'{ds[0]} {bln}'
    if qc and n >= 4 and ds == list(range(ds[0], ds[0] + n)): return f'{ds[0]} s/d {ds[-1]} {bln}'
    txt = ', '.join(map(str, ds[:-1])) + f' & {ds[-1]}'
    return txt + (' ' + bln if qc else '\n' + bln)


def fix_cached(root, date_txt):
    """Sel rumus yang menunjuk tanggal header (=$P$3, =O3) diberi nilai cache baru."""
    for c in root.iter(q('c')):
        f = c.find(q('f'))
        if f is not None and (f.text or '').replace('$', '') in ('P3', 'O3'):
            for v in c.findall(q('v')): c.remove(v)
            c.set('t', 'str'); etree.SubElement(c, q('v')).text = date_txt


class Sheet:
    def __init__(s, path):
        s.tpl = etree.parse(path).getroot(); s.tree = etree.parse(path); s.root = s.tree.getroot()
        s.sd = s.root.find(q('sheetData')); s.mc = s.root.find(q('mergeCells'))
        s.rows = {int(r.get('r')): r for r in s.sd}; s.trows = {int(r.get('r')): r for r in s.tpl.find(q('sheetData'))}
        s.tm = [m.get('ref') for m in s.tpl.find(q('mergeCells'))]

    def tmerges(s, a, b):
        return [ref for ref in s.tm if rownum(ref.split(':')[0]) >= a and rownum(ref.split(':')[1]) <= b]

    def drop_merges(s, a, b):
        for m in list(s.mc):
            x, y = m.get('ref').split(':')
            if rownum(x) >= a and rownum(y) <= b: s.mc.remove(m)

    def add_merge(s, ref):
        if not any(m.get('ref') == ref for m in s.mc):
            etree.SubElement(s.mc, q('mergeCell')).set('ref', ref)

    @staticmethod
    def shifted(ref, off):
        x, y = ref.split(':'); f = lambda z: col_of(z) + str(rownum(z) + off)
        return f(x) + ':' + f(y)

    def put_row(s, r, trow):
        trow = copy.deepcopy(trow); trow.set('r', str(r))
        for c in trow: c.set('r', col_of(c.get('r')) + str(r))
        if r in s.rows:
            old = s.rows[r]; old.getparent().replace(old, trow)
        else:
            s.sd.append(trow)
        s.rows[r] = trow; return trow

    def clone_block(s, a, b, dst):
        off = dst - a
        s.drop_merges(dst, dst + (b - a))
        for r in range(a, b + 1):
            row = s.put_row(r + off, s.trows[r])
            for c in row:
                f = c.find(q('f'))
                if f is not None: f.text = 'P3'
        for ref in s.tmerges(a, b): s.add_merge(s.shifted(ref, off))

    def truncate_after(s, last):
        for r in [r for r in s.rows if r > last]:
            s.sd.remove(s.rows[r]); del s.rows[r]
        s.drop_merges(last + 1, 10 ** 6)

    def slot(s, r0, tpl, vals):
        s.drop_merges(r0, r0 + 1)
        for k in (0, 1):
            row = s.put_row(r0 + k, s.trows[tpl + k])
            for c in row:
                for ch in list(c): c.remove(ch)
                c.attrib.pop('t', None)
            for (col, rr), v in vals.items():
                if rr == k: setcell(row, col, v)
        for ref in s.tmerges(tpl, tpl + 1): s.add_merge(s.shifted(ref, r0 - tpl))

    def finish(s, date_txt=None):
        if date_txt: fix_cached(s.root, date_txt)
        s.mc.set('count', str(len(s.mc)))
        s.root.find(q('dimension')).set('ref', f'A1:AC{max(s.rows)}')
        sv = s.root.find(q('sheetViews'))[0]; sv.set('topLeftCell', 'A1'); sv.attrib.pop('tabSelected', None)
        for sel in sv.findall(q('selection')): sel.set('activeCell', 'A1'); sel.set('sqref', 'A1')

    def set_breaks(s, ids):
        rb = s.root.find(q('rowBreaks'))
        if rb is not None: s.root.remove(rb)
        if not ids: return
        rb = etree.Element(q('rowBreaks'))
        for i in ids: etree.SubElement(rb, q('brk'), id=str(i), max='22', man='1')
        rb.set('count', str(len(ids))); rb.set('manualBreakCount', str(len(ids)))
        # urutan skema OOXML: ... pageSetup, headerFooter, rowBreaks, ... drawing  (Excel menolak urutan lain)
        hf = s.root.find(q('headerFooter'))
        (hf if hf is not None else s.root.find(q('pageSetup'))).addnext(rb)


def split_dates(ds, maxd, balanced):
    if len(ds) <= maxd: return [ds]
    if balanced:
        k = math.ceil(len(ds) / maxd); sz = math.ceil(len(ds) / k)
        return [ds[i:i + sz] for i in range(0, len(ds), sz)]
    return [ds[i:i + maxd] for i in range(0, len(ds), maxd)]


def person_rows(p, maxd, balanced):
    out = []
    for a, b, ds in p['ents']:
        for chunk in split_dates(ds, maxd, balanced):
            out.append((p['nik'], p['nama'], p['nama_hrd'], a, b, chunk))
    return out


def paginate(groups, cap, last_ok_offx):
    pages = []
    for blocks in groups:
        if not blocks: continue
        pages.append([])
        for ents in blocks:
            def fits(page):
                end = len(page) + len(ents)
                return end <= cap - 1 or (end == cap and last_ok_offx and ents[-1][3] == 'OFF')
            if not fits(pages[-1]): pages.append([])
            pages[-1].extend(ents)
    return pages or [[]]


def _drawing_fix(path, n, keep_n, clone_from=None):
    dt = etree.parse(path); dr = dt.getroot()
    anc = dr.findall('{%s}twoCellAnchor' % XDR)
    anc.sort(key=lambda a: int(a.find('{%s}from' % XDR).find('{%s}row' % XDR).text))
    if n > len(anc) and clone_from is not None:
        nw = copy.deepcopy(anc[clone_from])
        for tag in ('from', 'to'):
            e = nw.find('{%s}%s' % (XDR, tag)).find('{%s}row' % XDR); e.text = str(int(e.text) + 52)
        nw.find('.//{%s}cNvPr' % XDR).set('id', '9'); dr.append(nw)
    else:
        for a in anc[n:]: dr.remove(a)
    dt.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)


def render_inf(plan, rev, gen_date, template, out_path, alasan_kiln='Kebutuhan Rencana Produksi'):
    ym = plan['ym']; y, m = map(int, ym.split('-'))
    BULAN = ym_label(ym); BLN_QC = f'{BULAN_ABBR[m - 1]} {y}'
    kiln = [p for p in plan['people'] if p['seksi'] != 'QCSK']
    qc = [p for p in plan['people'] if p['seksi'] == 'QCSK']
    tmp = tempfile.mkdtemp()
    SRC = os.path.join(tmp, 'src'); OUTDIR = os.path.join(tmp, 'out')
    with zipfile.ZipFile(template) as z: z.extractall(SRC)
    shutil.copytree(SRC, OUTDIR)
    W = lambda sub: os.path.join(SRC, 'xl', 'worksheets', sub)
    O = lambda sub: os.path.join(OUTDIR, 'xl', 'worksheets', sub)
    date_txt = f'GRESIK, {gen_date.day} {BULAN_ID[gen_date.month - 1]} {gen_date.year}'
    info = {}

    def head_date(sheet_root, cell):
        r = int(re.sub('[A-Z]', '', cell))
        row = next(x for x in sheet_root.find(q('sheetData')) if int(x.get('r')) == r)
        setcell(row, col_of(cell), date_txt)

    isday = lambda p: p['pws'] in DAY_PWS
    # ---------------- SULTON (kiln)
    ks = Sheet(W('sheet2.xml'))
    day = [person_rows(p, 8, False) for p in kiln if isday(p) and p['ents']]
    shf = [person_rows(p, 8, False) for p in kiln if not isday(p) and p['ents']]
    pages = paginate([day, shf], 20, True)
    if len(pages) > 4: raise ValueError(f'Form SULTON butuh {len(pages)} halaman (maks 4).')
    npg = len(pages); info['kiln_pages'] = [len(p) for p in pages]
    starts = [12, 64, 116, 168]
    if npg == 4: ks.clone_block(105, 156, 157)
    elif npg < 3: ks.truncate_after(52 if npg == 1 else 104)
    for pi, page in enumerate(pages):
        for k in range(20):
            r0 = starts[pi] + 2 * k
            if k < len(page):
                nik, nm, full, a, b, ds = page[k]
                v = {('A', 0): k + 1, ('B', 0): int(nik), ('D', 0): full, ('J', 0): fmt_dates(ds, BULAN, False),
                     ('V', 0): alasan_kiln}
                tpl = 50 if k == 19 else (12 if b == 'OFF' else 14 if a == 'OFF' else 72)
                if b == 'OFF': v.update({('P', 0): a, ('P', 1): hrs(a), ('S', 0): 'OFF'})
                elif a == 'OFF': v.update({('P', 0): 'OFF', ('S', 0): b, ('S', 1): hrs(b)})
                else: v.update({('P', 0): a, ('P', 1): hrs(a), ('S', 0): b, ('S', 1): hrs(b)})
                ks.slot(r0, tpl, v)
            else:
                ks.slot(r0, 154 if k == 19 else 140, {})
    ks.finish(date_txt); ks.set_breaks([52 * i for i in range(1, npg)])
    head_date(ks.root, 'P3')
    ks.tree.write(O('sheet2.xml'), xml_declaration=True, encoding='UTF-8', standalone=True)
    _drawing_fix(os.path.join(OUTDIR, 'xl/drawings/drawing2.xml'), npg, npg, clone_from=2)
    kiln_last = 52 * npg
    # ---------------- SATRIO (QC)
    qs = Sheet(W('sheet1.xml'))
    day = [person_rows(p, 8, True) for p in qc if isday(p) and p['ents']]
    shf = [person_rows(p, 8, True) for p in qc if not isday(p) and p['ents']]
    qpages = paginate([day, shf], 17, False)
    if len(qpages) > 4: raise ValueError(f'Form SATRIO butuh {len(qpages)} halaman (maks 4).')
    nq = len(qpages); info['qc_pages'] = [len(p) for p in qpages]
    qstarts = [12, 58, 103, 148]; page_end = [47, 92, 137, 182]
    if nq < 4: qs.truncate_after(page_end[nq - 1])
    for pi, page in enumerate(qpages):
        for k in range(17):
            r0 = qstarts[pi] + 2 * k
            if k < len(page):
                nik, nm, full, a, b, ds = page[k]
                many = len(ds) > 2
                v = {('A', 0): k + 1, ('B', 0): int(nik), ('D', 0): nm, ('V', 0): 'Kebutuhan Pemenuhan Produksi'}
                if many: v[('J', 0)] = fmt_dates(ds, BULAN, False)
                else: v[('J', 0)] = fmt_dates(ds, BLN_QC, True); v[('J', 1)] = f'( {len(ds)} Hari )'
                if b == 'OFF': tpl = 12; v.update({('P', 0): a, ('P', 1): hrs(a), ('S', 0): 'OFF'})
                elif a == 'OFF': tpl = 14; v.update({('P', 0): 'OFF', ('S', 0): b, ('S', 1): hrs(b)})
                else: tpl = 119; v.update({('P', 0): a, ('P', 1): hrs(a), ('S', 0): b, ('S', 1): hrs(b)})
                qs.slot(r0, tpl, v)
                if many:
                    for mm in list(qs.mc):
                        if mm.get('ref') in (f'J{r0}:O{r0}', f'J{r0 + 1}:O{r0 + 1}'): qs.mc.remove(mm)
                    qs.add_merge(f'J{r0}:O{r0 + 1}')
            else:
                qs.slot(r0, 180 if k == 16 else 170, {})
    qs.finish(date_txt); qs.set_breaks([page_end[i - 1] for i in range(1, nq)])
    head_date(qs.root, 'P3')
    qs.tree.write(O('sheet1.xml'), xml_declaration=True, encoding='UTF-8', standalone=True)
    _drawing_fix(os.path.join(OUTDIR, 'xl/drawings/drawing1.xml'), nq, nq)
    q_last = page_end[nq - 1]
    # ---------------- PERUBAHAN REGU
    rt = etree.parse(W('sheet4.xml')); root = rt.getroot()
    rows = {int(r.get('r')): r for r in root.find(q('sheetData'))}
    for r in list(range(12, 58)) + list(range(69, 115)):
        if r in rows: clear_row(rows[r])

    def regu_rows(ps, first):
        n = 0
        for p in ps:
            g = p['regu']
            if not g: continue
            if n >= 23: raise ValueError('Form PERUBAHAN REGU penuh (maks 23 baris per seksi).')
            r0 = first + 2 * n; n += 1
            a, b = rows[r0], rows[r0 + 1]
            for col, v in [('A', n), ('B', int(p['nik'])), ('C', p['nama_hrd']), ('I', 'QC SK' if p['seksi'] == 'QCSK' else 'SK'),
                           ('J', g['lama']), ('L', g['baru']), ('N', f"{g['tgl']} {BULAN}"), ('R', 'Kebutuhan Produksi')]:
                setcell(a, col, v)
            for col, v in [('J', REGU_TXT.get(g['lama'], '')), ('L', REGU_TXT.get(g['baru'], '')), ('N', '(Seterusnya)')]:
                setcell(b, col, v)
        return n
    nk = regu_rows(kiln, 12); nqc = regu_rows(qc, 69)
    info['regu'] = (nk, nqc)
    head_date(root, 'O3'); fix_cached(root, date_txt)
    sv = root.find(q('sheetViews'))[0]; sv.set('topLeftCell', 'A1'); sv.attrib.pop('tabSelected', None)
    for sel in sv.findall(q('selection')): sel.set('activeCell', 'A1'); sel.set('sqref', 'A1')
    rt.write(O('sheet4.xml'), xml_declaration=True, encoding='UTF-8', standalone=True)
    reg_last = 115 if nqc else 58
    # ---------------- paket
    for f in ['xl/worksheets/sheet3.xml', 'xl/worksheets/_rels/sheet3.xml.rels', 'xl/drawings/drawing3.xml',
              'xl/printerSettings/printerSettings3.bin', 'xl/calcChain.xml']:
        fp = os.path.join(OUTDIR, f)
        if os.path.exists(fp): os.remove(fp)
    rd = lambda p: open(os.path.join(OUTDIR, p), encoding='utf-8').read()
    def wr(p, s): open(os.path.join(OUTDIR, p), 'w', encoding='utf-8').write(s)
    s = rd('[Content_Types].xml')
    for part in ['worksheets/sheet3.xml', 'drawings/drawing3.xml', 'calcChain.xml']:
        s = re.sub(r'<Override PartName="/xl/%s"[^>]*/>' % re.escape(part), '', s)
    wr('[Content_Types].xml', s)
    s = rd('xl/_rels/workbook.xml.rels')
    for rid in ['rId3', 'rId8']: s = re.sub(r'<Relationship Id="%s"[^>]*/>' % rid, '', s)
    wr('xl/_rels/workbook.xml.rels', s)
    suf = sheet_suffix(rev)
    N1, N2, N3 = 'SATRIO' + suf, 'SULTON' + suf, 'PERUBAHAN REGU' + suf
    s = rd('xl/workbook.xml')
    s = re.sub(r'<sheet name="SULTON R1"[^>]*/>', '', s)
    s = s.replace('<sheet name="SATRIO"', f'<sheet name="{N1}"').replace('<sheet name="SULTON"', f'<sheet name="{N2}"') \
         .replace('<sheet name="PERUBAHAN REGU"', f'<sheet name="{N3}"')
    s = re.sub(r'<definedNames>.*?</definedNames>',
               f"<definedNames><definedName name=\"_xlnm.Print_Area\" localSheetId=\"2\">'{N3}'!$A$2:$U${reg_last}</definedName>"
               f"<definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">'{N1}'!$A$2:$W${q_last - 1}</definedName>"
               f"<definedName name=\"_xlnm.Print_Area\" localSheetId=\"1\">'{N2}'!$A$1:$W${kiln_last}</definedName></definedNames>", s, flags=re.S)
    s = s.replace('activeTab="2"', 'activeTab="0"').replace('<calcPr calcId="191029" concurrentCalc="0"/>',
                                                           '<calcPr calcId="191029" concurrentCalc="0" fullCalcOnLoad="1"/>')
    wr('xl/workbook.xml', s)
    s = rd('docProps/app.xml')
    s = s.replace('<vt:i4>4</vt:i4></vt:variant><vt:variant><vt:lpstr>Named', '<vt:i4>3</vt:i4></vt:variant><vt:variant><vt:lpstr>Named', 1)
    s = re.sub(r'(Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>)4', '\\g<1>3', s)
    s = re.sub(r'<TitlesOfParts>.*?</TitlesOfParts>',
               f"<TitlesOfParts><vt:vector size=\"6\" baseType=\"lpstr\"><vt:lpstr>{N1}</vt:lpstr><vt:lpstr>{N2}</vt:lpstr><vt:lpstr>{N3}</vt:lpstr>"
               f"<vt:lpstr>'{N3}'!Print_Area</vt:lpstr><vt:lpstr>'{N1}'!Print_Area</vt:lpstr><vt:lpstr>'{N2}'!Print_Area</vt:lpstr></vt:vector></TitlesOfParts>", s, flags=re.S)
    wr('docProps/app.xml', s)
    if os.path.exists(out_path): os.remove(out_path)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(OUTDIR, '[Content_Types].xml'), '[Content_Types].xml')
        for dp, dn, fn in os.walk(OUTDIR):
            for f in fn:
                full = os.path.join(dp, f); arc = os.path.relpath(full, OUTDIR)
                if arc != '[Content_Types].xml': z.write(full, arc)
    shutil.rmtree(tmp, ignore_errors=True)
    info['sheets'] = [N1, N2, N3]
    return info
