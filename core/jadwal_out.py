"""Berkas 'Jadwal Kerja' hasil render: salinan file Jadwal Kerja yang diunggah (format & gambar asli utuh),
dengan 'Revisi : n' dan tanggal 'Gresik, ...' diperbarui."""
import io, re, zipfile
import openpyxl
from lxml import etree
from .consts import BULAN_ID

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def _locate(data):
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=False)
    for ws in wb.worksheets:
        rev = date = None
        for row in ws.iter_rows(min_row=1, max_row=12):
            for c in row:
                v = str(c.value or '').strip()
                if v.lower().startswith('revisi') and rev is None: rev = c.coordinate
                if v.lower().startswith('gresik,') and date is None: date = c.coordinate
        if rev or date:
            return ws.title, rev, date
    raise ValueError('Sel "Revisi : n" / "Gresik, ..." tidak ditemukan di Jadwal Kerja.')


def _setcell(root, ref, text):
    for c in root.iter('{%s}c' % NS):
        if c.get('r') == ref:
            for ch in list(c): c.remove(ch)
            c.set('t', 'inlineStr')
            t = etree.SubElement(etree.SubElement(c, '{%s}is' % NS), '{%s}t' % NS); t.text = text
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            return True
    return False


def make_jadwal(data, rev, gen_date):
    title, rev_ref, date_ref = _locate(data)
    zin = zipfile.ZipFile(io.BytesIO(data))
    wbx = etree.fromstring(zin.read('xl/workbook.xml'))
    rid = next(s.get('{%s}id' % RNS) for s in wbx.iter('{%s}sheet' % NS) if s.get('name') == title)
    rels = etree.fromstring(zin.read('xl/_rels/workbook.xml.rels'))
    target = next(r.get('Target') for r in rels if r.get('Id') == rid)
    part = 'xl/' + target.lstrip('/').replace('xl/', '', 1) if not target.startswith('/') else target.lstrip('/')
    root = etree.fromstring(zin.read(part))
    if rev_ref: _setcell(root, rev_ref, f'Revisi : {rev}')
    if date_ref: _setcell(root, date_ref, f'Gresik, {gen_date.day} {BULAN_ID[gen_date.month - 1]} {gen_date.year}')
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for item in zin.infolist():
            raw = zin.read(item.filename)
            if item.filename == part:
                raw = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            if item.filename == 'xl/workbook.xml' and b'fullCalcOnLoad' not in raw:
                raw = re.sub(rb'<calcPr([^>]*?)/>', rb'<calcPr\1 fullCalcOnLoad="1"/>', raw)
            z.writestr(item, raw)
    return out.getvalue()
