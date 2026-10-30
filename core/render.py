"""Pintu tunggal pembuat Excel per departemen."""
import os
from .depts import DEPTS
from .inf_writer import render_inf
from .qc_writer import render_qc

ASSETS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'assets')


def render(dept, plan, rev, gen_date, out_path, teams=None, bulk_pct=None):
    cfg = DEPTS[dept]; tpl = os.path.join(ASSETS, cfg['template'])
    if cfg['kind'] == 'qc':
        return render_qc(plan, rev, gen_date, tpl, out_path, cfg, teams, bulk_pct)
    return render_inf(plan, rev, gen_date, tpl, out_path)
