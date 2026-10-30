import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ['JADWAL_DB'] = tempfile.mktemp(suffix='.db')
from streamlit.testing.v1 import AppTest

def test_app_loads_and_pages():
    at = AppTest.from_file(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),'app.py'), default_timeout=60).run()
    assert not at.exception, at.exception
    at.text_input[0].set_value('admin'); at.text_input[1].set_value('salah'); at.button[0].click().run()
    assert not at.session_state['auth'] if 'auth' in at.session_state else True
    at.text_input[1].set_value('123a'); at.button[0].click().run()
    assert at.session_state['auth'] and not at.exception
    for p in ['Riwayat', 'Saldo Hari', 'Pengaturan']:
        at.sidebar.radio[0].set_value(p).run(); assert not at.exception, (p, at.exception)
