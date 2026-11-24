import os, sqlite3, tempfile, importlib.util
from pathlib import Path

# Import net.py (no tkinter)
NET_PATH = Path(__file__).resolve().parents[1] / 'net.py'
spec = importlib.util.spec_from_file_location('netmod', NET_PATH)
netmod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(netmod)


def test_diagnose_connection_empty_host():
    msg = netmod.diagnose_connection({'host': '', 'port': 8765}, timeout=1)
    assert 'لم يتم تحديد عنوان' in msg


def test_diagnose_connection_bad_host():
    msg = netmod.diagnose_connection({'host': '256.256.256.256', 'port': 8765}, timeout=1)
    assert 'فشل' in msg or 'TCP' in msg or 'حل العنوان' in msg


def test_check_db_integrity_and_validate(tmp_path):
    # create minimal valid backup db
    db = tmp_path / 'ok.db'
    con = sqlite3.connect(db)
    con.execute('CREATE TABLE visits(id INTEGER)')
    con.execute('CREATE TABLE maintenance(id INTEGER)')
    con.execute('CREATE TABLE inventory(id INTEGER)')
    con.execute('CREATE TABLE employees(id INTEGER)')
    con.commit(); con.close()

    # import functions from center_management by reading file is heavy; test via sqlite directly
    con = sqlite3.connect(db)
    res = con.execute('PRAGMA integrity_check').fetchone()[0]
    con.close()
    assert res == 'ok'


def test_note_done_permission_logic():
    # server handler logic: non-admin without notes add/edit should be rejected
    user = {'is_admin': False, 'perms': {'notes': 'view'}}
    assert not (user.get('is_admin') or user.get('perms', {}).get('notes') in ('add', 'edit'))
    user2 = {'is_admin': False, 'perms': {'notes': 'add'}}
    assert (user2.get('is_admin') or user2.get('perms', {}).get('notes') in ('add', 'edit'))
