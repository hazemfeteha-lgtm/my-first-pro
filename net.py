# -*- coding: utf-8 -*-
"""طبقة الشبكة لبرنامج «خدمات مصر» (تشغيل على أكثر من جهاز في شبكة LAN).

الأوضاع:
  • standalone : الوضع القديم — قاعدة بيانات محلية على نفس الجهاز.
  • server     : الجهاز الرئيسي يملك قاعدة البيانات ويشغّل خادمًا صغيرًا داخليًا.
  • client     : المحطات (الأمن/الاستقبال/الصيانة) تبعث كل عمليات قاعدة البيانات
                 للسيرفر وتستقبل النتيجة، فلا تُحفظ أي بيانات على المحطة.

كل استعلامات البرنامج تمر من دالة واحدة في center_management.py اسمها db()،
فتكفي إعادة توجيهها هنا لتشتغل كل الشاشات على الشبكة بدون لمس أي استعلام.

الأمان: توكن مشترك + فحص صلاحيات المستخدم على السيرفر نفسه (حسب جدول users)
قبل تنفيذ أي استعلام، وسجل عمليات (audit) لكل إضافة/تعديل/حذف.
"""
from __future__ import annotations

import base64
import json
import os
import re
import socket
import socketserver
import sqlite3
import threading
import urllib.parse
import urllib.request
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

DEFAULT_PORT = 8765
CONFIG_NAME = "net_settings.json"
QUEUE_NAME = "net_queue.jsonl"
PROTO = 1

MODE_STANDALONE = "standalone"
MODE_SERVER = "server"
MODE_CLIENT = "client"

# المستخدم الحالي (يُضبط بعد تسجيل الدخول) — يقرأه العميل في كل طلب.
CURRENT = {"username": None, "is_admin": False}


# ============================== الإعدادات ==============================
def default_config():
    return {
        "mode": MODE_STANDALONE,
        "bind_host": "0.0.0.0",
        "port": DEFAULT_PORT,
        "server_host": "",
        "token": "",
        "station": socket.gethostname(),
        "auto_refresh": 20,          # ثواني تحديث الشاشات في وضع المحطة (0 = إيقاف)
    }


def config_path(base):
    return os.path.join(base, CONFIG_NAME)


def load_config(base):
    cfg = default_config()
    try:
        with open(config_path(base), encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            cfg.update({k: v for k, v in data.items() if k in cfg})
    except Exception:
        pass
    try:
        cfg["port"] = int(cfg.get("port") or DEFAULT_PORT)
    except Exception:
        cfg["port"] = DEFAULT_PORT
    try:
        cfg["auto_refresh"] = int(cfg.get("auto_refresh", 20))
    except Exception:
        cfg["auto_refresh"] = 20
    return cfg


def save_config(base, cfg):
    path = config_path(base)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)
    return path


def new_token():
    import secrets
    return secrets.token_hex(12)


# ========================== تصنيف استعلامات SQL ==========================
_WRITE_RE = re.compile(
    r"^\s*(?:INSERT\s+(?:OR\s+[A-Z]+\s+)?INTO|REPLACE\s+INTO|UPDATE(?:\s+OR\s+[A-Z]+)?|DELETE\s+FROM)"
    r"\s+[\"`\[]?([A-Za-z_][A-Za-z0-9_]*)", re.I)
_READ_RE = re.compile(r"^\s*(SELECT|WITH|PRAGMA|EXPLAIN|VALUES)\b", re.I)
_FROM_RE = re.compile(r"\bFROM\s+[\"`\[]?([A-Za-z_][A-Za-z0-9_]*)", re.I | re.S)

_TABLE_PAGE = {
    "visits": "visits",
    "maintenance": "maintenance",
    "employees": "workers",
    "employee_notes": "employees",
    "purchases": "purchases",
    "attendance": "attend",
    "permits": "attend",
    "absences": "attend",
    "gov_attendance": "gov",
    "gov_staff": "govstaff",
    "inventory": "@inventory",
    "stock_in": "@inventory",
    "users": "@admin",
    "audit_log": "@admin",
}
_INV_PAGES = ("cleaning", "paper", "toner")
_LEVELS = ["none", "view", "add", "edit"]
# جداول مسموح بقراءتها لصفحات إضافية (مثال: شاشة الحضور محتاجة تقرأ أسماء العاملين)
_READ_PAGES = {
    "employees": ("workers", "attend", "employees", "maintenance"),
    "gov_staff": ("govstaff", "gov"),
    "inventory": _INV_PAGES + ("purchases",),
    "stock_in": _INV_PAGES + ("purchases",),
}


def write_table(sql):
    m = _WRITE_RE.match(sql or "")
    return m.group(1).lower() if m else None


def is_write(sql):
    return write_table(sql) is not None


def is_read(sql):
    return bool(_READ_RE.match(sql or ""))


def _level_ok(perms, page, need):
    lv = (perms or {}).get(page, "none")
    if lv not in _LEVELS:
        lv = "none"
    return _LEVELS.index(lv) >= _LEVELS.index(need)


def allowed(user, sql, params):
    """يفحص صلاحية المستخدم على الاستعلام. user = {is_admin, perms}.

    يرجع (مسموح؟, سبب الرفض).
    """
    if user.get("is_admin"):
        return True, ""
    tbl = write_table(sql)
    isw = tbl is not None
    if isw:
        need = "edit" if re.match(r"^\s*(UPDATE|DELETE)", sql or "", re.I) else "add"
    elif is_read(sql):
        need = "view"
        m = _FROM_RE.search(sql or "")
        tbl = m.group(1).lower() if m else None
        if tbl is None:
            return True, ""            # مثل SELECT 1 (بدون جدول)
    else:
        return False, "استعلام غير مسموح"
    page = _TABLE_PAGE.get(tbl)
    if page is None:
        return False, "جدول غير معروف: %s" % tbl
    if page == "@admin":
        return False, "هذا الجدول للمدير فقط"
    if not isw:                         # قراءة: اسمح لو المستخدم يشوف أي صفحة مرتبطة بالجدول
        pages = _READ_PAGES.get(tbl)
        if pages and any(_level_ok(user.get("perms"), p, "view") for p in pages):
            return True, ""
    if page == "@inventory":
        if any(_level_ok(user.get("perms"), p, need) for p in _INV_PAGES):
            return True, ""
        return False, "غير مسموح بالجرد"
    if _level_ok(user.get("perms"), page, need):
        return True, ""
    return False, "غير مسموح"


# ============================== عميل الشبكة ==============================
def base_url(cfg):
    host = (cfg.get("server_host") or "").strip()
    port = int(cfg.get("port") or DEFAULT_PORT)
    if not host:
        raise ValueError("لم يتم إدخال عنوان الجهاز الرئيسي")
    if host.startswith("http://") or host.startswith("https://"):
        base = host.rstrip("/")
        if ":" not in base.split("//", 1)[1]:
            base += ":" + str(port)
        return base
    return "http://%s:%d" % (host, port)


def _post(url, payload, timeout=20):
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        raw = r.read()
    return json.loads(raw.decode("utf-8") or "{}")


def _get(url, timeout=10):
    with urllib.request.urlopen(url, timeout=timeout) as r:
        raw = r.read()
    return json.loads(raw.decode("utf-8") or "{}")


class RemoteCursor:
    """يشبه مؤشر sqlite3 لكن نتائجه وصلت جاهزة من السيرفر."""

    def __init__(self, rows=None, lastrowid=0, rowcount=-1):
        self._rows = [tuple(r) for r in (rows or [])]
        self.lastrowid = lastrowid
        self.rowcount = rowcount
        self.description = None

    def __iter__(self):
        return iter(self._rows)

    def fetchall(self):
        return list(self._rows)

    def fetchone(self):
        return self._rows[0] if self._rows else None

    def fetchmany(self, size=1):
        out = self._rows[:size]
        self._rows = self._rows[size:]
        return out


class RemoteConnection:
    """اتصال قاعدة بيانات عبر الشبكة: نفس واجهة sqlite3.Connection."""

    def __init__(self, cfg, queue_path=None, timeout=20):
        self.cfg = cfg
        self.base = base_url(cfg)
        self.timeout = timeout
        self.queue_path = queue_path
        self.token = cfg.get("token", "")
        self.station = cfg.get("station", "")
        self._flushed = False

    # --- واجهة الاتصال ---
    def execute(self, sql, params=()):
        return self._call(sql, list(params or ()), False)

    def executemany(self, sql, seq):
        return self._call(sql, [list(p) for p in seq], True)

    def commit(self):
        pass

    def close(self):
        pass

    # --- داخلي ---
    def _call(self, sql, params, many):
        if self.queue_path and not self._flushed:
            self._flushed = True
            try:
                self._flush()
            except Exception:
                pass
        payload = {"token": self.token, "station": self.station,
                   "user": CURRENT.get("username"), "sql": sql,
                   "params": params, "many": many}
        try:
            data = _post(self.base + "/rpc", payload, self.timeout)
        except Exception as e:
            if is_write(sql) and self.queue_path:
                self._enqueue(sql, params, many)
                return RemoteCursor([])
            raise sqlite3.OperationalError("تعذر الاتصال بالسيرفر: %s" % e)
        if not data.get("ok"):
            raise sqlite3.OperationalError(data.get("error") or "خطأ من السيرفر")
        return RemoteCursor(data.get("rows"), data.get("lastrowid", 0), data.get("rowcount", -1))

    def _enqueue(self, sql, params, many):
        try:
            with open(self.queue_path, "a", encoding="utf-8") as f:
                f.write(json.dumps({"sql": sql, "params": params, "many": many}, ensure_ascii=False) + "\n")
        except Exception:
            pass

    def _flush(self):
        if not self.queue_path or not os.path.isfile(self.queue_path):
            return
        with open(self.queue_path, encoding="utf-8") as f:
            items = []
            for line in f:
                line = line.strip()
                if line:
                    try:
                        items.append(json.loads(line))
                    except ValueError:
                        pass
        if not items:
            os.remove(self.queue_path)
            return
        for it in items:
            data = _post(self.base + "/rpc",
                         {"token": self.token, "station": self.station, "user": CURRENT.get("username"),
                          "sql": it.get("sql"), "params": it.get("params"), "many": it.get("many")},
                         self.timeout)
            if not data.get("ok"):
                raise sqlite3.OperationalError(data.get("error") or "فشل إرسال العمليات المعلّقة")
        os.remove(self.queue_path)


def login(cfg, username, password):
    """تسجيل دخول على السيرفر (التحقق من كلمة المرور عند السيرفر). يرجع (user, reason)."""
    try:
        data = _post(base_url(cfg) + "/login",
                     {"token": cfg.get("token", ""), "station": cfg.get("station", ""),
                      "username": username, "password": password}, timeout=15)
    except Exception as e:
        raise sqlite3.OperationalError("تعذر الاتصال بالسيرفر: %s" % e)
    if not data.get("ok"):
        return None, data.get("reason", "wrong")
    u = data.get("user") or {}
    CURRENT["username"] = u.get("username")
    CURRENT["is_admin"] = bool(u.get("is_admin"))
    return u, ""


def ping(cfg, timeout=4):
    """اختبار اتصال سريع. يرجع dict أو يرفع استثناء."""
    q = urllib.parse.quote(cfg.get("token", ""))
    return _get(base_url(cfg) + "/ping?token=" + q, timeout)


def server_info(cfg):
    q = urllib.parse.quote(cfg.get("token", ""))
    return _get(base_url(cfg) + "/info?token=" + q, timeout=8)


def request_backup(cfg):
    return _post(base_url(cfg) + "/backup",
                 {"token": cfg.get("token", ""), "user": CURRENT.get("username")}, timeout=60)


def fetch_audit(cfg, limit=300):
    """يجلب سجل العمليات من السيرفر (للمدير فقط)."""
    data = _post(base_url(cfg) + "/audit",
                 {"token": cfg.get("token", ""), "user": CURRENT.get("username"), "limit": limit}, timeout=30)
    if not data.get("ok"):
        raise sqlite3.OperationalError(data.get("error") or "تعذر جلب السجل")
    return [tuple(r) for r in data.get("rows", [])]


def upload_photo(cfg, local_path, name):
    with open(local_path, "rb") as f:
        blob = base64.b64encode(f.read()).decode("ascii")
    return _post(base_url(cfg) + "/upload",
                 {"token": cfg.get("token", ""), "user": CURRENT.get("username"),
                  "name": name, "data": blob}, timeout=120)


def fetch_photo(cfg, name, dest_path):
    url = (base_url(cfg) + "/file?name=" + urllib.parse.quote(os.path.basename(name))
           + "&token=" + urllib.parse.quote(cfg.get("token", "")))
    with urllib.request.urlopen(url, timeout=120) as r:
        blob = r.read()
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    with open(dest_path, "wb") as f:
        f.write(blob)
    return dest_path


# ============================== السيرفر ==============================
class _ThreadingServer(ThreadingHTTPServer):
    """ThreadingHTTPServer بدون استعلام DNS عكسي (getfqdn) اللي بيعلّق التشغيل."""

    daemon_threads = True
    allow_reuse_address = True

    def server_bind(self):
        socketserver.TCPServer.server_bind(self)
        host, port = self.server_address[:2]
        self.server_name = host
        self.server_port = port

    def handle_error(self, request, client_address):
        pass            # تجاهل أخطاء انقطاع الاتصال (Broken pipe) بدل طباعة تتبّع مزعج


class Server:
    """خادم LAN صغير يملك قاعدة البيانات وينفّذ الاستعلامات القادمة من المحطات."""

    def __init__(self, db_path, token="", verify=None, photos_dir=None,
                 backup_fn=None, info_fn=None, log=None):
        self.db_path = db_path
        self.token = token or ""
        self.verify = verify            # callable(username, password) -> user dict | None
        self.photos_dir = photos_dir
        self.backup_fn = backup_fn      # callable() -> backup path
        self.info_fn = info_fn          # callable() -> dict (db_mb, photos_mb, total)
        self.log = log or (lambda m: None)
        self.version = 0
        self.lock = threading.Lock()
        self.httpd = None
        self.thread = None
        self.port = None
        self.stations = {}              # station -> آخر وقت ظهور
        self.write_count = 0

    # --- قاعدة البيانات ---
    def _conn(self):
        c = sqlite3.connect(self.db_path, timeout=15)
        c.execute("PRAGMA busy_timeout=10000")
        return c

    def ensure(self):
        c = self._conn()
        try:
            c.execute("""CREATE TABLE IF NOT EXISTS audit_log(
                id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, username TEXT, station TEXT,
                action TEXT, table_name TEXT, detail TEXT)""")
            c.commit()
        finally:
            c.close()

    def _auth(self, token):
        return (not self.token) or (token == self.token)

    def _user(self, username):
        if not username:
            return None
        c = self._conn()
        try:
            r = c.execute("SELECT is_admin, active, perms FROM users WHERE username=?",
                          (username,)).fetchone()
        finally:
            c.close()
        if not r or not r[1]:
            return None
        try:
            perms = json.loads(r[2] or "{}")
        except ValueError:
            perms = {}
        return {"is_admin": bool(r[0]), "perms": perms}

    def _audit(self, username, station, action, table, detail):
        c = self._conn()
        try:
            c.execute("INSERT INTO audit_log(ts,username,station,action,table_name,detail) "
                      "VALUES(?,?,?,?,?,?)",
                      (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), username, station,
                       action, table, detail))
            c.commit()
        finally:
            c.close()

    def _photo_path(self, name):
        name = os.path.basename(name or "")
        if not name or not self.photos_dir:
            return None
        p = os.path.join(self.photos_dir, name)
        return p if os.path.isfile(p) else None

    def _info(self):
        size = 0
        try:
            size = os.path.getsize(self.db_path) / 1048576
        except Exception:
            pass
        photos = 0
        if self.photos_dir and os.path.isdir(self.photos_dir):
            for root, _, files in os.walk(self.photos_dir):
                for fn in files:
                    try:
                        photos += os.path.getsize(os.path.join(root, fn))
                    except Exception:
                        pass
        total = 0
        try:
            c = self._conn()
            try:
                for t in ("visits", "maintenance", "inventory", "employee_notes", "employees"):
                    try:
                        total += c.execute("SELECT COUNT(*) FROM %s" % t).fetchone()[0]
                    except sqlite3.Error:
                        pass
            finally:
                c.close()
        except Exception:
            pass
        if self.info_fn:
            try:
                extra = self.info_fn() or {}
                size = extra.get("db_mb", size)
                photos = extra.get("photos_mb", photos)
                total = extra.get("total", total)
            except Exception:
                pass
        return {"version": self.version, "db_mb": round(size, 2),
                "photos_mb": round(photos, 2), "total": total,
                "writes": self.write_count, "stations": len(self.stations)}

    def _rpc(self, req):
        sql = req.get("sql") or ""
        params = req.get("params") or []
        many = bool(req.get("many"))
        username = req.get("user")
        station = req.get("station") or ""
        user = self._user(username)
        if not user:
            return {"ok": False, "error": "المستخدم غير مصرح"}
        ok, why = allowed(user, sql, params)
        if not ok:
            return {"ok": False, "error": why}
        tbl = write_table(sql)
        with self.lock:
            c = self._conn()
            try:
                if many:
                    cur = c.executemany(sql, [tuple(p) for p in params])
                    rows = []
                else:
                    cur = c.execute(sql, tuple(params))
                    rows = cur.fetchall() if cur.description else []
                lastrowid = cur.lastrowid or 0
                rowcount = cur.rowcount
                c.commit()
                if tbl:
                    self.version += 1
                    self.write_count += 1
                    if station:
                        self.stations[station] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    up = sql.strip().upper()
                    action = ("DELETE" if up.startswith("DELETE") else
                              "UPDATE" if up.startswith("UPDATE") else "INSERT")
                    try:
                        self._audit(username, station, action, tbl, sql.strip()[:300])
                    except Exception:
                        pass
            except sqlite3.Error as e:
                return {"ok": False, "error": str(e)}
            finally:
                c.close()
        return {"ok": True, "rows": rows, "lastrowid": lastrowid,
                "rowcount": rowcount, "version": self.version}

    def _upload(self, req):
        name = os.path.basename(req.get("name") or "")
        if not name or not self.photos_dir:
            return {"ok": False, "error": "اسم ملف غير صالح"}
        try:
            blob = base64.b64decode(req.get("data") or "")
        except Exception:
            return {"ok": False, "error": "بيانات غير صالحة"}
        os.makedirs(self.photos_dir, exist_ok=True)
        with open(os.path.join(self.photos_dir, name), "wb") as f:
            f.write(blob)
        return {"ok": True}

    # --- التشغيل ---
    def start(self, host="0.0.0.0", port=DEFAULT_PORT):
        srv = self

        class Handler(BaseHTTPRequestHandler):
            protocol_version = "HTTP/1.1"

            def log_message(self, *a):
                pass

            def _json(self, obj, code=200):
                data = json.dumps(obj).encode("utf-8")
                self.send_response(code)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def _body(self):
                try:
                    n = int(self.headers.get("Content-Length") or 0)
                except ValueError:
                    n = 0
                raw = self.rfile.read(n) if n else b""
                try:
                    return json.loads(raw.decode("utf-8") or "{}")
                except ValueError:
                    return None

            def do_GET(self):
                u = urllib.parse.urlparse(self.path)
                q = urllib.parse.parse_qs(u.query)
                tok = (q.get("token", [""])[0])
                if u.path == "/ping":
                    self._json({"ok": True, "proto": PROTO, "app": "KhadamatMisr",
                                "version": srv.version, "stations": len(srv.stations)})
                    return
                if not srv._auth(tok):
                    self._json({"ok": False, "error": "توكن غير صحيح"}, 403)
                    return
                if u.path == "/info":
                    self._json(dict({"ok": True}, **srv._info()))
                    return
                if u.path == "/file":
                    path = srv._photo_path(q.get("name", [""])[0])
                    if not path:
                        self._json({"ok": False, "error": "الملف غير موجود"}, 404)
                        return
                    with open(path, "rb") as f:
                        blob = f.read()
                    self.send_response(200)
                    self.send_header("Content-Type", "application/octet-stream")
                    self.send_header("Content-Length", str(len(blob)))
                    self.end_headers()
                    self.wfile.write(blob)
                    return
                self._json({"ok": False, "error": "not found"}, 404)

            def do_POST(self):
                u = urllib.parse.urlparse(self.path)
                req = self._body()
                if req is None:
                    self._json({"ok": False, "error": "طلب غير صالح"}, 400)
                    return
                if not srv._auth(req.get("token", "")):
                    self._json({"ok": False, "error": "توكن غير صحيح"}, 403)
                    return
                if u.path == "/login":
                    user = srv.verify(req.get("username", ""), req.get("password", "")) if srv.verify else None
                    if not user:
                        self._json({"ok": False, "reason": "wrong"})
                        return
                    srv.log("دخول: %s من %s" % (user.get("username"), req.get("station") or "?"))
                    self._json({"ok": True, "user": user})
                    return
                if u.path == "/rpc":
                    self._json(srv._rpc(req))
                    return
                if u.path == "/upload":
                    self._json(srv._upload(req))
                    return
                if u.path == "/backup":
                    user = srv._user(req.get("user"))
                    if not (user and user.get("is_admin")):
                        self._json({"ok": False, "error": "النسخ الاحتياطي للمدير فقط"})
                        return
                    try:
                        path = srv.backup_fn() if srv.backup_fn else None
                        self._json({"ok": True, "path": os.path.basename(path) if path else ""})
                    except Exception as e:
                        self._json({"ok": False, "error": str(e)})
                    return
                if u.path == "/audit":
                    user = srv._user(req.get("user"))
                    if not (user and user.get("is_admin")):
                        self._json({"ok": False, "error": "سجل العمليات للمدير فقط"})
                        return
                    try:
                        limit = int(req.get("limit") or 300)
                    except (TypeError, ValueError):
                        limit = 300
                    try:
                        c = srv._conn()
                        try:
                            c.execute("""CREATE TABLE IF NOT EXISTS audit_log(
                                id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, username TEXT, station TEXT,
                                action TEXT, table_name TEXT, detail TEXT)""")
                            rows = c.execute("SELECT ts,username,station,action,table_name,detail "
                                             "FROM audit_log ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
                        finally:
                            c.close()
                        self._json({"ok": True, "rows": rows})
                    except Exception as e:
                        self._json({"ok": False, "error": str(e)})
                    return
                self._json({"ok": False, "error": "not found"}, 404)

        self.ensure()
        self.httpd = _ThreadingServer((host, port), Handler)
        self.httpd.daemon_threads = True
        self.port = self.httpd.server_address[1]
        self.thread = threading.Thread(target=self.httpd.serve_forever, name="khm-net", daemon=True)
        self.thread.start()
        self.log("السيرفر يعمل على %s:%d" % (host, self.port))
        return self

    def stop(self):
        try:
            if self.httpd:
                self.httpd.shutdown()
                self.httpd.server_close()
        except Exception:
            pass
        self.httpd = None
        self.thread = None


def local_ip():
    """أفضل تخمين لعنوان الجهاز على الشبكة المحلية (للطباعة في نافذة الشبكة)."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"
    finally:
        s.close()


def audit_rows(db_path, limit=300):
    """آخر العمليات المسجّلة (للعرض في نافذة سجل العمليات)."""
    try:
        c = sqlite3.connect(db_path)
        try:
            c.execute("""CREATE TABLE IF NOT EXISTS audit_log(
                id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, username TEXT, station TEXT,
                action TEXT, table_name TEXT, detail TEXT)""")
            return c.execute("SELECT ts,username,station,action,table_name,detail "
                             "FROM audit_log ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
        finally:
            c.close()
    except Exception:
        return []
