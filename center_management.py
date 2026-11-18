# -*- coding: utf-8 -*-
"""
خدمات مصر - نظام إدارة ومتابعة المركز (الإسكندرية)
تصميم: حازم فتيحه

الصفحات: الزيارات، الصيانة، جرد النظافة، جرد الأدوات الورقية، جرد الأحبار،
         ملاحظات وأخطاء العاملين، استخراج التقارير.
التشغيل: python center_management.py
المتطلبات: pip install openpyxl
"""
import os, sys, csv, json, html, calendar, bisect, shutil, subprocess, webbrowser, sqlite3, hashlib, hmac, secrets, functools, re, traceback, tempfile, zipfile, tkinter as tk
from pathlib import Path
from contextlib import contextmanager
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkfont
from datetime import datetime, date, timedelta
# مكتبة Excel مرفقة داخل مجلد libs فلا يحتاج البرنامج إنترنت للتثبيت
_BASE = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_BASE, "libs"))
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

APP_NAME = "خدمات مصر"
APP_VERSION = "2.11.1"
DESIGNER = "حازم فتيحه"
BRANCH = "الإسكندرية"

# عند تحويله إلى exe تُحفظ قاعدة البيانات بجوار ملف exe وليس في مجلد مؤقت
BASE = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))


def _res_dir():
    """مجلد الموارد المرفقة (img / fonts) — يشتغل مع الـ exe (onefile) والعادي."""
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


RES_DIR = _res_dir()
DB_PATH = os.path.join(BASE, "center_management.db")
CRASH_LOG = os.path.join(BASE, "crash.log")

# ---- الشبكة المحلية (LAN): وضع مستقل / سيرفر (جهاز رئيسي) / محطة ----
try:
    import net as netmod
except Exception:                       # لو الملف مش موجود، البرنامج يشتغل عادي (مستقل)
    netmod = None

NET_CFG = netmod.load_config(BASE) if netmod else {"mode": "standalone"}
NET_MODE = NET_CFG.get("mode", "standalone")
NET_STATION = NET_CFG.get("station") or ""
NET_QUEUE = os.path.join(BASE, netmod.QUEUE_NAME) if netmod else None
NET_SERVER = None                       # كائن الخادم في وضع «سيرفر»
NET_STATUS = {"ok": True, "msg": "", "at": None}


def apply_cli():
    """يتيح تشغيل البرنامج بأوضاع الشبكة من سطر الأوامر (اختياري).

    أمثلة:  KhadamatMisr.exe --server
            KhadamatMisr.exe --client 192.168.1.10 --token xxxx --station الأمن
    """
    global NET_MODE, NET_STATION
    if not netmod:
        return
    args = sys.argv[1:]
    changed = False
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--server":
            NET_CFG["mode"] = "server"; changed = True
        elif a == "--standalone":
            NET_CFG["mode"] = "standalone"; changed = True
        elif a == "--client" and i + 1 < len(args):
            NET_CFG["mode"] = "client"; NET_CFG["server_host"] = args[i + 1]; changed = True; i += 1
        elif a == "--port" and i + 1 < len(args):
            try:
                NET_CFG["port"] = int(args[i + 1]); changed = True
            except ValueError:
                pass
            i += 1
        elif a == "--token" and i + 1 < len(args):
            NET_CFG["token"] = args[i + 1]; changed = True; i += 1
        elif a == "--station" and i + 1 < len(args):
            NET_CFG["station"] = args[i + 1]; changed = True; i += 1
        i += 1
    if changed:
        NET_MODE = NET_CFG.get("mode", "standalone")
        NET_STATION = NET_CFG.get("station") or NET_STATION
        try:
            netmod.save_config(BASE, NET_CFG)
        except Exception:
            pass

# ---- الألوان والخطوط (وضع نهاري / ليلي) ----
MONTHS_AR = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
             "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
FONT = "Tahoma"
EMOJI_FONT = "Segoe UI Emoji"
SETTINGS_PATH = os.path.join(BASE, "settings.json")
# ===== الثيمات: ٩ أنماط (٨ من صور المستخدم + نمط «ذهبي فاخر» من مواصفات التصميم) =====
# كل ثيم يحدّد: تدرّج الشريط (H1→H2)، القائمة (NAVY_2)، الخلفية (BG)، الكروت (WHITE)،
# اللون الأساسي (BLUE/PRIMARY)، الذهبي (GOLD)، ولون نص الشريط (HDR_TEXT/HDR_SUB) للثيمات الفاتحة.
THEMES = {
    # ١) أزرق بحري — شريط متدرّج (كحلي ← تركوازي) + محتوى فاتح
    "ocean": dict(NAVY="#0b1d3a", NAVY_2="#13315c", BLUE="#2a7fff", GOLD="#f5b82e", BG="#eef3fb", WHITE="#ffffff",
                  TEXT="#14233b", MUTED="#6a7d99", LINE="#dbe5f2", RED="#e5484d", GREEN="#12b886", ORANGE="#ff8a3d",
                  PURPLE="#7c5cff", TEAL="#14b8a6", CYAN="#06b6d4", AMBER="#f59e0b", PINK="#ec4899",
                  ODD="#f5f8fe", SOFT="#eaf0fa", SEL="#cfe0ff", INK="#0b1d3a", PRIMARY="#13315c",
                  H1="#0b1d3a", H2="#0f6a86", ENTRY="#ffffff", HDR_TEXT="#ffffff", HDR_SUB="#ffd978"),
    # ٢) كحلي كلاسيكي — شريط كحلي ثابت + محتوى فاتح
    "navy": dict(NAVY="#0a1c3c", NAVY_2="#0c2444", BLUE="#1667c8", GOLD="#e8b64c", BG="#eef3f9", WHITE="#ffffff",
                 TEXT="#14233b", MUTED="#6a7d99", LINE="#d8e2ee", RED="#e5484d", GREEN="#12b886", ORANGE="#ff8a3d",
                 PURPLE="#7c5cff", TEAL="#14b8a6", CYAN="#06b6d4", AMBER="#f59e0b", PINK="#ec4899",
                 ODD="#f5f8fc", SOFT="#e6eef8", SEL="#cfe0ff", INK="#0a1c3c", PRIMARY="#123a6b",
                 H1="#0a1c3c", H2="#0d2a52", ENTRY="#ffffff", HDR_TEXT="#ffffff", HDR_SUB="#ffd978"),
    # ٣) بيج رملي — شريط بيج + قائمة رمادية مزرقّة
    "sand": dict(NAVY="#e8d8b8", NAVY_2="#7d8f92", BLUE="#8a6d3b", GOLD="#c9a227", BG="#f5f2ea", WHITE="#ffffff",
                 TEXT="#3a352c", MUTED="#7d7364", LINE="#e3dccb", RED="#b23a3a", GREEN="#3f7d4e", ORANGE="#c47a2c",
                 PURPLE="#6b5b95", TEAL="#4f8a8b", CYAN="#2f7f8f", AMBER="#c9971f", PINK="#b05a7a",
                 ODD="#faf7f0", SOFT="#efe8d8", SEL="#e6d9b8", INK="#2f2a22", PRIMARY="#8a7f6a",
                 H1="#f8f0d8", H2="#c8b8a0", ENTRY="#ffffff", HDR_TEXT="#3a352c", HDR_SUB="#8a6d3b"),
    # ٤) فاتح بقائمة كحلية — شريط فاتح + قائمة كحلية غامقة
    "pearl": dict(NAVY="#eef0f4", NAVY_2="#232c3c", BLUE="#3b5bdb", GOLD="#c9a227", BG="#f4f5f7", WHITE="#ffffff",
                  TEXT="#22293a", MUTED="#6b7280", LINE="#dde1e7", RED="#dc2626", GREEN="#16a34a", ORANGE="#ea580c",
                  PURPLE="#7c3aed", TEAL="#0d9488", CYAN="#0891b2", AMBER="#d97706", PINK="#db2777",
                  ODD="#f8f9fb", SOFT="#e9ecf1", SEL="#dbe4ff", INK="#22293a", PRIMARY="#2b3648",
                  H1="#f0f0f8", H2="#e2e6ee", ENTRY="#ffffff", HDR_TEXT="#22293a", HDR_SUB="#3b5bdb"),
    # ٥) ليلي فحمي — فحمي + زيتوني
    "charcoal": dict(NAVY="#1c1c1c", NAVY_2="#101010", BLUE="#7f8c5a", GOLD="#c8b06a", BG="#0b0b0b", WHITE="#171717",
                     TEXT="#e2e2d6", MUTED="#8f9080", LINE="#2a2a26", RED="#e0736f", GREEN="#8fae6a", ORANGE="#d9a05a",
                     PURPLE="#a99bd6", TEAL="#6fae9a", CYAN="#6fc0c8", AMBER="#c8b06a", PINK="#d98fb0",
                     ODD="#131311", SOFT="#1a1a17", SEL="#3a3d2a", INK="#f0f0e4", PRIMARY="#3a3a30",
                     H1="#1c1c1c", H2="#303030", ENTRY="#121210", HDR_TEXT="#ffffff", HDR_SUB="#d8cfa8"),
    # ٦) ليلي كحلي — كحلي رمادي داكن
    "midnight": dict(NAVY="#182028", NAVY_2="#131b28", BLUE="#3b82f6", GOLD="#e0b64a", BG="#0b1320", WHITE="#182534",
                     TEXT="#dfe8f5", MUTED="#8a9bb0", LINE="#22303f", RED="#f87171", GREEN="#34d399", ORANGE="#fb923c",
                     PURPLE="#a78bfa", TEAL="#2dd4bf", CYAN="#22d3ee", AMBER="#fbbf24", PINK="#f472b6",
                     ODD="#101a28", SOFT="#16202e", SEL="#1e3a5f", INK="#eef4fb", PRIMARY="#2b4a7a",
                     H1="#182028", H2="#223040", ENTRY="#101828", HDR_TEXT="#ffffff", HDR_SUB="#e0b64a"),
    # ٧) ليلي أسود — أسود مزرق عميق
    "onyx": dict(NAVY="#000810", NAVY_2="#050d18", BLUE="#3b82f6", GOLD="#e8b64c", BG="#0a1420", WHITE="#16202e",
                 TEXT="#e6eef8", MUTED="#8aa0c0", LINE="#1e2c3d", RED="#f87171", GREEN="#34d399", ORANGE="#fb923c",
                 PURPLE="#a78bfa", TEAL="#2dd4bf", CYAN="#22d3ee", AMBER="#fbbf24", PINK="#f472b6",
                 ODD="#0d1826", SOFT="#131e2c", SEL="#1b3a63", INK="#f2f7fd", PRIMARY="#2a4a78",
                 H1="#000810", H2="#0a1a2c", ENTRY="#0c1622", HDR_TEXT="#ffffff", HDR_SUB="#e8b64c"),
    # ٨) ليلي أخضر — أخضر/تركوازي داكن
    "forest": dict(NAVY="#1b2b2f", NAVY_2="#16272c", BLUE="#2f8f74", GOLD="#d9b44a", BG="#121a20", WHITE="#1f2a30",
                   TEXT="#dfeae8", MUTED="#8fa5a2", LINE="#26363a", RED="#e57373", GREEN="#5fbf8f", ORANGE="#d9a05a",
                   PURPLE="#a99bd6", TEAL="#3fbfa5", CYAN="#4dd0e1", AMBER="#d9b44a", PINK="#d98fb0",
                   ODD="#16211f", SOFT="#1a2724", SEL="#22463c", INK="#eef6f4", PRIMARY="#2f6b57",
                   H1="#283038", H2="#184840", ENTRY="#141e1c", HDR_TEXT="#ffffff", HDR_SUB="#d9b44a"),
    # ٩) ذهبي فاخر — ليلي فاخر بلمسات ذهبية (من مواصفات تصميم المستخدم)
    "gold": dict(NAVY="#14181F", NAVY_2="#14181F", BLUE="#C89D46", GOLD="#E0B963", BG="#1B2028", WHITE="#323844",
                 TEXT="#FFFFFF", MUTED="#A0A8B5", LINE="#3A4150", RED="#E5484D", GREEN="#1D7A5F", ORANGE="#C89D46",
                 PURPLE="#8C7A3E", TEAL="#2F8F74", CYAN="#3AA88A", AMBER="#E0B963", PINK="#C89D46",
                 ODD="#2B313C", SOFT="#252A34", SEL="#3A4150", INK="#FFFFFF", PRIMARY="#C89D46",
                 H1="#0E3B33", H2="#151B2E", ENTRY="#14181F", HDR_TEXT="#FFFFFF", HDR_SUB="#E0B963",
                 BTN="gold", FONT_KEY="cairo",
                 LOGIN_G1="#0E3B33", LOGIN_G2="#151B2E", LOGIN_BOX="#16232A",
                 LOGIN_BTN=["#1D7A5F", "#8C7A3E", "#1B2B65"]),
}

THEME_LABELS = {"ocean": "أزرق بحري", "navy": "كحلي كلاسيكي", "sand": "بيج رملي",
                "pearl": "فاتح بقائمة كحلية", "charcoal": "ليلي فحمي", "midnight": "ليلي كحلي",
                "onyx": "ليلي أسود", "forest": "ليلي أخضر", "gold": "ذهبي فاخر"}

# ترتيب عرض الثيمات في شاشة المظهر
THEME_ORDER = ["ocean", "navy", "sand", "pearl", "charcoal", "midnight", "onyx", "forest", "gold"]
DARK_THEMES = ("charcoal", "midnight", "onyx", "forest", "gold")

# أشكال الواجهة: كثافة العرض وحجم الأزرار (تُختار من نافذة الألوان والمظهر)
UI_STYLES = {
    "comfort": {"label": "مريح", "ROW_H": 32, "BTN_PADX": 18, "BTN_PADY": 8, "NAV_PADY": 1},
    "compact": {"label": "مضغوط", "ROW_H": 26, "BTN_PADX": 12, "BTN_PADY": 5, "NAV_PADY": 0},
    "classic": {"label": "كلاسيكي", "ROW_H": 36, "BTN_PADX": 22, "BTN_PADY": 10, "NAV_PADY": 2},
}

# ---- محرك الخطوط: نوع الخط وحجمه ----
UI_FONTS = {"tahoma": "Tahoma", "segoe": "Segoe UI", "arial": "Arial",
            "calibri": "Calibri", "verdana": "Verdana", "majalla": "Sakkal Majalla",
            "cairo": "Cairo", "tajawal": "Tajawal"}
FONT_LABELS = {"tahoma": "Tahoma (افتراضي)", "segoe": "Segoe UI", "arial": "Arial",
               "calibri": "Calibri", "verdana": "Verdana", "majalla": "Sakkal Majalla",
               "cairo": "Cairo (عربي عصري)", "tajawal": "Tajawal (عربي عصري)"}
FONT_SIZES = {"small": 0.88, "normal": 1.0, "large": 1.16}
FONT_SIZE_LABELS = {"small": "صغير", "normal": "عادي", "large": "كبير"}

# ---- أشكال الأزرار: مملوء / محدّد / ناعم / ذهبي فاخر ----
BTN_STYLES = {"solid": "مملوء", "outline": "محدّد", "soft": "ناعم", "gold": "ذهبي فاخر"}


def load_settings():
    try:
        with open(SETTINGS_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(d):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
    except Exception:
        pass


def register_bundled_fonts():
    """يسجّل خطوط Cairo/Tajawal المرفقة في مجلد fonts عشان تبان في البرنامج.

    ويندوز: AddFontResourceExW بصلاحية خاصة للعملية (مش بتتثبت على الجهاز).
    لينكس: نسخ لمجلد خطوط المستخدم + fc-cache.
    لو فشل أي حاجة، البرنامج بيكمّل عادي بخط النظام.
    """
    d = os.path.join(RES_DIR, "fonts")
    if not os.path.isdir(d):
        return []
    files = [os.path.join(d, f) for f in sorted(os.listdir(d)) if f.lower().endswith((".ttf", ".otf"))]
    if not files:
        return []
    if sys.platform.startswith("win"):
        try:
            import ctypes
            gdi = ctypes.windll.gdi32
            for f in files:
                gdi.AddFontResourceExW(ctypes.c_wchar_p(f), 0x10, 0)   # FR_PRIVATE
        except Exception:
            pass
        return files
    try:
        import shutil as _sh
        tgt = os.path.join(os.path.expanduser("~"), ".local", "share", "fonts")
        os.makedirs(tgt, exist_ok=True)
        changed = False
        for f in files:
            t = os.path.join(tgt, os.path.basename(f))
            try:
                if not os.path.exists(t) or os.path.getsize(t) != os.path.getsize(f):
                    _sh.copy2(f, t); changed = True
            except Exception:
                pass
        if changed:
            subprocess.run(["fc-cache", "-f", tgt], capture_output=True, timeout=40)
    except Exception:
        pass
    return files


def apply_theme(name):
    """يغيّر ألوان البرنامج (المتغيرات العامة NAVY, BG, ...)."""
    name = name if name in THEMES else THEME_ORDER[0]
    globals().update({"HDR_TEXT": "#ffffff", "HDR_SUB": "#ffd978"})
    globals().update(THEMES[name])
    nav2 = THEMES[name].get("NAVY_2", "#13315c")
    r, g, b = (int(nav2[i:i + 2], 16) for i in (1, 3, 5))
    light_side = (0.299 * r + 0.587 * g + 0.114 * b) >= 140
    globals()["SIDE_MUTED"] = "#33403f" if light_side else "#9fc0e6"
    globals()["SIDE_TEXT"] = "#1f2a2c" if light_side else "#ffffff"
    globals()["SIDE_ACCENT"] = "#6b5a2a" if light_side else THEMES[name].get("GOLD", "#f5b82e")
    globals()["SIDE_ACCENT2"] = "#2f6b57" if light_side else "#9fe3c0"
    globals()["THEME"] = name


def apply_style(name):
    """يغيّر شكل الواجهة: ارتفاع صفوف الجداول وحشو الأزرار وتباعد القائمة."""
    name = name if name in UI_STYLES else "comfort"
    st = UI_STYLES[name]
    globals().update({"STYLE": name, "ROW_H": st["ROW_H"], "BTN_PADX": st["BTN_PADX"],
                      "BTN_PADY": st["BTN_PADY"], "NAV_PADY": st["NAV_PADY"]})


def apply_font(name):
    """يغيّر نوع الخط المستخدم في كل الواجهة."""
    name = name if name in UI_FONTS else "tahoma"
    globals()["FONT"] = UI_FONTS[name]
    globals()["FONT_KEY"] = name


def apply_font_size(name):
    """يغيّر حجم الخط العام (يُطبَّق عبر tk scaling)."""
    name = name if name in FONT_SIZES else "normal"
    globals()["FONT_SCALE"] = FONT_SIZES[name]
    globals()["FONT_SIZE_KEY"] = name


def apply_btn_style(name):
    """يغيّر شكل الأزرار: مملوء / محدّد / ناعم."""
    globals()["BTN_STYLE"] = name if name in BTN_STYLES else "solid"


register_bundled_fonts()          # خطوط Cairo/Tajawal المرفقة (قبل اختيار الخط)
apply_theme(load_settings().get("theme", "ocean"))
apply_style(load_settings().get("style", "comfort"))
apply_font(load_settings().get("font", "tahoma"))
apply_font_size(load_settings().get("font_size", "normal"))
apply_btn_style(load_settings().get("btn_style", "solid"))

# ---- اللغة (عربي / إنجليزي): طبقة الترجمة تُثبَّت قبل إنشاء أي واجهة ----
import i18n
i18n.set_lang(load_settings().get("lang", "ar"))
i18n.install()
T = i18n.T


def _rgb(c):
    return tuple(int(c[i:i + 2], 16) for i in (1, 3, 5))


def lerp_color(c1, c2, t):
    a, b = _rgb(c1), _rgb(c2)
    return "#%02x%02x%02x" % tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def paint_gradient(cv, w, h, c1, c2, horizontal=True, step=4):
    """يرسم تدرجًا لونيًا على Canvas."""
    cv.delete("grad")
    n = w if horizontal else h
    for i in range(0, max(n, 1), step):
        col = lerp_color(c1, c2, i / max(1, n - 1))
        if horizontal:
            cv.create_rectangle(i, 0, i + step, h, outline="", fill=col, tags="grad")
        else:
            cv.create_rectangle(0, i, w, i + step, outline="", fill=col, tags="grad")
    cv.tag_lower("grad")


def gradient_button(parent, text, cmd, stops, width=300, height=48):
    """زرار بتدرّج لوني حقيقي (Canvas) بإطار ذهبي ونص ذهبي — للثيمات الفاخرة (زي شاشة الدخول)."""
    try:
        base_bg = parent.cget("bg")
    except Exception:
        base_bg = "#14181F"
    cv = tk.Canvas(parent, width=width, height=height, highlightthickness=0, bd=0,
                   bg=base_bg, cursor="hand2", takefocus=True)
    st = {"hover": False}

    def _paint(_=None):
        try:
            cv.delete("all")
            w = cv.winfo_width()
            w = w if w and w > 1 else width          # قبل ما النافذة تظهر نستخدم المقاس المطلوب
            h = cv.winfo_height()
            h = h if h and h > 1 else height
            cols = [lerp_color(c, "#ffffff", 0.20) for c in stops] if st["hover"] else list(stops)
            n = max(1, len(cols) - 1)
            seg = max(2, w // 90)
            for i in range(0, w, seg):
                t = i / max(1, w - 1)
                idx = min(n - 1, int(t * n))
                cv.create_rectangle(i, 0, i + seg + 1, h, outline="",
                                    fill=lerp_color(cols[idx], cols[idx + 1], t * n - idx))
            cv.create_rectangle(1, 1, w - 1, h - 1, outline=GOLD, width=1)
            cv.create_text(w // 2, h // 2, text=text, fill=GOLD, font=(FONT, 13, "bold"))
        except Exception:
            pass

    def _go(_=None):
        try:
            cmd()
        except Exception:
            pass
    cv.bind("<Configure>", _paint)
    cv.bind("<Enter>", lambda e: (st.update(hover=True), _paint()))
    cv.bind("<Leave>", lambda e: (st.update(hover=False), _paint()))
    cv.bind("<Button-1>", _go)
    cv.bind("<Return>", _go)
    cv.bind("<space>", _go)
    cv.after(40, _paint)
    return cv

CLEANING_ITEMS = [
    "باركيه بليدج", "سيراميك بليدج", "جلانس", "ديتول",
    "صابون فيبا 2 لتر", "أرضية معطر", "جو معطر", "كلور",
    "ويفز", "شاور", "زعافة كورة", "عصا + مقشة",
    "عصا + مساحة", "عصا + جاروف", "إسفنجة", "فرش قواعد حمام",
    "فوطة مكتب + فوط أرضية", "مناديل سحب حمامات", "علب مناديل مكتب",
    "أكياس قمامة", "بيرسول", "هاربيك", "ليم",
    "جردل شرشوبة", "جردل عادي"
]
PAPER_ITEMS = [
    "ورق تصوير A4", "رول حراري POS", "رول Q", "دبابيس",
    "دباسة", "ملفات", "أقلام", "أظرف", "مناديل ورقية"
]
TONER_ITEMS = [
    "تونر Canon LBP226dw", "حبارة تصوير 3010",
    "تونر طابعة 2", "تونر طابعة 3", "درام / وحدة تصوير"
]
DEPARTMENTS = ["أمن", "خدمات", "صيانة"]
STATUS_ON, STATUS_OFF = "على رأس العمل", "ترك العمل"
PERMIT_LIMIT = 2                                      # أقصى عدد إذونات خروج للعامل في الشهر
DEPT_START = {"خدمات": "07:00", "أمن": "07:00", "صيانة": "07:30"}   # بداية الدوام (الفنيين = قسم الصيانة)
# مواعيد الدوام (بداية، نهاية): الأمن 7ص-7م ، الفنيين (الصيانة) 7:30ص-4م ، الخدمات 7ص-4م
DEPT_SHIFT = {"أمن": ("07:00", "19:00"), "صيانة": ("07:30", "16:00"), "خدمات": ("07:00", "16:00")}
DEPT_LABEL = {"أمن": "الأمن", "صيانة": "الفنيين", "خدمات": "الخدمات"}
ABSENCE_KINDS = ["غياب بدون إذن", "غياب بعذر", "إجازة مرضية", "إجازة اعتيادية", "تأخير"]
SUMMARY_ONLY = ("ملخص العاملين", "ملخص الحضور والغياب", "ملخص حضور الموظفين", "ملخص الجهات")
NOTE_KINDS = ["ملاحظة", "خطأ", "تنبيه شفهي", "إشادة"]
SEVERITIES = ["بسيطة", "متوسطة", "جسيمة"]
INV_NAMES = {"cleaning": "جرد النظافة", "paper": "الجرد الورقي", "toner": "جرد الأحبار"}
INV_ITEMS = {"cleaning": CLEANING_ITEMS, "paper": PAPER_ITEMS, "toner": TONER_ITEMS}
# المشتريات: ورقيات / أحبار / أدوات (القوائم اقتراحات فقط ويمكن كتابة أي صنف يدويًا)
PURCH_CATS = {"paper": "ورقيات", "toner": "أحبار", "tools": "أدوات"}
TOOLS_ITEMS = ["عدة يد", "مفك", "مفتاح ربط", "شاكوش", "بنسة", "شريط لاصق", "لمبة", "كابل كهرباء", "بطارية", "سلم"]
PURCH_ITEMS = {"paper": PAPER_ITEMS, "toner": TONER_ITEMS, "tools": TOOLS_ITEMS}

# لوجو خدمات مصر (مدمج داخل الملف). لتغييره ضع ملف logo.png بجوار البرنامج.
LOGO_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAD4AAAA4CAIAAAAq+twOAAANTUlEQVR42uWae7BX1XXH11p77/P4ve6DC4gIWIkQhREJdMZCU2dC"
    "kibDpDExhoSZhGRiI2LqH/2jzsQktZ1OpvE1dTomAaNTi4mY1EcnMRZESqNYwQS0PiJXeVy4MFzu+/c65+yz91r943fvjdgauXAh"
    "6XTPb34zZ84+53x+a9b6rsf5oUgO/zeXPgf3RAABgPFvBMDfnBT+H9vx9wed30YPb/sNE0dy9tznBp0FQAAnsE7lfufh75XDIICA"
    "CADzBCyO2xin7jk05eSCBERIiEQsDAikCAgZhNmLTATA78bqb38wigiAMHulDHtPqEW4mdTrWSMIjOTO5T7UYbHcrrRGEPAO2AsC"
    "GiU4FhUok/Z+PCNxpFOjDhFRRJJmPYoiYOp+9fVZfzBHDBTCgrDYpAZMAsYxVkKjyYOwaEKtRJGAoADyuGMpOk8OIyLeWecyl1ul"
    "Te74u/f8Yy2pt7W3txfbUXxgqNzWVe7oLJRi0NSwSe6taEKtGUCcJx43OU7Oj84WHZGIFAg5Z8Mw+uFD/7xt5/Y/XHGVy9PqSH+t"
    "Ojg83FevDeQu00a3txW8xirnFtijiDCIgIzbG38HYSqIEselwf6TGzdtvGThApvZ6shI1kiyRpY1bf/Q0In+41meEoAJlUVp+txm"
    "GbDQ23gFgem8oguScs4B4r9tfepkX9/cSy627Jq1BrIowczaIAy8y/O0xt4a0kor5/LMWu+9SEtH5QwkZ2qsrpQCgDfe6C6VysYo"
    "53KFSgkhUVwoUqA7O9ulWiXxkYlCJskcMgMLILJ3IAKIiEiC5xV9Is9orfI8F/aIogQISVCJ0bte+M/HtmyhJH3l4X9JjhyNRYVO"
    "NKB1uRB5BG4pJE8u1U5FNpVWrIH3npkBRIkgogDqMD4yeOyvbrlFDQx+4vFl3c8827e/50Pr/7xAul5PqBQ1bGICgwg0+RJhCtC9"
    "9+w9mNazmQRQBEEUGSvcNevCr1z/lbB/qLO9c0F713O79ixY8SdzFl1qAFQYZASogCYKTQFQ5xGduWXslroRCKAgALASQNRCN9/4"
    "F2UK4NDRQp6XE+sHB1HmqzhIkwzjQAAYBRhA8NzrOvPER5iV0i1qZkFERuO9KKMs2UCrqM4mVaBDqEScNmdmzTbfqAfZqOIQwyIb"
    "8gzAAAxIQPr05f1clF8ogIJjemcYNCoRAQICCACIcy/sARCQWttBROR8Z9P3jgQBBEQk8C2nQhZoJVGWs6rdzzE6grAf0w72CIKI"
    "MJaHBETkLOg1nI+F4+0RogAL45iLnFXhrqdA1IVlvFluVb+t/pNFvPfIPAbpnHc5EhKi8Bj22XRNBOdxicgU3u28ok/t+v+Gju84"
    "wPGREZMI/SYk8X+/duyEvO16PIN4nTS6gDh0AB48glci2gKCYkhrl+b5otHazPpozFnGdkh80wRWSCEJiAA3Eaugkwwi6w1JVaUN"
    "w9pRkBsUAzCpmvdMrC4MLK0hFyMAeEBFGmyydM5Fy0qluDHMZFPjC8VisW6ngUEiZm/rzTaMo6BoooKxrIDzQCwKCqJXwCQIfB58"
    "HcenDoigCAkR4kJcqsycd3GRAdNGCDx9pNnZANvTKwZEErFJpZYESQZxBEo55wLShOMV4/npTfGUQZeQiHOem8nsK64olNuyXa/2"
    "PfSzC44OxX3VdO9LLz/7lFBuXb+21czVm4ZzhRhE4gWcH+tNUcbC4xynJKRWLiFppXQUUUrltTS6eN6KP/vk8Xs3v/mjn57Y+7pj"
    "6jnYvfQvrycN8cDI8C9f6RsZiObNqsya2bSZKYSegXgiVFt+jucUHUDQO08iqEJCJGFEpLicj1bnfHhVpeF69r7wRv8xqDaXrF51"
    "6cc+Ar3dydO733p0R56kc5YujmZ21Um0CkITcp5jFAGzgCAaBP/OKfa759szmH4hsxNGIgWANrOkQGvFjSSvjioAFYVp74GR0f6y"
    "hOUm2GPHD+/Z09d9yFaT+SuWX/z564amdUgUqFKUiy+WKgrR6FCEBQBJK/Sn2J5oKtG994iIqGya6CBk59k5hSB5bkKTDPVXVR6G"
    "RPv277zljs6hJraVjxRp3qqVi//0Q81KxZs4LhaU0ToMi5UO7zObZWEYoYK0mcZRdGowviu6uu22b00qrJ213rEJIkTUxhAppcnl"
    "jtlniL98+b/QlJWOdS4mF2zmplwpL16wcN210ao/Oh4Hqq3CACLc1tZ+5HDPl7+09gPLl82aNTvNGoRKa/NOB3l3h5msr4s2Ggld"
    "bu/+hzujKOTMhcZcv/4mpOD17v2rP7vmtr+/86bPrkl7jpm5cxd/60YZGkErQwprpGeEZUJJwywOYxSe3t62+LJFhUIpc5ZMIOKF"
    "HZkQTq9Ko0knU0ARQVKDA4PM6JiTzOowSLM0DIKuGTNK0yqZc+Vp0/a+tf9723627fiBg2gtBapq9z31730HjrQVS8LYzGz3Wweu"
    "+dSnC4U4NAV2bHOrg1CYp05hxmZr4+iktdEA8J3v3DWx5dvf/rssS9es+bxOm+HgQKlgfvLTR2+/8/ZiqZil2YYb13/0Ix9tuuH7"
    "n9jcfUf3V2/46vr164dHqvc/tPmVV19l9nfccdeKlSsz62xmSSGBIkRA/O3mp0kZHJBsmtx//33bt2/N82x0dLher3rv9u/fv3fv"
    "vjy3hjBQKmnWN963ccGlC7Zt3f6+S97340d+QkpPnzF9848e+sY3v3nvvd/91a/2zpzeddfddz/99NNLrlj6jVu/Xq1WldZKaRBo"
    "qc17us1kHEYEUDcaze9//3tPPvlzIkNEzWYzy2xnZ0exWPTeCYAXzrKs70TfZZddppVauHDBaLVaq9UAsNlsfvzjHyOi5559LkmT"
    "JEk885e+vK7n6JHt27bGcSFJEu+ZWU7H3c8kTNvb22fPnq0UlcvlcrnSSobW2jzPATCO4vaOjiAwIkKKgtCIMBF2dnaISBQXLrpo"
    "dr1RK5balNJRHF911VWLLr+8+803XW6zzGqtWj6K79X/TQIdEdm7OC7NmTP34YcfTtPMe29tphTt3LlzxYoVF1wwq6OjY/PmzS++"
    "+OLIyGhPT8+Rnp6+vn4iSpJ006b7jh3rNSY4ceLkvn0v/e3f/LUwJ2kSBuHgwODu3bsbzaRYLHjviBAEgN6jLDjNlCTAwiLeO2PC"
    "I0d7H3zwwUOHDsZx7D3neT5v3ry1a9d2dU177bXXfvzII9bmxphjx44NDQ3WarXPfOa6devWPfDAAz09Pd77UqkUBMHJkyejMBSR"
    "uFhMmo0rlyz5whe+KCCIYowxxihSgHj2hcAYugg77xF1EEQTWgkgeW6TJBkZGYmiKDQGCY0J+vpOHDhwMIqi+fPna62DIAiDEAkR"
    "0WY2CAMR0Vqnado6OzQ0GAQhImultVGK9JQUAmPozF4EbO6zLGvNLay1WussS8vlcpbZLMuiMEySpjEmDKNiodBoNLz3Q8NDYRgW"
    "CsVKpdxoNJ1zjXq9UCjYPG/ZOMuyQiEGwDDUSimltGpBTxW6iBdBFvI+T9PUex8EYZ7nStGJE31xHLe1tb/y8ktxIe7q6tq2ddvh"
    "w4c/sGzZ0iuvnDFzRpqmv/71Gz9/8slpXdM+t+ZzURxlmdVGG63ZcxAGSmkAMVopTUSKkKbe6ixjr3kBwHt2LkdU11137ZIlS26/"
    "/a5rP3VNmibtHR3V6ujChQt/8Yv/eP/7L7/nnnt27Xr+1lu/vnz58t7e3vb29o0bN7buQKS8d0EQAKBWhAREpJQam1i/O/oksikh"
    "IilAFEHnckBSRCA5mSDLsjRJtVIiUiwVDx06uG7duquvvrqjo+OF3bvXrl27Z8/uHTueCYLg/gf+6V+feHTDhpve7O5etHiRtRYA"
    "oihWioiIWZQSRCWn0XRMroZBxNa4UGujSIkIkYriGBHLpRKO75k7b96qD6/y3tcbjRkzprdVKsePHVeKCoW4Vh1GwFK5PDwyopQS"
    "EaVofPKBSilAgtPrmc6kS1JELIKIAITIzAIAo9VqFBcQsV6vz77wQq20D7glIKSoWCqKCAKWKx2VckUrMkbnuWMWRCJCRKIx30Bq"
    "CSLh1PemrXdzACIMqJT3XgBmzpy5Y8czhw4ePHDgQG9v701f+xoheu+rtZoIbNmyZWBgoFar/+C+jf39J6vVWhCEeW7jOEYAQmrh"
    "j5l67L80U48OANR6rY8KEZCUCl1Qq9Xmz5+/cuXKP/7gB5VSmc3yzLZUrlgsjY4OF4vl3t6jjz32aJpm62+4YcGChQDAzEbr1vgX"
    "aLxUxNOabJyewpxSDJHzrJRi70TEeU8Ijz/+xM0337xp06ZPXvPpgf4+rTUiMfsoitI0cc5FUZzneasoHBke7pzWySwiUohjbYzW"
    "SpEaQz/tmcDke1MREfLMIoxAGzasf/75XZ2d01avXr1u3Rfz3LN32hgAQSQRJlII4JlFxLmcSIVBwMzGGG2MMdpoIyJIqLUhIkD+"
    "LQPOs2yrwXsmIAFwznV37x+tjs6be/GsWRcMDQ4677TWIC0xwomhKREBCLN475m9UlprbYwmUq1FisZGYApOE/2/AZq8agy3nOYr"
    "AAAAAElFTkSuQmCC"
)


# ============================ المستخدمون والصلاحيات ============================
DEFAULT_ADMIN = "admin"
# لا توجد كلمة مرور افتراضية داخل البرنامج: أول تشغيل يطلب من المدير اختيارها.
# القيمة التالية للتعرف على النسخ القديمة فقط (التي كانت تأتي بكلمة مرور admin/6945)
# وإجبارها على التغيير عند أول تشغيل بعد التحديث.
LEGACY_DEFAULT_PW = "6945"
PW_MIN_LEN = 6
WEAK_PASSWORDS = {"admin", "123456", "12345678", "password", "000000", "111111", "123456789", LEGACY_DEFAULT_PW}
_IDENT_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,63}$")
# مستويات الصلاحية لكل قسم: ممنوع < عرض فقط < عرض وإضافة < تعديل وحذف كامل
LEVELS = ["none", "view", "add", "edit"]
LEVEL_LABEL = {"none": "ممنوع", "view": "عرض فقط", "add": "عرض + إضافة", "edit": "تعديل وحذف كامل"}
PERM_PAGES = [("workers", "قاعدة بيانات العاملين"), ("visits", "زيارات المركز"), ("maintenance", "الصيانة"),
              ("cleaning", "جرد خامات النظافة"), ("paper", "جرد الأدوات الورقية"), ("toner", "جرد الأحبار"),
              ("purchases", "المشتريات"), ("employees", "ملاحظات وأخطاء العاملين"), ("notes", "ملاحظات وتعليمات الوحدات"),
              ("attend", "الحضور والانصراف"),
              ("govstaff", "قاعدة بيانات موظفي الجهات"), ("gov", "حضور وانصراف الجهات"), ("compare", "مقارنة الفترات"), ("reports", "استخراج وطباعة التقارير")]
GOV_ENTITIES = ["السجل التجاري", "نيابة الأسرة", "المرور", "التأمينات الاجتماعية",
                "الجوازات", "الشهر العقاري", "الأحوال المدنية", "الأدلة الجنائية"]
VIEW_ONLY_PAGES = set()                         # أقسام بلا إضافة/تعديل في الواجهة (حاليًا لا يوجد)
# ملاحظات وتعليمات الوحدات (من السيرفر للوحدات، ومعها زر «قدمها»)
NOTE_TARGET_ALL = "كل الوحدات"
NOTE_PRIORITIES = ["عادية", "عاجلة"]
NOTE_DEFAULT_TARGETS = ["الأمن", "الاستقبال", "الصيانة"]


def hash_pw(pw, salt):
    return hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), bytes.fromhex(salt), 120000).hex()


def authenticate_local(username, password):
    """التحقق من كلمة المرور على قاعدة البيانات المحلية (يُستخدم في وضع السيرفر).

    يتصل بالملف المحلي مباشرة (بدون db) حتى لا يتأثر بوضع المحطة.
    """
    con = sqlite3.connect(DB_PATH)
    try:
        r = con.execute("SELECT id,username,pw_hash,salt,is_admin,active,perms FROM users WHERE username=?",
                        (username.strip(),)).fetchone()
    finally:
        con.close()
    if not r or not r[2] or not r[3] or not hmac.compare_digest(hash_pw(password, r[3]), r[2]):
        return None, "wrong"
    if not r[5]:
        return None, "disabled"
    try:
        perms = json.loads(r[6] or "{}")
    except ValueError:
        perms = {}
    return {"id": r[0], "username": r[1], "is_admin": bool(r[4]), "perms": perms}, ""


def authenticate(username, password):
    """يرجع (بيانات المستخدم, "") عند النجاح أو (None, سبب الفشل).

    في وضع «محطة» تتحقق كلمة المرور عند السيرفر نفسه (لا تُنقل أي بيانات مستخدمين).
    """
    if netmod and NET_MODE == netmod.MODE_CLIENT:
        try:
            return netmod.login(NET_CFG, username, password)
        except Exception:
            return None, "offline"
    return authenticate_local(username, password)


def _ident(name):
    """يتحقق أن اسم جدول/عمود آمن قبل إدخاله في نص SQL (حماية من الحقن)."""
    if not isinstance(name, str) or not _IDENT_RE.match(name):
        raise ValueError("اسم جدول أو عمود غير صالح: %r" % (name,))
    return name


def admin_password_is_default():
    """True لو حساب المدير غير موجود أو ما زال على كلمة المرور القديمة/الفارغة."""
    if netmod and NET_MODE == netmod.MODE_CLIENT:
        return False                    # كلمة مرور المدير يديرها الجهاز الرئيسي
    try:
        with db() as con:
            r = con.execute("SELECT pw_hash, salt FROM users WHERE username=?", (DEFAULT_ADMIN,)).fetchone()
    except Exception:
        return True
    if not r or not r[0] or not r[1]:
        return True
    try:
        return hmac.compare_digest(hash_pw(LEGACY_DEFAULT_PW, r[1]), r[0])
    except Exception:
        return True


def set_admin_password(password):
    """يحفظ كلمة مرور جديدة لحساب المدير (ينشئ الحساب لو غير موجود)."""
    salt = secrets.token_hex(16)
    pw_hash = hash_pw(password, salt)
    with db() as con:
        row = con.execute("SELECT id FROM users WHERE username=?", (DEFAULT_ADMIN,)).fetchone()
        if row:
            con.execute("UPDATE users SET pw_hash=?, salt=?, is_admin=1, active=1 WHERE id=?",
                        (pw_hash, salt, row[0]))
        else:
            con.execute("INSERT INTO users(username,pw_hash,salt,is_admin,active,perms,created_at) "
                        "VALUES(?,?,?,?,?,?,?)",
                        (DEFAULT_ADMIN, pw_hash, salt, 1, 1, "{}",
                         datetime.now().strftime("%Y-%m-%d %H:%M:%S")))


def validate_password(pw, confirm):
    """يرجع رسالة خطأ أو None لو كلمة المرور مقبولة."""
    if not pw or len(pw) < PW_MIN_LEN:
        return f"كلمة المرور لازم تكون {PW_MIN_LEN} أحرف على الأقل."
    if pw != confirm:
        return "كلمتا المرور غير متطابقتين."
    if pw.strip().lower() in WEAK_PASSWORDS:
        return "كلمة المرور دي شائعة جدًا وسهلة التخمين، اختر واحدة أقوى."
    return None


def ask_new_password(root, heading, note):
    """نافذة لاختيار كلمة مرور جديدة. ترجع كلمة المرور أو None لو أُغلقت."""
    win = tk.Toplevel(root)
    win.title(T(heading))
    win.configure(bg=BG)
    win.resizable(False, False)
    w, h = 480, 340
    win.geometry(f"{w}x{h}+{(win.winfo_screenwidth() - w) // 2}+{max(0, (win.winfo_screenheight() - h) // 2 - 40)}")
    res = {"pw": None}

    tk.Label(win, text=T(heading), bg=BG, fg=INK, font=(FONT, 15, "bold")).pack(pady=(18, 4))
    tk.Label(win, text=T(note), bg=BG, fg=MUTED, font=(FONT, 9), wraplength=430,
             justify="center").pack(padx=20)
    box = tk.Frame(win, bg=BG)
    box.pack(pady=14)
    tk.Label(box, text="كلمة المرور الجديدة", bg=BG, fg=TEXT, font=(FONT, 10, "bold")).grid(row=0, column=0, sticky="e", pady=5)
    e1 = tk.Entry(box, show="•", width=28, font=(FONT, 12), bg=ENTRY, relief="solid", bd=1)
    e1.grid(row=0, column=1, padx=8, pady=5)
    tk.Label(box, text="تأكيد كلمة المرور", bg=BG, fg=TEXT, font=(FONT, 10, "bold")).grid(row=1, column=0, sticky="e", pady=5)
    e2 = tk.Entry(box, show="•", width=28, font=(FONT, 12), bg=ENTRY, relief="solid", bd=1)
    e2.grid(row=1, column=1, padx=8, pady=5)

    def save(event=None):
        err = validate_password(e1.get(), e2.get())
        if err:
            messagebox.showerror("كلمة المرور", T(err), parent=win)
            return
        res["pw"] = e1.get()
        win.destroy()

    def cancel(event=None):
        win.destroy()

    btns = tk.Frame(win, bg=BG)
    btns.pack(pady=4)
    tk.Button(btns, text=T("حفظ كلمة المرور"), command=save, bg=GREEN, fg="white", relief="flat", bd=0,
              font=(FONT, 11, "bold"), padx=22, pady=7, cursor="hand2").pack(side="right", padx=6)
    tk.Button(btns, text=T("إلغاء"), command=cancel, bg=MUTED, fg="white", relief="flat", bd=0,
              font=(FONT, 10, "bold"), padx=18, pady=7, cursor="hand2").pack(side="right", padx=6)
    win.protocol("WM_DELETE_WINDOW", cancel)
    win.bind("<Return>", save)
    win.bind("<Escape>", cancel)
    win.update_idletasks()
    win.deiconify()
    win.lift()
    win.grab_set()
    e1.focus_set()
    root.wait_window(win)
    return res["pw"]


def ensure_admin_password(root):
    """أول تشغيل: يطلب كلمة مرور للمدير بدلًا من أي كلمة مرور افتراضية.

    ترجع True لو الحساب جاهز، و False لو المستخدم أغلق النافذة.
    """
    if netmod and NET_MODE == netmod.MODE_CLIENT:
        return True                     # كلمة مرور المدير تُطلب على الجهاز الرئيسي فقط
    if not admin_password_is_default():
        return True
    pw = ask_new_password(root, "أول تشغيل - كلمة مرور المدير",
                          "مفيش كلمة مرور افتراضية جوه البرنامج. اختر كلمة مرور لحساب "
                          f"{DEFAULT_ADMIN} (6 أحرف على الأقل) - هتحتاجها في كل مرة تفتح البرنامج.")
    if not pw:
        return False
    set_admin_password(pw)
    return True


# ============================ قاعدة البيانات ============================
def init_db(con):
    con.execute("""CREATE TABLE IF NOT EXISTS visits(
        id INTEGER PRIMARY KEY AUTOINCREMENT, person TEXT NOT NULL,
        entity TEXT, reason TEXT, created_at TEXT NOT NULL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS maintenance(
        id INTEGER PRIMARY KEY AUTOINCREMENT, person TEXT,
        entity TEXT, notes TEXT NOT NULL, created_at TEXT NOT NULL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS inventory(
        id INTEGER PRIMARY KEY AUTOINCREMENT, week_from TEXT NOT NULL,
        week_to TEXT NOT NULL, item TEXT NOT NULL, quantity REAL NOT NULL,
        notes TEXT DEFAULT '', saved_at TEXT NOT NULL)""")
    cols = [r[1] for r in con.execute("PRAGMA table_info(inventory)")]
    if "category" not in cols:
        con.execute("ALTER TABLE inventory ADD COLUMN category TEXT DEFAULT 'cleaning'")
    con.execute("""CREATE TABLE IF NOT EXISTS employee_notes(
        id INTEGER PRIMARY KEY AUTOINCREMENT, employee TEXT NOT NULL,
        kind TEXT NOT NULL, severity TEXT DEFAULT '', details TEXT NOT NULL,
        action TEXT DEFAULT '', created_at TEXT NOT NULL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS employees(
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE,
        department TEXT NOT NULL, created_at TEXT NOT NULL)""")
    have = [r[1] for r in con.execute("PRAGMA table_info(employees)")]
    for col, ddl in (("job_title", "TEXT DEFAULT ''"), ("national_id", "TEXT DEFAULT ''"),
                     ("phone", "TEXT DEFAULT ''"), ("birth_date", "TEXT DEFAULT ''"),
                     ("address", "TEXT DEFAULT ''"), ("start_date", "TEXT DEFAULT ''"),
                     ("end_date", "TEXT DEFAULT ''"), ("status", f"TEXT DEFAULT '{STATUS_ON}'"),
                     ("notes", "TEXT DEFAULT ''"), ("photo_file", "TEXT DEFAULT ''"),
                     ("id_file", "TEXT DEFAULT ''")):
        if col not in have:
            con.execute(f"ALTER TABLE employees ADD COLUMN {_ident(col)} {ddl}")
    con.execute("""CREATE TABLE IF NOT EXISTS permits(
        id INTEGER PRIMARY KEY AUTOINCREMENT, employee TEXT NOT NULL, permit_date TEXT NOT NULL,
        out_time TEXT NOT NULL, back_time TEXT DEFAULT '', reason TEXT DEFAULT '', created_at TEXT NOT NULL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS absences(
        id INTEGER PRIMARY KEY AUTOINCREMENT, employee TEXT NOT NULL, absence_date TEXT NOT NULL,
        kind TEXT NOT NULL, arrival_time TEXT DEFAULT '', late_minutes INTEGER DEFAULT 0,
        notes TEXT DEFAULT '', created_at TEXT NOT NULL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS stock_in(
        id INTEGER PRIMARY KEY AUTOINCREMENT, category TEXT NOT NULL, item TEXT NOT NULL, qty REAL NOT NULL,
        in_date TEXT NOT NULL, notes TEXT DEFAULT '', created_at TEXT NOT NULL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS attendance(
        id INTEGER PRIMARY KEY AUTOINCREMENT, employee TEXT NOT NULL, att_date TEXT NOT NULL,
        in_time TEXT DEFAULT '', out_time TEXT DEFAULT '', notes TEXT DEFAULT '', created_at TEXT NOT NULL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS gov_staff(
        id INTEGER PRIMARY KEY AUTOINCREMENT, entity TEXT NOT NULL, name TEXT NOT NULL,
        job_title TEXT DEFAULT '', phone TEXT DEFAULT '', notes TEXT DEFAULT '', created_at TEXT NOT NULL,
        national_id TEXT DEFAULT '', status TEXT DEFAULT 'على رأس العمل')""")
    gc = [r[1] for r in con.execute("PRAGMA table_info(gov_staff)")]
    if "national_id" not in gc:
        con.execute("ALTER TABLE gov_staff ADD COLUMN national_id TEXT DEFAULT ''")
    if "status" not in gc:
        con.execute("ALTER TABLE gov_staff ADD COLUMN status TEXT DEFAULT 'على رأس العمل'")
    con.execute("""CREATE TABLE IF NOT EXISTS gov_attendance(
        id INTEGER PRIMARY KEY AUTOINCREMENT, entity TEXT NOT NULL, person TEXT NOT NULL, att_date TEXT NOT NULL,
        in_time TEXT DEFAULT '', out_time TEXT DEFAULT '', notes TEXT DEFAULT '', created_at TEXT NOT NULL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS purchases(
        id INTEGER PRIMARY KEY AUTOINCREMENT, category TEXT NOT NULL, item TEXT NOT NULL, qty REAL NOT NULL,
        unit_price REAL DEFAULT 0, total REAL DEFAULT 0, supplier TEXT DEFAULT '', invoice_no TEXT DEFAULT '',
        p_date TEXT NOT NULL, notes TEXT DEFAULT '', invoice_file TEXT DEFAULT '', stock_in_id INTEGER,
        created_at TEXT NOT NULL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS notes(
        id INTEGER PRIMARY KEY AUTOINCREMENT, body TEXT NOT NULL,
        target TEXT DEFAULT '', priority TEXT DEFAULT 'عادية',
        created_by TEXT DEFAULT '', created_at TEXT NOT NULL,
        status TEXT DEFAULT 'new', done_by TEXT DEFAULT '', done_at TEXT DEFAULT '')""")
    con.execute("""CREATE TABLE IF NOT EXISTS users(
        id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL UNIQUE COLLATE NOCASE,
        pw_hash TEXT NOT NULL, salt TEXT NOT NULL, is_admin INTEGER DEFAULT 0, active INTEGER DEFAULT 1,
        perms TEXT DEFAULT '{}', created_at TEXT NOT NULL)""")
    if not con.execute("SELECT 1 FROM users LIMIT 1").fetchone():      # أول تشغيل: حساب مدير بدون كلمة مرور
        # كلمة المرور تُطلب من المستخدم عند أول تشغيل (ensure_admin_password)
        # بدلًا من وجود كلمة مرور افتراضية معروفة داخل البرنامج.
        con.execute("INSERT INTO users(username,pw_hash,salt,is_admin,active,perms,created_at) VALUES(?,?,?,?,?,?,?)",
                    (DEFAULT_ADMIN, "", "", 1, 1, "{}",
                     datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    # أعمدة أضيفت لاحقًا (تُضاف تلقائيًا لقواعد البيانات القديمة)
    pc = [r[1] for r in con.execute("PRAGMA table_info(purchases)")]
    if "paid_by" not in pc:
        con.execute("ALTER TABLE purchases ADD COLUMN paid_by TEXT DEFAULT ''")
    sc = [r[1] for r in con.execute("PRAGMA table_info(stock_in)")]
    if "invoice_file" not in sc:
        con.execute("ALTER TABLE stock_in ADD COLUMN invoice_file TEXT DEFAULT ''")
    # فهارس لتسريع البحث والتقارير مع كبر حجم البيانات
    for sql in (
        "CREATE INDEX IF NOT EXISTS ix_visits_t ON visits(created_at)",
        "CREATE INDEX IF NOT EXISTS ix_maint_t ON maintenance(created_at)",
        "CREATE INDEX IF NOT EXISTS ix_notes_t ON employee_notes(created_at)",
        "CREATE INDEX IF NOT EXISTS ix_notes_e ON employee_notes(employee)",
        "CREATE INDEX IF NOT EXISTS ix_inv ON inventory(category, week_from, week_to)",
        "CREATE INDEX IF NOT EXISTS ix_stock ON stock_in(category, item, in_date)",
        "CREATE INDEX IF NOT EXISTS ix_perm ON permits(employee, permit_date)",
        "CREATE INDEX IF NOT EXISTS ix_abs ON absences(employee, absence_date)",
        "CREATE INDEX IF NOT EXISTS ix_att ON attendance(employee, att_date)",
        "CREATE INDEX IF NOT EXISTS ix_gov_att ON gov_attendance(entity, att_date)",
        "CREATE INDEX IF NOT EXISTS ix_gov_staff ON gov_staff(entity, name)",
        "CREATE INDEX IF NOT EXISTS ix_purch ON purchases(category, p_date)",
        "CREATE INDEX IF NOT EXISTS ix_notes ON notes(status, id)",
    ):
        con.execute(sql)


@contextmanager
def db():
    """اتصال يُغلق دائمًا ويحفظ التغييرات تلقائيًا.

    في وضع «محطة» يتم توجيه الاتصال للسيرفر عبر الشبكة بدل الملف المحلي،
    فتشترك كل الأجهزة في قاعدة بيانات واحدة على الجهاز الرئيسي.
    """
    if netmod and NET_MODE == netmod.MODE_CLIENT:
        con = netmod.RemoteConnection(NET_CFG, queue_path=NET_QUEUE)
        try:
            yield con
        finally:
            con.close()
        return
    con = sqlite3.connect(DB_PATH)
    try:
        init_db(con)
        yield con
        con.commit()
    finally:
        con.close()


BACKUP_DIR = os.path.join(BASE, "backups")


def backup_now(tag=None):
    """نسخة احتياطية متسقة من قاعدة البيانات داخل مجلد backups."""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    name = f"center_management_{tag or datetime.now().strftime('%Y-%m-%d_%H%M%S')}.db"
    dest = os.path.join(BACKUP_DIR, name)
    src, dst = sqlite3.connect(DB_PATH), sqlite3.connect(dest)
    try:
        src.backup(dst)
    finally:
        dst.close(); src.close()
    return dest


def auto_backup(keep=14):
    """نسخة تلقائية مرة كل يوم عند فتح البرنامج، وتُحفظ آخر 14 نسخة يومية."""
    if netmod and NET_MODE == netmod.MODE_CLIENT:
        return                          # النسخ الاحتياطي مسؤولية الجهاز الرئيسي
    if not os.path.exists(DB_PATH):
        return
    tag = date.today().strftime("%Y-%m-%d")
    if os.path.exists(os.path.join(BACKUP_DIR, f"center_management_{tag}.db")):
        return
    backup_now(tag)
    daily = sorted(f for f in os.listdir(BACKUP_DIR)
                   if f.startswith("center_management_") and len(f) == len(f"center_management_{tag}.db"))
    for f in daily[:-keep]:
        os.remove(os.path.join(BACKUP_DIR, f))


FULL_PREFIX = "full_"


def backup_full_now(tag=None):
    """نسخة كاملة: قاعدة البيانات + كل صور العاملين والفواتير في ملف zip واحد.

    تُحفظ داخل مجلد backups باسم full_التاريخ.zip. النسخة المتسقة من القاعدة
    تُؤخذ عبر واجهة SQLite الرسمية، ثم تُضاف الصور من فولدر photos.
    """
    os.makedirs(BACKUP_DIR, exist_ok=True)
    name = f"{FULL_PREFIX}{tag or datetime.now().strftime('%Y-%m-%d_%H%M%S')}.zip"
    dest = os.path.join(BACKUP_DIR, name)
    tmpdir = tempfile.mkdtemp(prefix="khm_full_")
    try:
        tmpdb = os.path.join(tmpdir, "center_management.db")
        src, dst = sqlite3.connect(DB_PATH), sqlite3.connect(tmpdb)
        try:
            src.backup(dst)
        finally:
            dst.close(); src.close()
        nphotos = 0
        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
            z.write(tmpdb, "center_management.db")
            if os.path.isdir(PHOTOS_DIR):
                for root_, _, files in os.walk(PHOTOS_DIR):
                    for fn in files:
                        if os.path.basename(root_) == "thumbs":      # الصور المصغرة تُعاد تلقائيًا
                            continue
                        p = os.path.join(root_, fn)
                        z.write(p, os.path.join("photos", os.path.relpath(p, PHOTOS_DIR)))
                        nphotos += 1
            info = (f"خدمات مصر — نسخة كاملة\n"
                    f"الإصدار: {APP_VERSION}\n"
                    f"التاريخ: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                    f"الصور: {nphotos} ملف\n")
            z.writestr("BACKUP_INFO.txt", info)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    return dest


def backup_kind(name):
    """يصنّف ملف النسخة الاحتياطية: كاملة / أمان قبل الاسترجاع / يدوية / يومية."""
    if name.startswith(FULL_PREFIX):
        return "كاملة"
    if name.startswith("center_management_before_restore_"):
        return "أمان قبل الاسترجاع"
    daily_len = len("center_management_YYYY-MM-DD.db")
    if name.endswith(".db") and len(name) == daily_len:
        return "يومية"
    return "يدوية"


def list_backups():
    """كل النسخ الاحتياطية في مجلد backups مرتبة من الأحدث للأقدم."""
    if not os.path.isdir(BACKUP_DIR):
        return []
    out = []
    for fn in os.listdir(BACKUP_DIR):
        if not (fn.endswith(".db") or fn.endswith(".zip")):
            continue
        p = os.path.join(BACKUP_DIR, fn)
        if not os.path.isfile(p):
            continue
        try:
            st = os.stat(p)
        except OSError:
            continue
        out.append({"name": fn, "path": p, "size": st.st_size,
                    "mtime": st.st_mtime, "kind": backup_kind(fn),
                    "full": fn.endswith(".zip")})
    out.sort(key=lambda r: r["mtime"], reverse=True)
    return out


def validate_backup_db(path):
    """يتأكد أن الملف نسخة سليمة من قاعدة بيانات البرنامج. يرجع (ok, msg, عدد السجلات)."""
    try:
        chk = sqlite3.connect(path)
        try:
            tables = {r[0] for r in chk.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            ok = chk.execute("PRAGMA integrity_check").fetchone()[0]
            nrec = sum(chk.execute(f"SELECT COUNT(*) FROM {_ident(t)}").fetchone()[0]
                       for t in ("visits", "maintenance", "inventory", "employees") if t in tables)
        finally:
            chk.close()
    except sqlite3.Error as ex:
        return False, f"تعذر قراءة الملف كقاعدة بيانات:\n{ex}", 0
    if not {"visits", "maintenance", "inventory"} <= tables or ok != "ok":
        return False, "هذا الملف ليس نسخة سليمة من قاعدة بيانات البرنامج.", 0
    return True, "", nrec


def open_backup_archive(path):
    """يفتح نسخة احتياطية (.db أو .zip). يرجع (db_file, photos_src, tmpdir).

    tmpdir لازم يتحذف بعد الاستخدام (لو مش None). يرفع ValueError لو الملف غير صالح.
    """
    if not path.lower().endswith(".zip"):
        return path, None, None
    tmpdir = tempfile.mkdtemp(prefix="khm_restore_")
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            dbname = (next((n for n in names if os.path.basename(n) == "center_management.db"), None)
                      or next((n for n in names if n.lower().endswith(".db")), None))
            if not dbname:
                raise ValueError("الملف المضغوط لا يحتوي على قاعدة بيانات البرنامج.")
            z.extract(dbname, tmpdir)
            db_file = os.path.join(tmpdir, dbname)
            photos = [n for n in names if n.startswith("photos/") and not n.endswith("/")]
            photos_src = None
            if photos:
                z.extractall(tmpdir, photos)
                photos_src = os.path.join(tmpdir, "photos")
    except Exception:
        shutil.rmtree(tmpdir, ignore_errors=True)
        raise
    return db_file, photos_src, tmpdir


def do_restore(db_file, photos_src=None):
    """يستبدل قاعدة البيانات الحالية بالنسخة، مع نسخة أمان وحفظ المستخدمين.

    photos_src: مجلد صور يُدمج (للنسخ الكاملة) أو None. يرجع مسار نسخة الأمان.
    يرفع استثناء عند الفشل (بعد إرجاع البيانات الحالية).
    """
    os.makedirs(BACKUP_DIR, exist_ok=True)
    safety = backup_now("before_restore_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    try:
        with db() as con:                   # المستخدمون وكلمات المرور لا تتغير بإدراج نسخة احتياطية
            users_keep = con.execute("SELECT id,username,pw_hash,salt,is_admin,active,perms,created_at FROM users").fetchall()
        src, dst = sqlite3.connect(db_file), sqlite3.connect(DB_PATH)
        try:
            src.backup(dst)
        finally:
            dst.close(); src.close()
        try:
            with db() as con:               # تهيئة/ترقية الجداول؛ لو فشلت نرجع للبيانات السابقة
                if users_keep:
                    con.execute("DELETE FROM users")
                    con.executemany("INSERT INTO users(id,username,pw_hash,salt,is_admin,active,perms,created_at) "
                                    "VALUES(?,?,?,?,?,?,?,?)", users_keep)
        except sqlite3.Error:
            src, dst = sqlite3.connect(safety), sqlite3.connect(DB_PATH)
            try:
                src.backup(dst)
            finally:
                dst.close(); src.close()
            raise
    except Exception:
        raise
    if photos_src and os.path.isdir(photos_src):     # دمج الصور بعد نجاح القاعدة (لا نحذف صورًا أحدث)
        os.makedirs(PHOTOS_DIR, exist_ok=True)
        for root_, _, files in os.walk(photos_src):
            for fn in files:
                s = os.path.join(root_, fn)
                d = os.path.join(PHOTOS_DIR, os.path.relpath(s, photos_src))
                try:
                    os.makedirs(os.path.dirname(d), exist_ok=True)
                    shutil.copy2(s, d)
                except OSError:
                    pass
    return safety


def db_info():
    """حجم قاعدة البيانات (ميجابايت) وإجمالي السجلات."""
    if netmod and NET_MODE == netmod.MODE_CLIENT:
        try:
            info = netmod.server_info(NET_CFG)
            return info.get("db_mb", 0), info.get("total", 0), info.get("photos_mb", 0)
        except Exception:
            return 0, 0, 0
    size = os.path.getsize(DB_PATH) / 1048576 if os.path.exists(DB_PATH) else 0
    with db() as con:
        total = sum(con.execute(f"SELECT COUNT(*) FROM {_ident(t)}").fetchone()[0]
                    for t in ("visits", "maintenance", "inventory", "employee_notes", "employees"))
    return size, total, photos_size_mb()


PHOTOS_DIR = os.path.join(BASE, "photos")


def save_photo(src, prefix):
    """ينسخ الصورة إلى مجلد photos ويرجع اسم الملف المحفوظ.

    في وضع «محطة» تُرفع الصورة للجهاز الرئيسي حتى تظهر لكل الأجهزة.
    """
    os.makedirs(PHOTOS_DIR, exist_ok=True)
    ext = os.path.splitext(src)[1].lower() or ".jpg"
    name = f"{prefix}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}{ext}"
    shutil.copy2(src, os.path.join(PHOTOS_DIR, name))
    if netmod and NET_MODE == netmod.MODE_CLIENT:
        try:
            netmod.upload_photo(NET_CFG, os.path.join(PHOTOS_DIR, name), name)
        except Exception:
            pass
    return name


def ensure_photo(name):
    """يتأكد أن الصورة موجودة محليًا؛ في وضع المحطة ينزّلها من السيرفر عند الحاجة."""
    if not name:
        return None
    path = os.path.join(PHOTOS_DIR, name)
    if os.path.exists(path):
        return path
    if netmod and NET_MODE == netmod.MODE_CLIENT:
        try:
            netmod.fetch_photo(NET_CFG, name, path)
            return path
        except Exception:
            return None
    return None


def make_thumb(name, size):
    """يرجع مسار صورة PNG مصغرة (ينشئها عند الحاجة عبر ويندوز) أو None لو تعذر."""
    src = ensure_photo(name)
    if not src or not os.path.exists(src):
        return None
    tdir = os.path.join(PHOTOS_DIR, "thumbs")
    os.makedirs(tdir, exist_ok=True)
    dst = os.path.join(tdir, f"{size}_{os.path.splitext(name)[0]}.png")
    if os.path.exists(dst):
        return dst
    if sys.platform == "win32":
        q = lambda p: p.replace("'", "''")
        ps = ("Add-Type -AssemblyName System.Drawing;"
              f"$i=[System.Drawing.Image]::FromFile('{q(src)}');"
              f"$r=[Math]::Min({size}/$i.Width,{size}/$i.Height);"
              "$w=[Math]::Max(1,[int]($i.Width*$r));$h=[Math]::Max(1,[int]($i.Height*$r));"
              "$b=New-Object System.Drawing.Bitmap $w,$h;"
              "$g=[System.Drawing.Graphics]::FromImage($b);"
              "$g.InterpolationMode='HighQualityBicubic';$g.DrawImage($i,0,0,$w,$h);"
              f"$b.Save('{q(dst)}',[System.Drawing.Imaging.ImageFormat]::Png);"
              "$g.Dispose();$b.Dispose();$i.Dispose()")
        try:
            subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                           capture_output=True, timeout=60,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except Exception:
            pass
        if os.path.exists(dst):
            return dst
    return None


def photos_size_mb():
    total = 0
    for root_, _, files in os.walk(PHOTOS_DIR):
        total += sum(os.path.getsize(os.path.join(root_, f)) for f in files)
    return total / 1048576


def today_text():
    return date.today().strftime("%d/%m/%Y")


def parse_date(s):
    return datetime.strptime(s.strip(), "%d/%m/%Y").date()


def fmt_stamp(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %I:%M %p")
    except (ValueError, TypeError):
        return s


def fmt_iso_date(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d").strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return s


def fmt_time(hhmm):
    """'13:30' -> '01:30 م'"""
    if not hhmm:
        return ""
    try:
        h, m = map(int, hhmm.split(":"))
    except ValueError:
        return hhmm
    return f"{(h % 12) or 12:02d}:{m:02d} {'ص' if h < 12 else 'م'}"


def time_options(h0, h1, step):
    return [fmt_time(f"{h:02d}:{m:02d}") for h in range(h0, h1 + 1) for m in range(0, 60, step)
            if not (h == h1 and m > 0)]


def parse_time_label(label):
    """'01:30 م' -> '13:30' (والفارغ يرجع فارغًا)"""
    label = (label or "").strip()
    if not label:
        return ""
    t, ap = label.split()
    h, m = map(int, t.split(":"))
    if ap == "م" and h != 12:
        h += 12
    if ap == "ص" and h == 12:
        h = 0
    return f"{h:02d}:{m:02d}"


def _mins(hhmm):
    h, m = map(int, hhmm.split(":"))
    return h * 60 + m


def late_minutes(dept, arrival):
    return max(0, _mins(arrival) - _mins(DEPT_START.get(dept, "07:00")))


def worker_dept(name):
    with db() as con:
        r = con.execute("SELECT department FROM employees WHERE name=?", (name,)).fetchone()
    return r[0] if r else ""


def permits_in_month(con, employee, iso_date):
    return con.execute("SELECT COUNT(*) FROM permits WHERE employee=? AND substr(permit_date,1,7)=?",
                       (employee, iso_date[:7])).fetchone()[0]


def add_permit(employee, iso_date, out_t, back_t="", reason=""):
    """يسجل إذن خروج. الحد الأقصى PERMIT_LIMIT في الشهر. يرجع (True, رقم الإذن بالشهر) أو (False, رسالة)."""
    with db() as con:
        n = permits_in_month(con, employee, iso_date)
        if n >= PERMIT_LIMIT:
            y, m = int(iso_date[:4]), int(iso_date[5:7])
            return False, (f"لا يسمح بإذن ثالث!\nالعامل «{employee}» استخدم إذنيه بالفعل خلال شهر {MONTHS_AR[m - 1]} {y}.\n"
                           f"الحد الأقصى {PERMIT_LIMIT} إذنين خروج في الشهر.")
        con.execute("INSERT INTO permits(employee,permit_date,out_time,back_time,reason,created_at) VALUES(?,?,?,?,?,?)",
                    (employee, iso_date, out_t, back_t, reason, now_stamp()))
        return True, n + 1


def add_absence(employee, iso_date, kind, arrival="", notes=""):
    """يسجل غيابًا أو تأخيرًا (حساب دقائق التأخير من بداية دوام قسم العامل). يرجع (True, دقائق) أو (False, رسالة)."""
    late = 0
    if kind == "تأخير":
        if not arrival:
            return False, "اختر وقت حضور العامل لحساب مدة التأخير."
        late = late_minutes(worker_dept(employee), arrival)
        if late <= 0:
            return False, "العامل حضر في الميعاد أو قبله، لا يوجد تأخير."
    else:
        arrival = ""
    with db() as con:
        if con.execute("SELECT 1 FROM absences WHERE employee=? AND absence_date=? AND kind=?",
                       (employee, iso_date, kind)).fetchone():
            return False, "هذا السجل مسجل بالفعل لنفس العامل في نفس اليوم."
        con.execute("""INSERT INTO absences(employee,absence_date,kind,arrival_time,late_minutes,notes,created_at)
            VALUES(?,?,?,?,?,?,?)""", (employee, iso_date, kind, arrival, late, notes, now_stamp()))
    return True, late


def parse_time_flex(text, is_out=False, after=None):
    """يقبل كتابة يدوية مرنة: 7 / 7:30 / 07:30 / 4 م / 04:00 pm / 16:00 ... ويرجع 'HH:MM' أو '' للفارغ.
    لو مفيش ص/م: الانصراف الأصغر من الحضور بيتحسب مساءً (4 => 16:00، 7 => 19:00)."""
    t = (text or "").strip().lower()
    if not t:
        return ""
    t = t.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")).replace("：", ":").replace(".", ":")
    ap = ""
    for tok, val in (("pm", "م"), ("am", "ص"), ("م", "م"), ("ص", "ص")):
        if t.endswith(tok):
            ap, t = val, t[:-len(tok)].strip(); break
    parts = t.split(":")
    try:
        h = int(parts[0]); m = int(parts[1]) if len(parts) > 1 and parts[1] else 0
    except ValueError:
        raise ValueError("اكتب الوقت بالشكل 7:30 أو 4:00 م")
    if not (0 <= h <= 23 and 0 <= m <= 59) or len(parts) > 2:
        raise ValueError("وقت غير صحيح")
    if ap == "م" and h < 12: h += 12
    if ap == "ص" and h == 12: h = 0
    hhmm = f"{h:02d}:{m:02d}"
    if not ap and is_out and after and h < 12 and hhmm <= after:
        hhmm = f"{h + 12:02d}:{m:02d}"
    return hhmm


def shift_of(dept):
    return DEPT_SHIFT.get(dept, ("07:00", "16:00"))


def attendance_calc(dept, in_t, out_t):
    """يرجع (دقائق العمل، دقائق التأخير عن البداية، دقائق الانصراف المبكر)."""
    s0, s1 = shift_of(dept)
    worked = _mins(out_t) - _mins(in_t) if in_t and out_t else None
    late = max(0, _mins(in_t) - _mins(s0)) if in_t else 0
    early = max(0, _mins(s1) - _mins(out_t)) if out_t else 0
    return worked, late, early


def fmt_hours(mins):
    if mins is None:
        return "-"
    return f"{mins // 60}:{mins % 60:02d} س"


# ============ استخراج وإدراج بيانات الحضور والانصراف (ملف Excel / CSV) ============
_H_NAME = {"اسم العامل", "العامل", "الاسم", "اسم الموظف", "الموظف", "اسم الشخص", "اسم"}
_H_ENT = {"الجهة", "جهة", "اسم الجهة"}
_H_DATE = {"التاريخ", "تاريخ", "اليوم"}
_H_IN = {"وقت الحضور", "الحضور", "حضور", "وقت الدخول", "الدخول"}
_H_OUT = {"وقت الانصراف", "الانصراف", "انصراف", "وقت الخروج", "الخروج"}
_H_NOTE = {"ملاحظات", "ملاحظة"}
_AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")


def _hnorm(v):
    return " ".join(str(v or "").replace("*", " ").split())


def _ent_key(v):
    t = _hnorm(v)
    return t[2:] if t.startswith("ال") else t


def read_sheet_rows(path):
    """يقرأ ملف .xlsx أو .csv ويرجع قائمة صفوف."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        wb = load_workbook(path, data_only=True, read_only=True)
        try:
            ws = wb["الحضور والانصراف"] if "الحضور والانصراف" in wb.sheetnames else wb.worksheets[0]
            return [list(r) for r in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
    if ext == ".csv":
        raw = open(path, "rb").read()
        for enc in ("utf-8-sig", "cp1256"):
            try:
                text = raw.decode(enc); break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError("تعذر قراءة ترميز ملف CSV.")
        first = text.splitlines()[0] if text.strip() else ""
        delim = max((",", ";", "\t"), key=first.count)
        return [list(r) for r in csv.reader(text.splitlines(), delimiter=delim)]
    raise ValueError("نوع الملف غير مدعوم. استخدم ملف Excel (.xlsx) أو CSV.")


def _cell_date(v):
    """يحوّل خلية تاريخ (تاريخ إكسل / نص / رقم تسلسلي) إلى yyyy-mm-dd."""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return (date(1899, 12, 30) + timedelta(days=int(v))).isoformat()
    t = str(v or "").strip().translate(_AR_DIGITS)
    if not t:
        raise ValueError("التاريخ فارغ")
    t = t.split(" ")[0]
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%Y/%m/%d", "%d/%m/%y"):
        try:
            return datetime.strptime(t, f).date().isoformat()
        except ValueError:
            pass
    raise ValueError(f"تاريخ غير مفهوم «{t}» (اكتبه يوم/شهر/سنة)")


def _cell_time(v, is_out=False, after=None):
    """يحوّل خلية وقت إلى 'HH:MM' أو '' لو فارغة."""
    from datetime import time as _t
    if v is None or (isinstance(v, str) and not v.strip()):
        return ""
    if isinstance(v, datetime):
        v = v.time()
    if isinstance(v, _t):
        return f"{v.hour:02d}:{v.minute:02d}"
    if isinstance(v, (int, float)):
        if isinstance(v, float) and 0 <= v < 1:
            m = round(v * 1440); return f"{(m // 60) % 24:02d}:{m % 60:02d}"
        v = str(int(v)) if float(v).is_integer() else str(v)
    t = str(v).strip().translate(_AR_DIGITS)
    if t.count(":") == 2:                       # 07:30:00 -> 07:30
        t = ":".join(t.split(":")[:2])
    return parse_time_flex(t, is_out, after)


def parse_attendance_rows(rows):
    """يحلل صفوف الملف. يرجع (السجلات الصالحة، الأخطاء [(رقم الصف، السبب)]).
    يلاقي صف العناوين تلقائيًا (أول 15 صف) ويقبل أسماء أعمدة بديلة."""
    hdr_i, cols = None, {}
    for i, r in enumerate(rows[:15]):
        m = {}
        for j, c in enumerate(r):
            h = _hnorm(c)
            for key, names in (("name", _H_NAME), ("ent", _H_ENT), ("date", _H_DATE), ("in", _H_IN),
                               ("out", _H_OUT), ("notes", _H_NOTE)):
                if h in names and key not in m:
                    m[key] = j
        if "name" in m and "date" in m and ("in" in m or "out" in m):
            hdr_i, cols = i, m; break
    if hdr_i is None:
        raise ValueError("لم أجد صف العناوين. لازم الملف يحتوي أعمدة: اسم العامل، التاريخ، وقت الحضور، وقت الانصراف (والملاحظات اختياري).\n"
                         "الأفضل تستخرج ملف من زر «استخراج بيانات» وتستخدمه كنموذج.")
    recs, errs = [], []
    get = lambda r, k: (r[cols[k]] if k in cols and cols[k] < len(r) else None)
    for i in range(hdr_i + 1, len(rows)):
        r, n = rows[i], i + 1
        name = " ".join(str(get(r, "name") or "").split())
        if not name and not any(str(c or "").strip() for c in r):
            continue                                              # صف فاضي
        try:
            if not name:
                raise ValueError("الاسم فارغ")
            iso = _cell_date(get(r, "date"))
            in_t = _cell_time(get(r, "in"))
            out_t = _cell_time(get(r, "out"), True, in_t or None)
            if not in_t and not out_t:
                raise ValueError("لا يوجد وقت حضور أو انصراف")
            if in_t and out_t and out_t <= in_t:
                raise ValueError("وقت الانصراف لازم يكون بعد الحضور")
        except ValueError as ex:
            errs.append((n, f"{name or '-'}: {ex}")); continue
        recs.append({"row": n, "name": name, "ent": _hnorm(get(r, "ent")), "iso": iso, "in": in_t, "out": out_t,
                     "notes": str(get(r, "notes") or "").strip()})
    return recs, errs


def apply_attendance_import(kind, recs, update_existing=False, default_entity=None, dry=False):
    """kind = 'main' (حضور العاملين) أو 'gov' (حضور الجهات). يرجع إحصاءات + أخطاء.
    dry=True: يحسب فقط دون كتابة (للمعاينة قبل التأكيد)."""
    res = {"added": 0, "updated": 0, "dups": 0, "errors": []}
    seen = set()
    with db() as con:
        if kind == "main":
            known = {" ".join(r[0].split()): r[0] for r in con.execute("SELECT name FROM employees")}
        else:
            ents = {_ent_key(e): e for e in GOV_ENTITIES}
        for r in recs:
            if kind == "main":
                name = known.get(r["name"])
                if not name:
                    res["errors"].append((r["row"], f"{r['name']}: العامل غير مسجل في قاعدة بيانات العاملين")); continue
                key = (name, r["iso"])
            else:
                ent = ents.get(_ent_key(r["ent"])) if r["ent"] else default_entity
                if not ent:
                    res["errors"].append((r["row"], f"{r['name']}: الجهة «{r['ent']}» غير معروفة")); continue
                name, key = r["name"], (ent, r["name"], r["iso"])
            if key in seen:
                res["errors"].append((r["row"], f"{r['name']}: مكرر داخل الملف لنفس اليوم")); continue
            seen.add(key)
            if kind == "main":
                old = con.execute("SELECT id FROM attendance WHERE employee=? AND att_date=?", key).fetchone()
            else:
                old = con.execute("SELECT id FROM gov_attendance WHERE entity=? AND person=? AND att_date=?", key).fetchone()
            if old:
                if not update_existing:
                    res["dups"] += 1; continue
                if not dry:
                    con.execute(f"UPDATE {'attendance' if kind == 'main' else 'gov_attendance'} "
                                "SET in_time=?, out_time=?, notes=? WHERE id=?", (r["in"], r["out"], r["notes"], old[0]))
                res["updated"] += 1; continue
            if not dry:
                if kind == "main":
                    con.execute("INSERT INTO attendance(employee,att_date,in_time,out_time,notes,created_at) VALUES(?,?,?,?,?,?)",
                                (name, r["iso"], r["in"], r["out"], r["notes"], now_stamp()))
                else:
                    con.execute("INSERT INTO gov_attendance(entity,person,att_date,in_time,out_time,notes,created_at) "
                                "VALUES(?,?,?,?,?,?,?)", (ent, name, r["iso"], r["in"], r["out"], r["notes"], now_stamp()))
                    if not con.execute("SELECT 1 FROM gov_staff WHERE entity=? AND name=?", (ent, name)).fetchone():
                        con.execute("INSERT INTO gov_staff(entity,name,created_at) VALUES(?,?,?)", (ent, name, now_stamp()))
            res["added"] += 1
    return res


def write_simple_xlsx(path, sheet, headers, rows, widths, date_col=None, help_lines=()):
    """ملف بيانات نظيف (عناوين + صفوف) قابل لإعادة الإدراج، مع ورقة تعليمات."""
    wb = Workbook(); ws = wb.active; ws.title = sheet
    ws.sheet_view.rightToLeft = True
    thin = Side(style="thin", color="C9D3E0")
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.font = Font(bold=True, color="FFFFFF", name="Tahoma"); c.fill = PatternFill("solid", fgColor="1F3A5F")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, r in enumerate(rows, 2):
        for j, v in enumerate(r, 1):
            c = ws.cell(row=i, column=j, value=v)
            c.alignment = Alignment(horizontal="center"); c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            if date_col is not None and j - 1 == date_col:
                c.number_format = "DD/MM/YYYY"
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"
    if help_lines:
        hs = wb.create_sheet("تعليمات"); hs.sheet_view.rightToLeft = True
        hs.column_dimensions["A"].width = 110
        for i, line in enumerate(help_lines, 1):
            hs.cell(row=i, column=1, value=line).alignment = Alignment(horizontal="right", wrap_text=True)
        hs["A1"].font = Font(bold=True, size=13)
    wb.save(path)


ATT_HELP = ["طريقة إدراج بيانات الحضور والانصراف",
            "1) اكتب كل يوم في صف مستقل تحت العناوين في الورقة الأولى (لا تغيّر أسماء الأعمدة).",
            "2) التاريخ: يوم/شهر/سنة مثل 15/03/2026 (أو تاريخ Excel عادي).",
            "3) الوقت: مثل 7:30 أو 07:30 ص أو 4:00 م أو 16:00. الانصراف بدون ص/م وأصغر من الحضور بيتحسب مساءً.",
            "4) اسم العامل لازم يطابق الاسم في قاعدة البيانات (أو اسم الشخص مسجلًا للجهة).",
            "5) لو العامل له سجل في نفس اليوم البرنامج بيسألك: تحديث السجل القديم أم تخطيه.",
            "6) الأعمدة الإضافية في الملف بيتم تجاهلها. الملفات المقبولة: .xlsx و .csv"]


GOV_HELP = [l.replace("العامل", "الموظف") for l in ATT_HELP[:4]] + [
    "4) عمود «الجهة» اختياري: لو تُرك فارغًا تُستخدم الجهة المفتوحة وقت الإدراج (الجهات الثمانية فقط).",
    "5) لو الموظف له سجل في نفس اليوم والجهة البرنامج بيسألك: تحديث السجل القديم أم تخطيه.",
    ATT_HELP[6]]


def add_attendance(employee, iso_date, in_t, out_t, notes=""):
    if not in_t and not out_t:
        return False, "اكتب وقت الحضور أو الانصراف على الأقل."
    if in_t and out_t and out_t <= in_t:
        return False, "وقت الانصراف لازم يكون بعد وقت الحضور."
    with db() as con:
        if con.execute("SELECT 1 FROM attendance WHERE employee=? AND att_date=?", (employee, iso_date)).fetchone():
            return False, "تم تسجيل حضور هذا العامل في نفس اليوم بالفعل. احذف السجل القديم من الجدول لو عايز تعيد تسجيله."
        con.execute("INSERT INTO attendance(employee,att_date,in_time,out_time,notes,created_at) VALUES(?,?,?,?,?,?)",
                    (employee, iso_date, in_t, out_t, notes, now_stamp()))
    return True, ""


def fmt_num(x):
    if x is None or x == "" or x == "-":
        return "-"
    x = float(x)
    return str(int(x)) if x.is_integer() else f"{x:g}"


def stock_status(category):
    """لكل صنف: (كمية آخر جرد، تاريخه، الوارد بعده، الرصيد المتوقع = آخر جرد + الوارد بعد تاريخه)."""
    with db() as con:
        last = {}
        for item, wt, q in con.execute("SELECT item, week_to, quantity FROM inventory WHERE category=? ORDER BY week_to, id",
                                       (category,)):
            last[item] = (wt, q)
        incs = {}
        for item, d, q in con.execute("SELECT item, in_date, qty FROM stock_in WHERE category=?", (category,)):
            incs.setdefault(item, []).append((d, q))
    out = {}
    for item in INV_ITEMS[category]:
        wt, lq = last.get(item, (None, None))
        after = sum(q for d, q in incs.get(item, []) if wt is None or d > wt)
        out[item] = (lq, wt, after, None if (wt is None and not after) else (lq or 0) + after)
    return out


def inventory_rows(con, cat, start, end):
    """صفوف تقرير الجرد مع: الرصيد السابق، الوارد، الاستهلاك = السابق + الوارد - الجرد الجديد."""
    rows = con.execute("""SELECT week_from,week_to,item,quantity,notes,saved_at FROM inventory
        WHERE category=? AND week_to>=? AND week_from<=? ORDER BY week_from,saved_at,item""", (cat, start, end)).fetchall()
    hist = {}
    for item, wt, q in con.execute("SELECT item, week_to, quantity FROM inventory WHERE category=? ORDER BY week_to", (cat,)):
        h = hist.setdefault(item, ([], [])); h[0].append(wt); h[1].append(q)
    inc = {}
    for item, d, q in con.execute("SELECT item, in_date, qty FROM stock_in WHERE category=? ORDER BY in_date", (cat,)):
        h = inc.setdefault(item, ([], [0.0])); h[0].append(d); h[1].append(h[1][-1] + q)

    def incoming(item, after, upto):                       # الوارد في (after, upto]
        ds, cum = inc.get(item, ([], [0.0]))
        hi = cum[bisect.bisect_right(ds, upto)]
        return round(hi - (cum[bisect.bisect_right(ds, after)] if after else 0.0), 3)
    out = []
    for wf, wt, item, q, notes, saved in rows:
        ds, qs = hist.get(item, ([], []))
        i = bisect.bisect_left(ds, wt) - 1                  # آخر جرد قبل نهاية هذه الفترة
        if i >= 0:
            got = incoming(item, ds[i], wt)
            out.append((fmt_iso_date(wf), fmt_iso_date(wt), item, q, qs[i], got, round(qs[i] + got - q, 3),
                        notes, fmt_stamp(saved)))
        else:
            out.append((fmt_iso_date(wf), fmt_iso_date(wt), item, q, "-", incoming(item, None, wt), "-",
                        notes, fmt_stamp(saved)))
    return out


def now_stamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ============================ منطق التقارير ============================
def period_counts(start, end):
    """عدد السجلات داخل الفترة (تواريخ ISO) - تستخدم في المعاينة والملخص."""
    with db() as con:
        def one(sql, *a):
            return con.execute(sql, a).fetchone()[0]
        d = "date(created_at) BETWEEN ? AND ?"
        res = {
            "الزيارات": one(f"SELECT COUNT(*) FROM visits WHERE {d}", start, end),
            "الصيانة": one(f"SELECT COUNT(*) FROM maintenance WHERE {d}", start, end),
            "ملاحظات وأخطاء العاملين": one(f"SELECT COUNT(*) FROM employee_notes WHERE {d}", start, end),
        }
        for cat, name in INV_NAMES.items():
            res[name] = one("SELECT COUNT(*) FROM inventory WHERE category=? AND week_to>=? AND week_from<=?",
                            cat, start, end)
        res["الوارد للمخازن"] = one("SELECT COUNT(*) FROM stock_in WHERE in_date BETWEEN ? AND ?", start, end)
        res["إذونات الخروج"] = one("SELECT COUNT(*) FROM permits WHERE permit_date BETWEEN ? AND ?", start, end)
        res["الحضور والانصراف"] = one("SELECT COUNT(*) FROM attendance WHERE att_date BETWEEN ? AND ?", start, end)
        res["الغياب والتأخير"] = one("SELECT COUNT(*) FROM absences WHERE absence_date BETWEEN ? AND ?", start, end)
        res["حضور موظفي الجهات"] = one("SELECT COUNT(*) FROM gov_attendance WHERE att_date BETWEEN ? AND ?", start, end)
        res["قاعدة بيانات موظفي الجهات (الإجمالي)"] = one("SELECT COUNT(*) FROM gov_staff")
        res["قاعدة العاملين (الإجمالي)"] = one("SELECT COUNT(*) FROM employees")
    return res


def _style_sheet(sh, title, period, headers, widths):
    ncol = len(headers)
    sh.sheet_view.rightToLeft = True
    sh.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    c = sh.cell(1, 1, f"{APP_NAME} - فرع {BRANCH} - {title}")
    c.font = Font(name="Arial", size=15, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="0B1D3A")
    c.alignment = Alignment(horizontal="center", vertical="center")
    sh.row_dimensions[1].height = 30
    sh.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)
    c = sh.cell(2, 1, period)
    c.font = Font(name="Arial", size=10, bold=True, color="0B1D3A")
    c.alignment = Alignment(horizontal="center")
    for i, h in enumerate(headers, 1):
        c = sh.cell(4, i, h)
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2A7FFF")
        c.alignment = Alignment(horizontal="center", vertical="center")
        sh.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    sh.freeze_panes = "A5"
    sh.page_setup.orientation = "landscape"
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
    sh.sheet_properties.pageSetUpPr.fitToPage = True
    sh.oddFooter.center.text = f"{APP_NAME} - تصميم: {DESIGNER}"


def _write_rows(sh, rows, ncol):
    # الأنماط تُنشأ مرة واحدة (أسرع بكثير مع التقارير الكبيرة)
    thin = Border(bottom=Side(style="hair", color="B8C7D6"))
    font = Font(name="Arial", size=10)
    al_txt = Alignment(horizontal="right", vertical="center", wrap_text=True)
    al_num = Alignment(horizontal="center", vertical="center", wrap_text=True)
    zebra = PatternFill("solid", fgColor="F3F7FB")
    for r, row in enumerate(rows, 5):
        for c, val in enumerate(row, 1):
            cell = sh.cell(r, c, val)
            cell.font = font
            cell.alignment = al_txt if isinstance(val, str) else al_num
            cell.border = thin
            if r % 2 == 0:
                cell.fill = zebra
    last = max(sh.max_row, 4)
    sh.auto_filter.ref = f"A4:{get_column_letter(ncol)}{last}"
    # سطر التوقيع أسفل الجدول
    foot = last + 2
    sh.merge_cells(start_row=foot, start_column=1, end_row=foot, end_column=ncol)
    c = sh.cell(foot, 1, f"{APP_NAME} - تصميم: {DESIGNER}")
    c.font = Font(name="Arial", size=9, italic=True, color="6B7F92")
    c.alignment = Alignment(horizontal="center")


def report_sections(kind, start, end):
    """تجمّع بيانات التقرير كأقسام (مشتركة بين Excel والطباعة)."""
    period = f"الفترة: من {fmt_iso_date(start)} إلى {fmt_iso_date(end)}   |   تاريخ الاستخراج: {today_text()}"
    S = []

    def add(name, title, headers, widths, rows, per=None):
        S.append({"name": name, "title": title, "headers": headers, "widths": widths,
                  "rows": rows, "period": per or period})
    with db() as con:
        if kind in ("all", "visits"):
            rows = con.execute("""SELECT person,entity,reason,created_at FROM visits
                WHERE date(created_at) BETWEEN ? AND ? ORDER BY created_at""", (start, end)).fetchall()
            add("الزيارات", "سجل زيارات المركز", ["اسم الشخص", "الجهة", "سبب الزيارة", "تاريخ ووقت الزيارة"],
                [28, 28, 40, 24], [(a, b, c, fmt_stamp(t)) for a, b, c, t in rows])
        if kind in ("all", "maintenance"):
            rows = con.execute("""SELECT person,entity,notes,created_at FROM maintenance
                WHERE date(created_at) BETWEEN ? AND ? ORDER BY created_at""", (start, end)).fetchall()
            add("الصيانة", "سجل أعمال الصيانة", ["الفني / المسؤول", "الجهة / الشركة", "ملاحظات الصيانة", "تاريخ ووقت التسجيل"],
                [28, 28, 40, 24], [(a, b, c, fmt_stamp(t)) for a, b, c, t in rows])
        for cat, name in INV_NAMES.items():
            if kind in ("all", cat):
                add(name, name,
                    ["من تاريخ", "إلى تاريخ", "الصنف", "الكمية (الجرد)", "الرصيد السابق", "الوارد", "الاستهلاك", "ملاحظات", "تاريخ الحفظ"],
                    [14, 14, 30, 14, 14, 12, 12, 28, 22], inventory_rows(con, cat, start, end))
        if kind in ("all", "stockin"):
            rows = con.execute("""SELECT in_date, category, item, qty, notes FROM stock_in
                WHERE in_date BETWEEN ? AND ? ORDER BY in_date, category, item""", (start, end)).fetchall()
            add("الوارد للمخازن", "حركة الوارد (الأصناف المستلمة)",
                ["تاريخ الاستلام", "القسم", "الصنف", "الكمية", "ملاحظات (المورد / رقم الاستلام)"], [16, 18, 32, 12, 40],
                [(fmt_iso_date(d), INV_NAMES.get(c, c), i, q, n) for d, c, i, q, n in rows])
        if kind in ("all", "employees"):
            rows = con.execute("""SELECT n.employee, COALESCE(e.department,'غير مسجل'), n.kind, n.severity,
                    n.details, n.action, n.created_at
                FROM employee_notes n LEFT JOIN employees e ON e.name = n.employee
                WHERE date(n.created_at) BETWEEN ? AND ? ORDER BY n.created_at""", (start, end)).fetchall()
            add("ملاحظات وأخطاء العاملين", "ملاحظات وأخطاء العاملين",
                ["اسم العامل", "القسم", "النوع", "الدرجة", "التفاصيل", "الإجراء المتخذ", "التاريخ والوقت"],
                [24, 12, 14, 12, 46, 30, 24], [(a, b, c, d, e, f, fmt_stamp(t)) for a, b, c, d, e, f, t in rows])
            summ = con.execute("""SELECT n.employee, COALESCE(e.department,'غير مسجل'),
                    SUM(n.kind='ملاحظة'), SUM(n.kind='خطأ'), SUM(n.kind='تنبيه شفهي'), SUM(n.kind='إشادة'), COUNT(*)
                FROM employee_notes n LEFT JOIN employees e ON e.name = n.employee
                WHERE date(n.created_at) BETWEEN ? AND ?
                GROUP BY n.employee ORDER BY SUM(n.kind='خطأ') DESC, COUNT(*) DESC""", (start, end)).fetchall()
            add("ملخص العاملين", "ملخص ملاحظات وأخطاء كل عامل",
                ["اسم العامل", "القسم", "ملاحظات", "أخطاء", "تنبيهات شفهية", "إشادات", "الإجمالي"],
                [28, 12, 14, 14, 16, 14, 14], summ)
        if kind in ("all", "permits", "attall"):
            rows = con.execute("""SELECT p.employee, COALESCE(e.department,'غير مسجل'), p.permit_date, p.out_time,
                    p.back_time, p.reason
                FROM permits p LEFT JOIN employees e ON e.name = p.employee
                WHERE p.permit_date BETWEEN ? AND ? ORDER BY p.permit_date, p.out_time""", (start, end)).fetchall()
            add("إذونات الخروج", "إذونات خروج العاملين",
                ["اسم العامل", "القسم", "تاريخ الإذن", "وقت الخروج", "وقت العودة", "السبب"],
                [26, 12, 16, 14, 14, 40],
                [(a, b, fmt_iso_date(c), fmt_time(d), fmt_time(e_), f) for a, b, c, d, e_, f in rows])
        if kind in ("all", "attendance", "attall"):
            rows = con.execute("""SELECT t.employee, COALESCE(e.department,'غير مسجل'), t.att_date, t.in_time, t.out_time, t.notes
                FROM attendance t LEFT JOIN employees e ON e.name = t.employee
                WHERE t.att_date BETWEEN ? AND ? ORDER BY t.att_date, e.department, t.employee""", (start, end)).fetchall()
            att = []
            for n_, d_, dt_, i_, o_, nt_ in rows:
                w_, l_, e2_ = attendance_calc(d_, i_, o_)
                att.append((n_, DEPT_LABEL.get(d_, d_), fmt_iso_date(dt_), fmt_time(i_) or "-", fmt_time(o_) or "-",
                            fmt_hours(w_), l_ or "", e2_ or "", nt_))
            add("الحضور والانصراف", "سجل الحضور والانصراف",
                ["اسم العامل", "القسم", "التاريخ", "الحضور", "الانصراف", "ساعات العمل", "دقائق التأخير", "دقائق الانصراف المبكر", "ملاحظات"],
                [26, 12, 16, 14, 14, 14, 14, 20, 30], att)
        if kind in ("all", "absences", "attall"):
            rows = con.execute("""SELECT a.employee, COALESCE(e.department,'غير مسجل'), a.absence_date, a.kind,
                    a.arrival_time, a.late_minutes, a.notes
                FROM absences a LEFT JOIN employees e ON e.name = a.employee
                WHERE a.absence_date BETWEEN ? AND ? ORDER BY a.absence_date, a.employee""", (start, end)).fetchall()
            add("الغياب والتأخير", "الغياب والتأخير",
                ["اسم العامل", "القسم", "التاريخ", "النوع", "وقت الحضور", "دقائق التأخير", "ملاحظات"],
                [26, 12, 16, 18, 14, 14, 36],
                [(a, b, fmt_iso_date(c), d, fmt_time(e_), f or "", g) for a, b, c, d, e_, f, g in rows])
        if kind in ("all", "permits", "absences", "attall"):
            agg = {}

            def acc(name, dept):
                return agg.setdefault(name, {"dept": dept, "perm": 0, "abs": 0, "exc": 0, "leave": 0, "late": 0, "lmin": 0})
            for name, dept in con.execute("""SELECT p.employee, COALESCE(e.department,'غير مسجل') FROM permits p
                    LEFT JOIN employees e ON e.name = p.employee WHERE p.permit_date BETWEEN ? AND ?""", (start, end)):
                acc(name, dept)["perm"] += 1
            for name, dept, k, lm in con.execute("""SELECT a.employee, COALESCE(e.department,'غير مسجل'), a.kind, a.late_minutes
                    FROM absences a LEFT JOIN employees e ON e.name = a.employee
                    WHERE a.absence_date BETWEEN ? AND ?""", (start, end)):
                x = acc(name, dept)
                if k == "غياب بدون إذن": x["abs"] += 1
                elif k == "غياب بعذر": x["exc"] += 1
                elif k == "تأخير": x["late"] += 1; x["lmin"] += lm or 0
                else: x["leave"] += 1
            srows = sorted(((n, x["dept"], x["perm"], x["abs"], x["exc"], x["leave"], x["late"], x["lmin"])
                            for n, x in agg.items()), key=lambda r: (-r[3], -r[6], -r[2], r[0]))
            add("ملخص الحضور والغياب", "ملخص الإذونات والغياب والتأخير لكل عامل",
                ["اسم العامل", "القسم", "إذونات الخروج", "غياب بدون إذن", "غياب بعذر", "إجازات", "مرات التأخير", "دقائق التأخير"],
                [28, 12, 14, 14, 12, 12, 14, 14], srows)
        if kind in ("all", "workers"):
            rows = con.execute("""SELECT name, department, job_title, status, start_date, end_date,
                    national_id, phone, birth_date, address, notes
                FROM employees ORDER BY (status=?), department, name""", (STATUS_OFF,)).fetchall()
            add("قاعدة العاملين", "قاعدة بيانات العاملين",
                ["اسم العامل", "القسم", "الوظيفة", "الحالة", "تاريخ بدء العمل", "تاريخ انتهاء العمل",
                 "الرقم القومي", "الهاتف", "تاريخ الميلاد", "العنوان", "ملاحظات"],
                [28, 12, 18, 16, 16, 18, 20, 16, 16, 32, 30],
                [(n, d, j, st, fmt_iso_date(sd), fmt_iso_date(ed), nid, ph, fmt_iso_date(bd), ad, nt)
                 for n, d, j, st, sd, ed, nid, ph, bd, ad, nt in rows],
                f"تاريخ الاستخراج: {today_text()}")
        if kind in ("all", "govatt"):
            S.extend(gov_att_sections(start, end))
        if kind in ("all", "govstaff"):
            rows = con.execute("""SELECT entity, name, job_title, phone, national_id, status, notes
                FROM gov_staff ORDER BY entity, (status=?), name""", (STATUS_OFF,)).fetchall()
            add("قاعدة بيانات موظفي الجهات", "قاعدة بيانات موظفي الجهات",
                ["الجهة", "اسم الموظف", "الوظيفة", "الهاتف", "الرقم القومي", "الحالة", "ملاحظات"],
                [24, 26, 20, 16, 20, 12, 32],
                [(e, n, j or "", p or "", nid or "", st, nt or "") for e, n, j, p, nid, st, nt in rows],
                f"تاريخ الاستخراج: {today_text()}")
    return S


def gov_att_sections(start, end, entity="", person=""):
    """أقسام تقرير حضور موظفي الجهات: التفاصيل + ملخص بالموظف + ملخص بالجهة.

    entity / person فارغان = كل الجهات / كل الموظفين. مشتركة بين الطباعة وExcel.
    """
    period = f"الفترة: من {fmt_iso_date(start)} إلى {fmt_iso_date(end)}   |   تاريخ الاستخراج: {today_text()}"
    if entity:
        period += f"   |   الجهة: {entity}"
    if person:
        period += f"   |   الموظف: {person}"
    sql = ("SELECT entity, person, att_date, in_time, out_time, notes FROM gov_attendance "
           "WHERE att_date BETWEEN ? AND ?")
    args = [start, end]
    if entity:
        sql += " AND entity=?"
        args.append(entity)
    if person:
        sql += " AND person=?"
        args.append(person)
    with db() as con:
        rows = con.execute(sql + " ORDER BY entity, att_date, person", args).fetchall()

    details, by_emp, by_ent = [], {}, {}
    for ent, nm, d, tin, tout, nt in rows:
        mins = _mins(tout) - _mins(tin) if tin and tout else None
        details.append((ent, nm, fmt_iso_date(d), fmt_time(tin) or "-", fmt_time(tout) or "-",
                        fmt_hours(mins) if mins is not None else "-", nt or ""))
        x = by_emp.setdefault((ent, nm), {"days": 0, "closed": 0, "open": 0, "mins": 0})
        if tin or tout:
            x["days"] += 1
        if mins is not None:
            x["closed"] += 1
            x["mins"] += mins
        elif tin:
            x["open"] += 1
        g = by_ent.setdefault(ent, {"staff": set(), "rec": 0, "mins": 0})
        g["rec"] += 1
        if tin or tout:
            g["staff"].add(nm)
        if mins is not None:
            g["mins"] += mins

    emp_rows = [(ent, nm, v["days"], v["closed"], v["open"], fmt_hours(v["mins"]),
                 fmt_hours(round(v["mins"] / v["closed"])) if v["closed"] else "-")
                for (ent, nm), v in by_emp.items()]
    emp_rows.sort(key=lambda r: (r[0], r[1]))
    ent_rows = [(ent, len(v["staff"]), v["rec"], fmt_hours(v["mins"])) for ent, v in by_ent.items()]
    ent_rows.sort(key=lambda r: r[0])

    return [
        {"name": "حضور موظفي الجهات", "title": "تفاصيل حضور وانصراف موظفي الجهات",
         "headers": ["الجهة", "اسم الموظف", "التاريخ", "الدخول", "الخروج", "ساعات العمل", "ملاحظات"],
         "widths": [22, 26, 16, 12, 12, 14, 34], "rows": details, "period": period},
        {"name": "ملخص حضور الموظفين", "title": "ملخص حضور كل موظف",
         "headers": ["الجهة", "اسم الموظف", "أيام مسجلة", "أيام مكتملة", "أيام بدون انصراف",
                     "إجمالي الساعات", "متوسط الساعات/يوم"],
         "widths": [22, 26, 12, 12, 16, 14, 18], "rows": emp_rows, "period": period},
        {"name": "ملخص الجهات", "title": "ملخص الحضور حسب الجهة",
         "headers": ["الجهة", "عدد الموظفين المسجلين", "عدد السجلات", "إجمالي الساعات"],
         "widths": [30, 20, 16, 18], "rows": ent_rows, "period": period},
    ]



def build_report(path, kind, start, end):
    """تنشئ ملف Excel. start/end بصيغة YYYY-MM-DD. ترجع قاموس {اسم الشيت: عدد الصفوف}."""
    wb = Workbook()
    wb.remove(wb.active)
    result = {}
    sections = report_sections(kind, start, end)
    for sec in sections:
        sh = wb.create_sheet(sec["name"])
        _style_sheet(sh, sec["title"], sec["period"], sec["headers"], sec["widths"])
        _write_rows(sh, sec["rows"], len(sec["headers"]))
        result[sec["name"]] = len(sec["rows"])
    if kind == "all":
        period = f"الفترة: من {fmt_iso_date(start)} إلى {fmt_iso_date(end)}   |   تاريخ الاستخراج: {today_text()}"
        sh = wb.create_sheet("ملخص التقرير", 0)
        _style_sheet(sh, "ملخص التقرير", period, ["القسم", "عدد السجلات"], [40, 18])
        _write_rows(sh, [(k, v) for k, v in result.items() if k not in SUMMARY_ONLY], 2)
    if not wb.sheetnames:
        wb.create_sheet("لا توجد بيانات")
    wb.save(path)
    return result


# ============================ الطباعة (HTML جاهز للطباعة) ============================
_HTML = """<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>%%TITLE%%</title><style>
@page{size:A4;margin:12mm 10mm 14mm}
*{box-sizing:border-box}body{font-family:Tahoma,"Segoe UI",Arial,sans-serif;color:#14233b;margin:0;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.head{display:flex;align-items:center;gap:14px;background:linear-gradient(90deg,#0b1d3a,#0f6a86);color:#fff;padding:10px 16px;border-radius:10px;border-bottom:5px solid #f5b82e}
.head img{height:60px;background:#fff;border:3px solid #f5b82e;border-radius:6px}.head h1{margin:0;font-size:21px}.head p{margin:3px 0 0;font-size:12px;color:#cfe3ff}
h2{font-size:15px;color:#0b1d3a;border-right:5px solid #2a7fff;padding-right:8px;margin:18px 0 7px}h2 small{color:#6a7d99;font-weight:normal}
table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:6px}th{background:#13315c;color:#fff;padding:7px 6px}
td{padding:6px;border-bottom:1px solid #e1e8f3;text-align:center}tbody tr:nth-child(even) td{background:#f5f8fe}thead{display:table-header-group}tr{page-break-inside:avoid}
.empty{color:#6a7d99;text-align:center;padding:10px;font-size:12px;border:1px dashed #c8d6ea;border-radius:8px}
.ins{border-right:5px solid #7c5cff;background:#f3f0ff;padding:8px 12px;border-radius:8px;margin:6px 0;font-size:13px}
.bars{margin:6px 0 10px}.br{display:flex;align-items:center;gap:8px;margin:4px 0;font-size:12px}.br .n{width:150px;text-align:right}.br .t{flex:1}
.bar{height:13px;border-radius:4px;margin:2px 0;min-width:2px;color:#fff;font-size:10px;line-height:13px;padding-right:4px}.a{background:#2a7fff}.b{background:#9aa9c2}
.up{color:#12b886;font-weight:bold}.dn{color:#e5484d;font-weight:bold}
.sign{display:flex;justify-content:space-around;margin-top:38px;font-size:12px}.sign div{width:28%;border-top:1px solid #6a7d99;text-align:center;padding-top:6px}
.foot{text-align:center;color:#6a7d99;font-size:10px;margin-top:16px}.noprint{text-align:center;margin:8px}
.noprint button{background:#13315c;color:#fff;border:0;border-radius:8px;padding:9px 22px;font-size:14px;cursor:pointer}
@media print{.noprint{display:none}}
</style></head><body>
<div class="noprint"><button onclick="window.print()">🖨️ طباعة / حفظ PDF</button></div>
<div class="head"><img src="data:image/png;base64,%%LOGO%%"><div><h1>%%APP%% - %%TITLE%%</h1><p>%%SUB%%</p></div></div>
%%BODY%%
<div class="sign"><div>إعداد</div><div>مراجعة</div><div>اعتماد</div></div>
<div class="foot">%%APP%% - تصميم وتطوير: %%DESIGNER%% &nbsp;|&nbsp; تم الاستخراج %%NOW%%</div>
<script>window.addEventListener('load',function(){setTimeout(function(){window.print()},600)})</script></body></html>"""


class _Raw(str):
    """نص HTML جاهز لا يُعاد تهريبه."""


def _e(x):
    if isinstance(x, _Raw):
        return str(x)
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    return html.escape("" if x is None else str(x))


def html_page(title, subtitle, body):
    return (_HTML.replace("%%TITLE%%", _e(title)).replace("%%SUB%%", _e(subtitle)).replace("%%BODY%%", body)
            .replace("%%LOGO%%", "".join(LOGO_B64)).replace("%%APP%%", f"{APP_NAME} - فرع {BRANCH}").replace("%%DESIGNER%%", DESIGNER)
            .replace("%%NOW%%", "<bdi>" + datetime.now().strftime("%d/%m/%Y %I:%M %p") + "</bdi>"))


def html_table(title, headers, rows, count=True):
    if not rows:
        return f'<h2>{_e(title)}</h2><div class="empty">لا توجد بيانات في هذه الفترة</div>'
    th = "".join(f"<th>{_e(h)}</th>" for h in headers)
    tr = "".join("<tr>" + "".join(f"<td><bdi>{_e(c)}</bdi></td>" for c in r) + "</tr>" for r in rows)
    n = f" <small>({len(rows)})</small>" if count else ""
    return f"<h2>{_e(title)}{n}</h2><table><thead><tr>{th}</tr></thead><tbody>{tr}</tbody></table>"


def open_html(html_text, prefix):
    """يحفظ صفحة الطباعة في مجلد print ويفتحها في المتصفح (وتبدأ الطباعة تلقائيًا)."""
    d = os.path.join(BASE, "print")
    os.makedirs(d, exist_ok=True)
    path = os.path.join(d, f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html_text)
    webbrowser.open(Path(path).as_uri())
    return path


def build_sections_excel(path, sections):
    """ينشئ ملف Excel من قائمة أقسام (name/title/headers/widths/rows/period). يرجع {اسم الشيت: عدد الصفوف}."""
    wb = Workbook()
    wb.remove(wb.active)
    result = {}
    for sec in sections:
        sh = wb.create_sheet(sec["name"][:31])
        _style_sheet(sh, sec["title"], sec["period"], sec["headers"], sec["widths"])
        _write_rows(sh, sec["rows"], len(sec["headers"]))
        result[sec["name"]] = len(sec["rows"])
    if not wb.sheetnames:
        wb.create_sheet("لا توجد بيانات")
    wb.save(path)
    return result


def sections_html(title, sections, subtitle=""):
    body = "".join(html_table(x["title"], x["headers"], x["rows"]) for x in sections)
    return html_page(title, subtitle or (sections[0]["period"] if sections else ""), body)


def filter_sections_by_employee(sections, name):
    """يُبقي صفوف العامل المحدد فقط (العمود الأول = اسم العامل)."""
    if not name:
        return sections
    out = []
    for s in sections:
        s = dict(s); s["rows"] = [r for r in s["rows"] if r and r[0] == name]; out.append(s)
    return out


ATT_TITLES = {"attendance": "تقرير الحضور والانصراف", "permits": "تقرير إذونات الخروج",
              "absences": "تقرير الغياب والتأخير", "attall": "تقرير الحضور والإذونات والغياب"}


def inventory_excel_sections(category, typed=None):
    """الجرد الحالي (الموجود فعليًا بالمخزن) كقسم Excel بأرقام حقيقية."""
    typed = typed or {}
    st = stock_status(category)
    rows = []
    for n, item in enumerate(INV_ITEMS[category], 1):
        lq, wt, after, exp = st[item]
        t = str(typed.get(item, "")).strip()
        try:
            f = float(t.replace(",", "."))
            t = int(f) if f.is_integer() else f
        except ValueError:
            pass
        status = "لم يُجرد بعد" if exp is None else ("نفد" if exp <= 0 else "متوفر")
        rows.append((n, item, lq if lq is not None else "-", fmt_iso_date(wt) if wt else "-", after or 0,
                     exp if exp is not None else "-", t, status))
    return [{"name": INV_NAMES[category], "title": f"{INV_NAMES[category]} - الجرد الحالي (الموجود بالمخزن)",
             "headers": ["م", "الصنف", "كمية آخر جرد", "تاريخ آخر جرد", "وارد بعد آخر جرد",
                         "الموجود حاليًا بالمخزن", "الجرد الفعلي", "الحالة"],
             "widths": [6, 34, 16, 18, 18, 22, 16, 16], "rows": rows,
             "period": f"الرصيد حتى تاريخ: {today_text()}"}]


def purchases_excel_sections(cat, start, end):
    rows = purchases_rows(cat, start, end)[::-1]
    data = [(fmt_iso_date(r[1]), PURCH_CATS.get(r[2], r[2]), r[3], r[4], r[5] or 0, r[6] or 0, r[7], r[8],
             r[10] or "", "نعم" if r[9] else "لا") for r in rows]
    total = sum(r[6] or 0 for r in rows)
    if data:
        data.append(("الإجمالي", "", "", "", "", total, "", "", "", ""))
    period = (f"الفترة: من {fmt_iso_date(start)} إلى {fmt_iso_date(end)}   |   النوع: "
              f"{'الكل' if cat == 'all' else PURCH_CATS[cat]}   |   تاريخ الاستخراج: {today_text()}")
    S = [{"name": "المشتريات", "title": "سجل المشتريات",
          "headers": ["التاريخ", "النوع", "الصنف", "الكمية", "سعر الوحدة", "الإجمالي", "المورد", "رقم الفاتورة",
                      "تم السداد بواسطة", "صورة فاتورة"],
          "widths": [14, 12, 32, 12, 14, 16, 26, 16, 22, 14], "rows": data, "period": period}]
    sums = {}
    for r in rows:
        s = sums.setdefault(r[2], [0, 0.0]); s[0] += 1; s[1] += r[6] or 0
    S.append({"name": "ملخص حسب النوع", "title": "ملخص المشتريات حسب النوع",
              "headers": ["النوع", "عدد البنود", "إجمالي المبلغ"], "widths": [20, 16, 20],
              "rows": [(PURCH_CATS.get(k, k), v[0], v[1]) for k, v in sums.items()], "period": period})
    return S


def fmt_money(x):
    try:
        return f"{float(x):,.2f}".rstrip("0").rstrip(".")
    except (TypeError, ValueError):
        return "-"


def inventory_print_html(category, typed=None):
    """صفحة طباعة الجرد الحالي: الرصيد الموجود فعليًا بالمخزن = آخر جرد + الوارد بعده."""
    typed = typed or {}
    st = stock_status(category)
    rows, out_n, uncounted = [], 0, 0
    for n, item in enumerate(INV_ITEMS[category], 1):
        lq, wt, after, exp = st[item]
        if exp is None:
            status = _Raw('<b style="color:#6a7d99">لم يُجرد بعد</b>'); uncounted += 1
        elif exp <= 0:
            status = _Raw('<b style="color:#e5484d">نفد</b>'); out_n += 1
        else:
            status = _Raw('<b style="color:#12b886">متوفر</b>')
        rows.append((n, item, fmt_num(lq), fmt_iso_date(wt) if wt else "-", fmt_num(after) if after else "-",
                     _Raw("<b>" + _e(fmt_num(exp)) + "</b>"), typed.get(item, ""), status))
    headers = ["م", "الصنف", "كمية آخر جرد", "تاريخ آخر جرد", "وارد بعد آخر جرد",
               "الموجود حاليًا بالمخزن", "الجرد الفعلي (للكتابة)", "الحالة"]
    body = html_table(f"{INV_NAMES[category]} - الرصيد الحالي", headers, rows, False)
    body += (f'<div class="ins">عدد الأصناف: <bdi>{len(rows)}</bdi> &nbsp;|&nbsp; أصناف نفدت: <bdi>{out_n}</bdi>'
             f' &nbsp;|&nbsp; لم تُجرد بعد: <bdi>{uncounted}</bdi><br>'
             'الموجود حاليًا = كمية آخر جرد + الوارد المسجل بعده. اترك عمود «الجرد الفعلي» للكتابة اليدوية عند المطابقة.</div>')
    return html_page(f"{INV_NAMES[category]} - الموجود حاليًا", f"الرصيد حتى تاريخ: {today_text()}", body)


def workers_sections(dept="الكل", status="الكل", key=""):
    """قاعدة العاملين بنفس فلاتر الشاشة (القسم / الحالة / البحث)."""
    sql = ("SELECT name, department, job_title, status, start_date, end_date, national_id, phone, birth_date, "
           "address, notes FROM employees WHERE 1=1")
    args = []
    if dept != "الكل":
        sql += " AND department=?"; args.append(dept)
    if status != "الكل":
        sql += " AND status=?"; args.append(status)
    if key:
        sql += " AND (name LIKE ? OR phone LIKE ? OR national_id LIKE ?)"; args += [f"%{key}%"] * 3
    with db() as con:
        rows = con.execute(sql + " ORDER BY (status=?), department, name", args + [STATUS_OFF]).fetchall()
    flt = []
    if dept != "الكل": flt.append(f"القسم: {dept}")
    if status != "الكل": flt.append(f"الحالة: {status}")
    if key: flt.append(f"بحث: {key}")
    period = f"تاريخ الاستخراج: {today_text()}" + ("   |   " + "   |   ".join(flt) if flt else "")
    return [{"name": "قاعدة العاملين", "title": "قاعدة بيانات العاملين",
             "headers": ["اسم العامل", "القسم", "الوظيفة", "الحالة", "تاريخ بدء العمل", "تاريخ انتهاء العمل",
                         "الرقم القومي", "الهاتف", "تاريخ الميلاد", "العنوان", "ملاحظات"],
             "widths": [28, 12, 18, 16, 16, 18, 20, 16, 16, 32, 30],
             "rows": [(n, d, j, st, fmt_iso_date(sd), fmt_iso_date(ed), nid, ph, fmt_iso_date(bd), ad, nt)
                      for n, d, j, st, sd, ed, nid, ph, bd, ad, nt in rows],
             "period": period}]


def earliest_record_date(table):
    """أقدم تاريخ تسجيل في جدول (لتكون الفترة الافتراضية = كل السجلات)."""
    try:
        with db() as con:
            r = con.execute(f"SELECT MIN(date(created_at)) FROM {_ident(table)}").fetchone()
        return datetime.strptime(r[0], "%Y-%m-%d").date() if r and r[0] else date.today()
    except Exception:
        return date.today()


def purchases_rows(cat, start, end):
    q = ("SELECT id,p_date,category,item,qty,unit_price,total,supplier,invoice_no,invoice_file,paid_by FROM purchases "
         "WHERE p_date BETWEEN ? AND ?")
    args = [start, end]
    if cat != "all":
        q += " AND category=?"; args.append(cat)
    with db() as con:
        return con.execute(q + " ORDER BY p_date DESC, id DESC", args).fetchall()


def purchases_html(cat, start, end):
    rows = purchases_rows(cat, start, end)[::-1]
    total = sum(r[6] or 0 for r in rows)
    data = [(fmt_iso_date(r[1]), PURCH_CATS.get(r[2], r[2]), r[3], fmt_num(r[4]), fmt_money(r[5]), fmt_money(r[6]),
             r[7], r[8], r[10] or "", "✓" if r[9] else "-") for r in rows]
    if data:
        data.append((_Raw("<b>الإجمالي</b>"), "", "", "", "", _Raw("<b>" + _e(fmt_money(total)) + "</b>"), "", "", "", ""))
    body = html_table("سجل المشتريات", ["التاريخ", "النوع", "الصنف", "الكمية", "سعر الوحدة", "الإجمالي",
                                         "المورد", "رقم الفاتورة", "تم السداد بواسطة", "صورة فاتورة"], data, False)
    sums = {}
    for r in rows:
        s = sums.setdefault(r[2], [0, 0.0]); s[0] += 1; s[1] += r[6] or 0
    if sums:
        body += html_table("ملخص حسب النوع", ["النوع", "عدد البنود", "إجمالي المبلغ"],
                           [(PURCH_CATS.get(k, k), v[0], fmt_money(v[1])) for k, v in sums.items()], False)
    imgs = []
    for f in dict.fromkeys(r[9] for r in rows if r[9]):
        if f.lower().endswith((".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp")) and os.path.exists(os.path.join(PHOTOS_DIR, f)):
            imgs.append('<div style="display:inline-block;margin:6px;text-align:center;page-break-inside:avoid">'
                        f'<img src="{Path(os.path.join(PHOTOS_DIR, f)).as_uri()}" style="max-width:340px;max-height:420px;border:1px solid #c8d6ea"></div>')
    if imgs:
        body += "<h2>صور الفواتير</h2>" + "".join(imgs)
    sub = f"الفترة: من {fmt_iso_date(start)} إلى {fmt_iso_date(end)}   |   النوع: {'الكل' if cat == 'all' else PURCH_CATS[cat]}"
    return html_page("تقرير المشتريات", sub, body)


def build_report_html(kind, start, end):
    S = report_sections(kind, start, end)
    body = ""
    if kind == "all":
        body += html_table("ملخص التقرير", ["القسم", "عدد السجلات"],
                           [(x["name"], len(x["rows"])) for x in S if x["name"] not in SUMMARY_ONLY], False)
    body += "".join(html_table(x["title"], x["headers"], x["rows"]) for x in S)
    sub = S[0]["period"] if S else ""
    title = "التقرير الشامل" if kind == "all" else (S[0]["title"] if S else "تقرير")
    return html_page(title, sub, body)


# ============================ المقارنة بين فترتين ============================
POLARITY = {"الأخطاء": -1, "التنبيهات الشفهية": -1, "الإشادات": 1, "أيام الغياب": -1, "مرات التأخير": -1}   # الزيادة سيئة (-1) أو جيدة (+1)


def metrics(start, end):
    with db() as con:
        def one(sql, *a):
            return con.execute(sql, a).fetchone()[0]
        d = "date(created_at) BETWEEN ? AND ?"
        return {
            "الزيارات": one(f"SELECT COUNT(*) FROM visits WHERE {d}", start, end),
            "أعمال الصيانة": one(f"SELECT COUNT(*) FROM maintenance WHERE {d}", start, end),
            "الملاحظات": one(f"SELECT COUNT(*) FROM employee_notes WHERE kind='ملاحظة' AND {d}", start, end),
            "الأخطاء": one(f"SELECT COUNT(*) FROM employee_notes WHERE kind='خطأ' AND {d}", start, end),
            "التنبيهات الشفهية": one(f"SELECT COUNT(*) FROM employee_notes WHERE kind='تنبيه شفهي' AND {d}", start, end),
            "الإشادات": one(f"SELECT COUNT(*) FROM employee_notes WHERE kind='إشادة' AND {d}", start, end),
            "إذونات الخروج": one("SELECT COUNT(*) FROM permits WHERE permit_date BETWEEN ? AND ?", start, end),
            "أيام الغياب": one("SELECT COUNT(*) FROM absences WHERE kind LIKE 'غياب%' AND absence_date BETWEEN ? AND ?", start, end),
            "مرات التأخير": one("SELECT COUNT(*) FROM absences WHERE kind='تأخير' AND absence_date BETWEEN ? AND ?", start, end),
            "عمليات الجرد": one("SELECT COUNT(*) FROM (SELECT DISTINCT category, week_from, week_to FROM inventory "
                               "WHERE week_to>=? AND week_from<=?)", start, end),
        }


def emp_stats(start, end):
    with db() as con:
        rows = con.execute("""SELECT n.employee, COALESCE(e.department,'غير مسجل'),
                SUM(n.kind='خطأ'), SUM(n.kind='ملاحظة'), SUM(n.kind='تنبيه شفهي'), SUM(n.kind='إشادة')
            FROM employee_notes n LEFT JOIN employees e ON e.name = n.employee
            WHERE date(n.created_at) BETWEEN ? AND ? GROUP BY n.employee""", (start, end)).fetchall()
    return {r[0]: {"dept": r[1], "err": r[2], "note": r[3], "warn": r[4], "praise": r[5]} for r in rows}


_ZERO = {"dept": "", "err": 0, "note": 0, "warn": 0, "praise": 0}


def make_insights(rows, emps):
    out, cand = [], []
    for name, a, b, d, pct in rows:
        if d:
            cand.append((abs(d) * (2 if name in POLARITY else 1), name, a, b, d, pct))
    for _, name, a, b, d, pct in sorted(cand, reverse=True)[:3]:
        pol = POLARITY.get(name, 0)
        icon = "✅" if pol * d > 0 else "⚠️" if pol * d < 0 else "📊"
        ptxt = f" ({'▲' if d > 0 else '▼'} {abs(pct)}%)" if pct is not None else " (جديد - لم يكن في الفترة السابقة)"
        out.append(f"{icon} {name}: {a} مقابل {b}{ptxt}")
    worst = max(emps, key=lambda e: e["err_a"], default=None)
    if worst and worst["err_a"] > 0:
        out.append(f"👤 أكثر العاملين أخطاءً في الفترة الحالية: {worst['name']} ({worst['dept']}) - {worst['err_a']}")
    best = max(emps, key=lambda e: e["praise_a"], default=None)
    if best and best["praise_a"] > 0:
        out.append(f"🌟 الأكثر إشادةً: {best['name']} - {best['praise_a']}")
    dept = {}
    for e in emps:
        dept[e["dept"]] = dept.get(e["dept"], 0) + e["err_a"]
    if dept:
        dn, dv = max(dept.items(), key=lambda kv: kv[1])
        if dv > 0:
            out.append(f"🏷️ أكثر قسم أخطاءً: {dn} ({dv})")
    return out or ["لا توجد فروقات بين الفترتين أو لا توجد بيانات كافية للمقارنة."]


def compare_periods(a1, a2, b1, b2):
    """أ = الفترة الحالية، ب = الفترة السابقة (تواريخ ISO)."""
    ma, mb = metrics(a1, a2), metrics(b1, b2)
    rows = []
    for k in ma:
        d = ma[k] - mb[k]
        rows.append((k, ma[k], mb[k], d, None if mb[k] == 0 else round(d * 100 / mb[k])))
    ea, eb = emp_stats(a1, a2), emp_stats(b1, b2)
    emps = []
    for name in set(ea) | set(eb):
        x, y = ea.get(name, _ZERO), eb.get(name, _ZERO)
        emps.append({"name": name, "dept": x["dept"] or y["dept"], "err_a": x["err"], "err_b": y["err"],
                     "note_a": x["note"], "warn_a": x["warn"], "praise_a": x["praise"], "praise_b": y["praise"]})
    emps.sort(key=lambda e: (-e["err_a"], -(e["err_a"] - e["err_b"]), e["name"]))
    return {"a": (a1, a2), "b": (b1, b2), "rows": rows, "emps": emps, "insights": make_insights(rows, emps)}


def _add_months(d, m):
    y, mo = divmod(d.month - 1 + m, 12)
    y, mo = d.year + y, mo + 1
    last = (date(y + (mo == 12), mo % 12 + 1, 1) - timedelta(days=1)).day
    return date(y, mo, min(d.day, last))


def preset_periods(kind, today=None):
    """يرجع (a1,a2,b1,b2) لفترتين متساويتين في عدد الأيام."""
    t = today or date.today()
    if kind == "month":
        a1 = t.replace(day=1); b1 = _add_months(a1, -1)
        b2 = min(b1 + timedelta(days=(t - a1).days), a1 - timedelta(days=1))
    elif kind == "week":                                  # الأسبوع يبدأ السبت
        a1 = t - timedelta(days=(t.weekday() - 5) % 7); b1 = a1 - timedelta(days=7)
        b2 = b1 + (t - a1)
    else:
        a1 = t.replace(month=1, day=1); b1 = a1.replace(year=a1.year - 1)
        b2 = _add_months(t, -12)
    return a1.isoformat(), t.isoformat(), b1.isoformat(), b2.isoformat()


def _cmp_period_text(data):
    return (f"الفترة الحالية: {fmt_iso_date(data['a'][0])} ← {fmt_iso_date(data['a'][1])}   |   "
            f"الفترة السابقة: {fmt_iso_date(data['b'][0])} ← {fmt_iso_date(data['b'][1])}")


def _sign_txt(d, pct):
    return ("▲ " if d > 0 else "▼ " if d < 0 else "■ ") + str(abs(d)), ("-" if pct is None else f"{pct:+d}%")


def build_compare_excel(path, data):
    from openpyxl.chart import BarChart, Reference
    wb = Workbook(); sh = wb.active; sh.title = "المقارنة"
    period = _cmp_period_text(data)
    rows = [(n, a, b, d, "جديد" if pct is None else pct / 100) for n, a, b, d, pct in data["rows"]]
    _style_sheet(sh, "مقارنة الفترات", period, ["البند", "الفترة الحالية", "الفترة السابقة", "الفرق", "النسبة"],
                 [30, 16, 16, 12, 12])
    _write_rows(sh, rows, 5)
    for i, (n, a, b, d, pct) in enumerate(data["rows"], 5):
        sh.cell(i, 5).number_format = "0%"
        pol = POLARITY.get(n, 0)
        if d and pol:
            sh.cell(i, 4).font = Font(name="Arial", size=10, bold=True,
                                      color="12B886" if pol * d > 0 else "E5484D")
    ch = BarChart(); ch.type = "col"; ch.title = "مقارنة الفترتين"; ch.height = 9; ch.width = 20
    ch.add_data(Reference(sh, min_col=2, max_col=3, min_row=4, max_row=4 + len(rows)), titles_from_data=True)
    ch.set_categories(Reference(sh, min_col=1, min_row=5, max_row=4 + len(rows)))
    sh.add_chart(ch, "G4")
    s2 = wb.create_sheet("ملخص ذكي")
    _style_sheet(s2, "ملخص ذكي للمقارنة", period, ["الملاحظة"], [110])
    _write_rows(s2, [(x,) for x in data["insights"]], 1)
    s3 = wb.create_sheet("العاملون")
    _style_sheet(s3, "مقارنة العاملين", period,
                 ["اسم العامل", "القسم", "أخطاء حالية", "أخطاء سابقة", "الفرق", "إشادات حالية", "إشادات سابقة"],
                 [28, 12, 14, 14, 10, 14, 14])
    _write_rows(s3, [(e["name"], e["dept"], e["err_a"], e["err_b"], e["err_a"] - e["err_b"],
                      e["praise_a"], e["praise_b"]) for e in data["emps"]], 7)
    wb.save(path)


def build_compare_html(data):
    bars = ""
    mx = max([1] + [r[1] for r in data["rows"]] + [r[2] for r in data["rows"]])
    for n, a, b, d, pct in data["rows"]:
        pol = POLARITY.get(n, 0)
        cls = "" if not d or not pol else ("up" if pol * d > 0 else "dn")
        t1, t2 = _sign_txt(d, pct)
        bars += (f'<div class="br"><div class="n">{_e(n)}</div><div class="t">'
                 f'<div class="bar a" style="width:{max(1, a * 100 // mx)}%">{a}</div>'
                 f'<div class="bar b" style="width:{max(1, b * 100 // mx)}%">{b}</div></div>'
                 f'<div class="{cls}" style="width:120px"><bdi dir="ltr">{t1} ({t2})</bdi></div></div>')
    body = "<h2>الملخص الذكي</h2>" + "".join(f'<div class="ins">{_e(x)}</div>' for x in data["insights"])
    body += '<h2>الرسم البياني <small>(أزرق = الفترة الحالية، رمادي = السابقة)</small></h2><div class="bars">' + bars + "</div>"
    body += html_table("جدول المقارنة", ["البند", "الحالية", "السابقة", "الفرق", "النسبة"],
                       [(n, a, b, _sign_txt(d, pct)[0], _Raw(f'<bdi dir="ltr">{_sign_txt(d, pct)[1]}</bdi>'))
                        for n, a, b, d, pct in data["rows"]], False)
    body += html_table("مقارنة العاملين", ["اسم العامل", "القسم", "أخطاء حالية", "أخطاء سابقة", "الفرق", "إشادات حالية"],
                       [(e["name"], e["dept"], e["err_a"], e["err_b"], e["err_a"] - e["err_b"], e["praise_a"])
                        for e in data["emps"]])
    return html_page("مقارنة الفترات", _cmp_period_text(data), body)


# ============================ الواجهة ============================
def draw_bars(cv, labels, a_vals, b_vals):
    """رسم أعمدة مزدوج: الفترة الحالية (أزرق) مقابل السابقة (برتقالي)."""
    cv.delete("all")
    w, h = max(cv.winfo_width(), 360), max(cv.winfo_height(), 240)
    top, bottom, right, left = 40, 58, 14, 14
    n = max(len(labels), 1)
    area_w, area_h = w - left - right, h - top - bottom
    mx = max([1] + list(a_vals) + list(b_vals))
    for i in range(5):
        y = top + area_h * i / 4
        cv.create_line(left, y, w - right, y, fill=LINE)
        cv.create_text(left + 2, y - 1, text=str(round(mx * (4 - i) / 4)), anchor="sw", fill=MUTED, font=(FONT, 7))
    gw = area_w / n
    bw = min(28, gw / 2.6)
    for i, (lab, a, b) in enumerate(zip(labels, a_vals, b_vals)):
        cx = (w - right) - (gw * i + gw / 2)                  # الاتجاه من اليمين لليسار
        for off, val, col in ((-bw - 1, a, BLUE), (1, b, ORANGE)):
            x0 = cx + off
            bh = area_h * val / mx
            cv.create_rectangle(x0, top + area_h - bh, x0 + bw, top + area_h, fill=col, outline="")
            cv.create_text(x0 + bw / 2, top + area_h - bh - 2, text=str(val), anchor="s", fill=TEXT, font=(FONT, 8, "bold"))
        cv.create_text(cx, top + area_h + 6, text=lab, anchor="n", width=max(gw - 4, 30), fill=TEXT,
                       font=(FONT, 8), justify="center")
    cv.create_rectangle(w - 120, 10, w - 108, 22, fill=BLUE, outline="")
    cv.create_text(w - 126, 16, text="الفترة الحالية", anchor="e", fill=TEXT, font=(FONT, 9, "bold"))
    cv.create_rectangle(w - 250, 10, w - 238, 22, fill=ORANGE, outline="")
    cv.create_text(w - 256, 16, text="الفترة السابقة", anchor="e", fill=TEXT, font=(FONT, 9, "bold"))


def login_dialog(root):
    """نافذة تسجيل الدخول. ترجع بيانات المستخدم أو None لو أغلقها.

    لو الثيم الحالي بيدعم الهوية الفاخرة (LOGIN_G1/LOGIN_G2) بنرسم خلفية متدرّجة
    وصندوق داكن بإطار ذهبي وزرار دخول متدرّج — زي مواصفات التصميم.
    """
    win = tk.Toplevel(root)
    win.title(f"{APP_NAME} - تسجيل الدخول"); win.configure(bg=BG); win.resizable(False, False)
    w, h = 440, 560
    try:
        h = min(h, max(360, int(win.winfo_screenheight()) - 80))
        w = min(w, max(340, int(win.winfo_screenwidth()) - 40))
    except Exception:
        pass
    win.geometry(f"{w}x{h}+{max(0, (win.winfo_screenwidth() - w) // 2)}+{max(0, (win.winfo_screenheight() - h) // 2 - 20)}")
    res = {"user": None, "tries": 0, "retry": False}

    # ---- الهوية الفاخرة: خلفية متدرّجة + صندوق داكن (تقريب الزجاج) + زرار متدرّج ----
    G1, G2 = globals().get("LOGIN_G1"), globals().get("LOGIN_G2")
    fancy = bool(G1 and G2)
    PBG = (globals().get("LOGIN_BOX") or WHITE) if fancy else BG
    if fancy:
        bgcv = tk.Canvas(win, highlightthickness=0, bd=0, bg=G1)
        bgcv.place(x=0, y=0, relwidth=1, relheight=1)
        bgcv.tk.call("lower", bgcv._w)          # خلفية تحت كل العناصر

        def _paint_bg(_=None):
            try:
                paint_gradient(bgcv, max(1, win.winfo_width()), max(1, win.winfo_height()),
                               G1, G2, horizontal=False, step=3)
            except Exception:
                pass
        win.bind("<Configure>", _paint_bg, add="+")
        win.after(30, _paint_bg)
        shell = tk.Frame(win, bg=PBG, highlightbackground=GOLD, highlightthickness=1)
        shell.pack(fill="both", expand=True, padx=36, pady=28)
    else:
        shell = win

    def switch_lang():
        i18n.set_lang("ar" if i18n.is_en() else "en")
        cfg = load_settings(); cfg["lang"] = i18n.LANG; save_settings(cfg)
        res["retry"] = True
        win.destroy()

    lang_btn = tk.Button(win, text="🌐  العربية" if i18n.is_en() else "🌐  English", command=switch_lang,
                         bg=PBG, fg=PRIMARY, activebackground=SOFT, relief="flat", bd=0,
                         font=(FONT, 9, "bold"), padx=8, pady=3, cursor="hand2")
    lang_btn.place(relx=0.04 if i18n.is_en() else 0.96, y=12, anchor="nw" if i18n.is_en() else "ne")
    try:
        win._logo = tk.PhotoImage(data="".join(LOGO_B64))
        tk.Label(shell, image=win._logo, bg="white", highlightbackground=GOLD, highlightthickness=3).pack(pady=(18, 8))
    except Exception:
        pass
    tk.Label(shell, text=APP_NAME, bg=PBG, fg=INK, font=(FONT, 20, "bold")).pack()
    tk.Label(shell, text=f"نظام إدارة المركز - فرع {BRANCH}", bg=PBG, fg=MUTED, font=(FONT, 10)).pack(pady=(0, 14))
    if netmod and NET_MODE == netmod.MODE_CLIENT:
        tk.Label(shell, text=f"🖧  الجهاز الرئيسي: {NET_CFG.get('server_host') or '؟'}   •   المحطة: {NET_STATION}",
                 bg=PBG, fg=PRIMARY, font=(FONT, 9, "bold")).pack(pady=(0, 6))
    box = tk.Frame(shell, bg=PBG, highlightbackground=GOLD if fancy else LINE, highlightthickness=1)
    box.pack(fill="x", padx=30, pady=4)
    tk.Label(box, text="اسم المستخدم", bg=PBG, fg=TEXT, font=(FONT, 10, "bold"), anchor="e").pack(fill="x", padx=16, pady=(14, 3))
    u_var, p_var = tk.StringVar(), tk.StringVar()
    ue = tk.Entry(box, textvariable=u_var, justify="right", font=(FONT, 12), relief="solid", bd=1,
                  bg=ENTRY, fg=TEXT, insertbackground=TEXT); ue.pack(fill="x", padx=16)
    tk.Label(box, text="كلمة المرور", bg=PBG, fg=TEXT, font=(FONT, 10, "bold"), anchor="e").pack(fill="x", padx=16, pady=(10, 3))
    pe = tk.Entry(box, textvariable=p_var, show="●", justify="right", font=(FONT, 12), relief="solid", bd=1,
                  bg=ENTRY, fg=TEXT, insertbackground=TEXT); pe.pack(fill="x", padx=16, pady=(0, 14))
    msg = tk.Label(shell, text="", bg=PBG, fg=RED, font=(FONT, 10, "bold")); msg.pack(pady=(8, 0))

    def attempt(_=None):
        if not u_var.get().strip() or not p_var.get():
            msg.config(text="اكتب اسم المستخدم وكلمة المرور."); return
        user, why = authenticate(u_var.get(), p_var.get())
        if user:
            res["user"] = user; win.destroy(); return
        if why == "offline":
            messagebox.showerror("الشبكة", "تعذر الاتصال بالجهاز الرئيسي.\nاتأكد أن الجهاز الرئيسي مفتوح والبرنامج شغال عليه، وأن الشبكة متصلة.", parent=win)
            return
        res["tries"] += 1
        if res["tries"] >= 5:
            messagebox.showerror("تسجيل الدخول", "تم إدخال بيانات خاطئة 5 مرات. سيُغلق البرنامج.", parent=win)
            win.destroy(); return
        msg.config(text="هذا الحساب موقوف. راجع المدير." if why == "disabled"
                   else f"اسم المستخدم أو كلمة المرور غير صحيحة. (المحاولة {res['tries']} من 5)")
        p_var.set(""); pe.focus_set()
    gstops = globals().get("LOGIN_BTN")
    if fancy and gstops:
        gradient_button(shell, "🔓  دخول", attempt, gstops, width=300, height=48).pack(pady=14)
    else:
        tk.Button(shell, text="🔓  دخول", command=attempt, bg=PRIMARY, fg="white", activebackground=BLUE, activeforeground="white",
                  relief="flat", font=(FONT, 12, "bold"), padx=30, pady=8, cursor="hand2").pack(pady=14)
    tk.Label(shell, text=f"تصميم وتطوير: {DESIGNER}   •   الإصدار {APP_VERSION}", bg=PBG, fg=MUTED, font=(FONT, 8)).pack(side="bottom", pady=8)
    ue.bind("<Return>", lambda e: pe.focus_set()); pe.bind("<Return>", attempt)
    win.protocol("WM_DELETE_WINDOW", win.destroy)
    win.lift(); win.focus_force(); ue.focus_set()
    try:
        win.grab_set()
    except Exception:
        pass
    root.wait_window(win)
    return "RETRY" if res["retry"] else res["user"]


# ============================ الأيقونات من فولدر img ============================
# ضع ملفات PNG في فولدر img بجوار البرنامج. الاسم: <المفتاح>.png (أو بمقاس جاهز <المفتاح>@30.png).
# المفاتيح: attend gov cleaning paper toner purchases employees maintenance visits
#          visits_today maint_month notes_month  (وأي مفتاح قسم آخر مثل workers / govstaff / compare / reports / home)
# أي أيقونة غير موجودة يظهر مكانها الرمز (الإيموجي) القديم كما هو.
IMG_DIR = os.path.join(BASE, "img")
# نسخة exe (onefile): لو مجلد img غير موجود بجوار الملف نستخدم النسخة المضمّنة داخله
if getattr(sys, "frozen", False) and not os.path.isdir(IMG_DIR):
    _MEI = getattr(sys, "_MEIPASS", "")
    if _MEI and os.path.isdir(os.path.join(_MEI, "img")):
        IMG_DIR = os.path.join(_MEI, "img")
_ICON_CACHE = {}


def load_icon(key, size=24):
    ck = (key, size)
    if ck in _ICON_CACHE:
        return _ICON_CACHE[ck]
    img = None
    for name in (f"{key}@{size}.png", f"{key}.png", f"{key}.gif"):
        p = os.path.join(IMG_DIR, name)
        if os.path.isfile(p):
            try:
                img = tk.PhotoImage(file=p)
                if f"@{size}" not in name:
                    big = max(img.width(), img.height())
                    if big > size:
                        f = -(-big // size)
                        img = img.subsample(f, f)
                    elif big * 2 <= size:
                        img = img.zoom(size // big)
                break
            except Exception:
                img = None
    _ICON_CACHE[ck] = img
    return img


class CenterApp:
    NAV = [
        ("home", "🏠", "الرئيسية"),
        ("notes", "📢", "ملاحظات"),
        ("workers", "👷", "قاعدة بيانات العاملين"),
        ("govstaff", "🗂️", "قاعدة بيانات موظفي الجهات"),
        ("attend", "🕒", "حضور وانصراف العاملين"),
        ("gov", "🏛️", "حضور وانصراف الجهات"),
        ("visits", "👥", "زيارات المركز"),
        ("maintenance", "🔧", "الصيانة"),
        ("cleaning", "🧴", "جرد خامات النظافة"),
        ("paper", "📄", "جرد الأدوات الورقية"),
        ("toner", "🖨️", "جرد الأحبار"),
        ("purchases", "🛒", "المشتريات"),
        ("employees", "📝", "ملاحظات وأخطاء العاملين"),
        ("compare", "📈", "مقارنة الفترات"),
        ("reports", "📊", "استخراج وطباعة التقارير"),
    ]

    def __init__(self, root, user=None):
        self.root = root
        self.user = user or {"id": 0, "username": "admin", "is_admin": True, "perms": {}}
        self.root.title(T(f"{APP_NAME} - نظام إدارة المركز - {BRANCH}"))
        self.fit_to_screen()
        self.root.configure(bg=BG)
        self.nav_buttons = {}
        self._apply_scaling()
        self.setup_style()
        self.build_shell()
        self.bind_keys()
        self.go("home")

    def fit_to_screen(self):
        """يضبط حجم النافذة على مقاس الشاشة الحقيقية، فتبان الواجهة كاملة على أي شاشة صغيرة أو كبيرة."""
        try:
            sw = int(self.root.winfo_screenwidth() or 1240)
            sh = int(self.root.winfo_screenheight() or 820)
        except Exception:
            sw, sh = 1240, 820
        # هامش بسيط لشريط المهام وإطار النافذة، مع حد أدنى معقول
        w = min(1280, max(760, sw - 24))
        h = min(880, max(520, sh - 90))
        if sw < w: w = sw
        if sh < h: h = sh
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 3)
        try:
            self.root.geometry(f"{w}x{h}+{x}+{y}")
            self.root.minsize(min(720, sw), min(470, sh))
        except Exception:
            pass
        self._screen_size = (sw, sh)

    def _clamp_size(self, size, min_w=340, min_h=260):
        """يصغّر أي مقاس نافذة (مثل 940x780) لو أكبر من الشاشة."""
        try:
            sw, sh = getattr(self, "_screen_size", (0, 0))
            if not sw or not sh:
                sw, sh = int(self.root.winfo_screenwidth()), int(self.root.winfo_screenheight())
            w, h = (int(v) for v in str(size).lower().split("x")[:2])
            w = min(w, max(min_w, sw - 30))
            h = min(h, max(min_h, sh - 110))
            return f"{w}x{h}"
        except Exception:
            return size

    def _minsize(self, win, w, h):
        """حد أدنى لمقاس النافذة، مع عدم تجاوز مقاس الشاشة."""
        try:
            sw, sh = getattr(self, "_screen_size", (0, 0))
            if not sw or not sh:
                sw, sh = int(self.root.winfo_screenwidth()), int(self.root.winfo_screenheight())
            win.minsize(min(int(w), max(320, sw - 30)), min(int(h), max(240, sh - 110)))
        except Exception:
            try:
                win.minsize(w, h)
            except Exception:
                pass

    # ---------- الصلاحيات ----------
    def level(self, page):
        if self.user["is_admin"]:
            return "edit"
        lv = self.user["perms"].get(page, "none")
        if page == "notes" and lv == "none":
            lv = "view"          # الملاحظات تعليمات موجهة للوحدات: متاحة للعرض لكل وحدة
        if page in VIEW_ONLY_PAGES and lv in ("add", "edit"):
            lv = "view"
        return lv if lv in LEVELS else "none"

    def can(self, page, need="view"):
        return LEVELS.index(self.level(page)) >= LEVELS.index(need)

    def need(self, page, need="edit"):
        """يتحقق من الصلاحية ويعرض رسالة لو غير مسموح."""
        if self.can(page, need):
            return True
        what = {"add": "الإضافة", "edit": "التعديل أو الحذف", "view": "الدخول"}[need]
        messagebox.showwarning("غير مسموح", f"ليس لديك صلاحية {what} في هذا القسم.\n"
                               f"صلاحيتك الحالية: {LEVEL_LABEL[self.level(page)]}.\nراجع المدير لو محتاج صلاحية أعلى.")
        return False

    def need_admin(self):
        if self.user["is_admin"]:
            return True
        messagebox.showwarning("غير مسموح", "هذه العملية للمدير فقط.")
        return False

    def all_data_visible(self):
        return self.user["is_admin"] or all(self.can(k, "view") for k, _ in PERM_PAGES if k not in VIEW_ONLY_PAGES)

    # ---------- الشكل العام ----------
    def setup_style(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("Treeview", rowheight=ROW_H, font=(FONT, 10), background=WHITE,
                     fieldbackground=WHITE, foreground=TEXT, borderwidth=0)
        st.configure("Treeview.Heading", font=(FONT, 10, "bold"), background=NAVY_2,
                     foreground="white", padding=8, relief="flat")
        st.map("Treeview.Heading", background=[("active", BLUE)])
        st.map("Treeview", background=[("selected", SEL)], foreground=[("selected", INK)])
        st.configure("TEntry", padding=7, fieldbackground=ENTRY, foreground=TEXT, insertcolor=TEXT,
                     bordercolor=LINE, lightcolor=LINE, darkcolor=LINE)
        st.configure("TCombobox", padding=6, fieldbackground=ENTRY, background=SOFT, foreground=TEXT,
                     arrowcolor=TEXT, bordercolor=LINE, lightcolor=LINE, darkcolor=LINE)
        st.map("TCombobox", fieldbackground=[("readonly", ENTRY)], foreground=[("readonly", TEXT)],
               selectbackground=[("readonly", ENTRY)], selectforeground=[("readonly", TEXT)])
        st.configure("TScrollbar", background=SOFT, troughcolor=BG, bordercolor=BG, arrowcolor=TEXT)
        st.configure("TRadiobutton", background=WHITE, foreground=TEXT, font=(FONT, 10))
        st.map("TRadiobutton", background=[("active", WHITE)])
        st.configure("TNotebook", background=BG, borderwidth=0)
        st.configure("TNotebook.Tab", padding=(18, 8), font=(FONT, 10, "bold"), background=SOFT, foreground=TEXT)
        st.map("TNotebook.Tab", background=[("selected", WHITE)], foreground=[("selected", INK)])
        self.root.option_add("*TCombobox*Listbox.background", ENTRY)
        self.root.option_add("*TCombobox*Listbox.foreground", TEXT)
        self.root.option_add("*TCombobox*Listbox.selectBackground", BLUE)
        self.root.option_add("*TCombobox*Listbox.selectForeground", "white")

    def load_logo(self):
        """يحمّل logo.png من جوار البرنامج إن وُجد، وإلا يستخدم اللوجو المدمج."""
        ext = os.path.join(BASE, "logo.png")
        try:
            if os.path.exists(ext):
                img = tk.PhotoImage(file=ext)
                f = max(1, -(-img.height() // 64))   # تصغير تلقائي لارتفاع 64 تقريبًا
                return img.subsample(f, f) if f > 1 else img
            return tk.PhotoImage(data="".join(LOGO_B64))
        except Exception:
            return None

    def build_shell(self):
        scale = max(1.0, FONT_SCALE)          # يكبر الهيدر والقائمة مع تكبير الخط
        H = int(86 * scale)
        hdr = tk.Canvas(self.root, height=H, highlightthickness=0, bd=0, bg=NAVY)
        hdr.pack(fill="x")
        self.hdr = hdr
        self.logo_img = self.load_logo()
        logo = None
        if self.logo_img:
            logo = tk.Label(hdr, image=self.logo_img, bg="white", bd=0, highlightbackground=GOLD, highlightthickness=3)
        else:
            logo = tk.Label(hdr, text="خ", bg=GOLD, fg="#1a2233", font=(FONT, 26, "bold"), width=3)
        en = i18n.is_en()
        logo_win = hdr.create_window(0, 0, anchor="w" if en else "e", window=logo)
        title_id = hdr.create_text(0, 0, anchor="center", text=APP_NAME, fill=HDR_TEXT, font=(FONT, 26, "bold"))
        sub_id = hdr.create_text(0, 0, anchor="center", text=f"فرع {BRANCH}", fill=HDR_SUB, font=(FONT, 13, "bold"))
        self.clock_id = hdr.create_text(0, 0, anchor="e" if en else "w", text="", fill=HDR_TEXT, font=(FONT, 11, "bold"))
        self.theme_btn = tk.Button(hdr, text="🌙  الوضع الليلي" if THEME not in DARK_THEMES else "☀️  الوضع النهاري",
                                   command=self.toggle_theme, bg=GOLD, fg="#1a2233", activebackground="white",
                                   relief="flat", bd=0, font=(FONT, 9, "bold"), padx=12, pady=4, cursor="hand2")
        btn_win = hdr.create_window(0, 0, anchor="e" if en else "w", window=self.theme_btn)
        # زر تبديل اللغة: يعرض اسم اللغة الأخرى (بدون ترجمة)
        self.lang_btn = tk.Button(hdr, text="🌐  العربية" if en else "🌐  English",
                                  command=self.toggle_lang, bg="white", fg="#1a2233", activebackground=GOLD,
                                  relief="flat", bd=0, font=(FONT, 9, "bold"), padx=12, pady=4, cursor="hand2")
        lang_win = hdr.create_window(0, 0, anchor="e" if en else "w", window=self.lang_btn)
        who = self.user["username"] + (" (مدير)" if self.user["is_admin"] else "")
        self.user_btn = tk.Menubutton(hdr, text=f"👤  {who}  ▾", bg="white", fg=NAVY, activebackground=GOLD,
                                      relief="flat", bd=0, font=(FONT, 9, "bold"), padx=12, pady=4, cursor="hand2")
        um = tk.Menu(self.user_btn, tearoff=0, font=(FONT, 10)); self.user_btn.config(menu=um)
        um.add_command(label="🔑  تغيير كلمة المرور", command=self.change_password_dialog)
        um.add_command(label="🎨  المظهر والثيمات", command=self.appearance_dialog)
        um.add_command(label="⛶  ملء الشاشة", command=self.toggle_fullscreen)
        um.add_command(label="⌨️  اختصارات لوحة المفاتيح", command=self.shortcuts_dialog)
        if self.user["is_admin"]:
            um.add_command(label="👥  المستخدمون والصلاحيات", command=self.users_dialog)
            um.add_command(label="🌐  الشبكة (LAN)", command=self.net_dialog)
            um.add_command(label="📋  سجل العمليات", command=self.audit_dialog)
            um.add_command(label="🗂️  النسخ الاحتياطية", command=self.backups_dialog)
        um.add_separator()
        um.add_command(label="🚪  تسجيل الخروج", command=self.logout)
        umb_win = hdr.create_window(0, 0, anchor="e" if en else "w", window=self.user_btn)

        def layout(_=None):
            w = max(hdr.winfo_width(), 400)
            paint_gradient(hdr, w, H, H1, H2)
            hdr.create_rectangle(0, H - 4, w, H, fill=GOLD, outline="", tags="grad")
            hdr.tag_lower("grad")
            # عنوان أصغر على الشاشات الضيقة عشان مايتداخلش مع الأزرار
            narrow = w < 1060
            tiny = w < 900
            dark = THEME in DARK_THEMES
            t_txt = ("🌙" if dark else "☀️") if tiny else ("🌙  الوضع الليلي" if dark else "☀️  الوضع النهاري")
            l_txt = "🌐" if tiny else ("🌐  العربية" if en else "🌐  English")
            u_txt = "👤  ▾" if tiny else f"👤  {who}  ▾"
            self.theme_btn.config(text=t_txt)
            self.lang_btn.config(text=l_txt)
            self.user_btn.config(text=u_txt)
            hdr.itemconfig(title_id, font=(FONT, 20 if narrow else 26, "bold"))
            hdr.itemconfig(sub_id, font=(FONT, 11 if narrow else 13, "bold"))
            f9 = tkfont.Font(font=(FONT, 9, "bold"))
            gap = max(8, int(10 * scale))
            row1, row2 = int(28 * scale), int(58 * scale)
            logo_w = max(logo.winfo_reqwidth(), 40)
            if en:                          # English interface: logo left, controls right (LTR)
                hdr.coords(logo_win, 22, H // 2 - 2)
                hdr.coords(self.clock_id, w - 24, row1)
                x = w - 20
                for txt, wid in ((t_txt, btn_win), (l_txt, lang_win), (u_txt, umb_win)):
                    hdr.coords(wid, x, row2)
                    x -= max(f9.measure(txt) + 30, 42) + gap
                lo, hi = 22 + logo_w + 14, x + gap - 12
            else:                           # Arabic interface (RTL)
                hdr.coords(logo_win, w - 22, H // 2 - 2)
                hdr.coords(self.clock_id, 24, row1)
                x = 20
                for txt, wid in ((t_txt, btn_win), (l_txt, lang_win), (u_txt, umb_win)):
                    hdr.coords(wid, x, row2)
                    x += max(f9.measure(txt) + 30, 42) + gap
                lo, hi = x - gap + 12, (w - 22) - logo_w - 14
            cx = w // 2 if (w >= 1100 or hi <= lo + 40) else (lo + hi) // 2
            hdr.coords(title_id, cx, int(33 * scale)); hdr.coords(sub_id, cx, int(63 * scale))
        hdr.bind("<Configure>", layout)
        self.hdr = hdr
        self.title_id = title_id
        self.sub_id = sub_id
        self.tick()

        footer = tk.Frame(self.root, bg=NAVY, height=30)
        footer.pack(side="bottom", fill="x")
        footer.pack_propagate(False)
        self.footer_lbl = tk.Label(footer, text=self.footer_text(),
                                   bg=NAVY, fg=GOLD, font=(FONT, 9, "bold"))
        self.footer_lbl.pack(expand=True)

        shell = tk.Frame(self.root, bg=BG)
        shell.pack(fill="both", expand=True)
        self.sidebar = tk.Frame(shell, bg=NAVY_2, width=self._side_width())
        self.sidebar.pack(side="right", fill="y")
        self.sidebar.pack_propagate(False)
        tk.Label(self.sidebar, text="القائمة الرئيسية", bg=NAVY_2, fg=SIDE_MUTED,
                 font=(FONT, 11, "bold")).pack(pady=(16, 8))
        # الأزرار السفلية تُحجز قبل منطقة التمرير عشان تفضل مثبّتة في آخر القائمة
        tk.Button(self.sidebar, text="المظهر والثيمات   🎨", command=self.appearance_dialog, anchor="e",
                  bg=NAVY_2, fg=SIDE_ACCENT, activebackground=BLUE, activeforeground="white", relief="flat", bd=0,
                  font=(FONT, 10, "bold"), padx=18, pady=8, cursor="hand2").pack(side="bottom", fill="x", padx=10, pady=(0, 6))
        if self.user["is_admin"]:
            tk.Button(self.sidebar, text="النسخ الاحتياطية   🗂️", command=self.backups_dialog, anchor="e",
                      bg=NAVY, fg=SIDE_ACCENT, activebackground=BLUE, activeforeground="white", relief="flat", bd=0,
                      font=(FONT, 10, "bold"), padx=18, pady=9, cursor="hand2").pack(side="bottom", fill="x", padx=10, pady=(0, 14))
            tk.Button(self.sidebar, text="إدراج نسخة احتياطية   📂", command=self.restore_backup, anchor="e",
                      bg=NAVY, fg=SIDE_ACCENT2, activebackground=BLUE, activeforeground="white", relief="flat", bd=0,
                      font=(FONT, 10, "bold"), padx=18, pady=9, cursor="hand2").pack(side="bottom", fill="x", padx=10, pady=(0, 6))
        # منطقة أزرار القائمة جوه Canvas بمسطرة تمرير، عشان توصل لآخر القائمة على أي مقاس شاشة
        nav_wrap = tk.Frame(self.sidebar, bg=NAVY_2)
        nav_wrap.pack(fill="both", expand=True)
        nav_wrap.rowconfigure(0, weight=1)
        nav_wrap.columnconfigure(1, weight=1)
        self.nav_canvas = tk.Canvas(nav_wrap, bg=NAVY_2, highlightthickness=0, bd=0)
        self.nav_vsb = ttk.Scrollbar(nav_wrap, orient="vertical", command=self.nav_canvas.yview)
        self.nav_canvas.configure(yscrollcommand=self.nav_vsb.set)
        self.nav_vsb.grid(row=0, column=0, sticky="ns")
        self.nav_canvas.grid(row=0, column=1, sticky="nsew")
        self._nav_vsb_shown = True
        self.nav_inner = tk.Frame(self.nav_canvas, bg=NAVY_2)
        self.nav_inner_id = self.nav_canvas.create_window((0, 0), window=self.nav_inner, anchor="nw")
        for key, icon, label in self.NAV:
            if key != "home" and not self.can(key, "view"):
                continue
            ic = load_icon(key, 30)
            look = dict(text=" " + label, image=ic, compound="right") if ic else dict(text=f"{label}   {icon}")
            b = tk.Button(self.nav_inner, command=lambda k=key: self.go(k),
                          anchor="e", bg=NAVY_2, fg=SIDE_TEXT, activebackground=BLUE,
                          activeforeground="white", relief="flat", bd=0, takefocus=True,
                          font=(FONT, 11, "bold"), padx=18, pady=NAV_PADY if ic else NAV_PADY + 2, cursor="hand2", **look)
            b.pack(fill="x", padx=10, pady=1)
            b.bind("<Enter>", lambda e, k=key: self.nav_hover(k, True))
            b.bind("<Leave>", lambda e, k=key: self.nav_hover(k, False))
            # الأسهم تتحرّك في القائمة (RTL: يمين = لأعلى، شمال = لأسفل)
            b.bind("<Up>", lambda e: (self.nav_move(-1), "break")[1])
            b.bind("<Down>", lambda e: (self.nav_move(1), "break")[1])
            b.bind("<Right>", lambda e: (self.nav_move(-1), "break")[1])
            b.bind("<Left>", lambda e: (self.nav_move(1), "break")[1])
            self.nav_buttons[key] = b
        self.nav_inner.bind("<Configure>", lambda e: self._fit_nav())
        self.nav_canvas.bind("<Configure>", lambda e: self._fit_nav())
        self.active_key = None
        # منطقة المحتوى داخل Canvas بشرائط تمرير (رأسي + أفقي) عشان كل الصفحة تبان على أي مقاس شاشة
        self.content_wrap = tk.Frame(shell, bg=BG)
        self.content_wrap.pack(side="left", fill="both", expand=True)
        self.content_wrap.rowconfigure(0, weight=1)
        self.content_wrap.columnconfigure(1, weight=1)
        self.content_canvas = tk.Canvas(self.content_wrap, bg=BG, highlightthickness=0, bd=0)
        self.content_vsb = ttk.Scrollbar(self.content_wrap, orient="vertical", command=self.content_canvas.yview)
        self.content_hsb = ttk.Scrollbar(self.content_wrap, orient="horizontal", command=self.content_canvas.xview)
        self.content_canvas.configure(yscrollcommand=self.content_vsb.set, xscrollcommand=self.content_hsb.set)
        self.content_vsb.grid(row=0, column=0, sticky="ns")
        self.content_canvas.grid(row=0, column=1, sticky="nsew")
        self.content_hsb.grid(row=1, column=1, sticky="ew")
        self._vsb_shown = True
        self._hsb_shown = True
        self.content = tk.Frame(self.content_canvas, bg=BG)
        self.content_id = self.content_canvas.create_window((0, 0), window=self.content, anchor="nw")
        self._last_size = None
        self.content.bind("<Configure>", lambda e: self._fit_content())
        self.content_canvas.bind("<Configure>", lambda e: self._fit_content())
        self.root.bind("<Configure>", self._on_root_config, add="+")
        self._bind_wheel()
        self.root.after(80, self._fit_nav)

    def _side_width(self):
        """عرض القائمة الجانبية حسب عرض النافذة (تضيق تلقائيًا على الشاشات الصغيرة)."""
        scale = max(1.0, FONT_SCALE)
        try:
            w = self.root.winfo_width()
        except Exception:
            w = 0
        if w <= 1:
            w = 1240
        if w < 950:
            return int(205 * scale)
        if w < 1150:
            return int(240 * scale)
        return int(285 * scale)

    def _on_root_config(self, e=None):
        """عند تغيير حجم النافذة: يعدّل عرض القائمة الجانبية ويعيد ضبط التمرير."""
        if e is not None and getattr(e, "widget", None) is not self.root:
            return
        try:
            want = self._side_width()
            if getattr(self, "_side_now", None) != want:
                self._side_now = want
                self.sidebar.config(width=want)
        except Exception:
            pass
        self._fit_content()
        self._fit_nav()

    def _fit_content(self):
        """يخلي المحتوى يملأ المكان المتاح، ويضيف تمرير لما يكون أطول/أعرض من الشاشة."""
        try:
            c = self.content_canvas
            cw, ch = max(1, c.winfo_width()), max(1, c.winfo_height())
            rw, rh = max(1, self.content.winfo_reqwidth()), max(1, self.content.winfo_reqheight())
            w, h = max(cw, rw), max(ch, rh)
            c.configure(scrollregion=(0, 0, w, h))
            if self._last_size != (w, h):
                self._last_size = (w, h)
                c.itemconfig(self.content_id, width=w, height=h)
            self._show_scrollbars(rw > cw + 1, rh > ch + 1)
        except Exception:
            pass

    def _show_scrollbars(self, need_h, need_v):
        """يخفي شرائط التمرير لو مش محتاجينها (شكل أنضف)."""
        try:
            if bool(need_v) != self._vsb_shown:
                self._vsb_shown = bool(need_v)
                self.content_vsb.grid() if need_v else self.content_vsb.grid_remove()
            if bool(need_h) != self._hsb_shown:
                self._hsb_shown = bool(need_h)
                self.content_hsb.grid() if need_h else self.content_hsb.grid_remove()
        except Exception:
            pass

    def _bind_wheel(self):
        """تمرير المحتوى بعجلة الفأرة (ويندوز ولينكس) — بدون التأثير على النوافذ المنبثقة."""
        try:
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                self.root.unbind_all(seq)
                self.root.bind_all(seq, self._on_wheel)
        except Exception:
            pass

    def _on_wheel(self, e):
        try:
            w = self.root.winfo_containing(e.x_root, e.y_root)
            if w is None:
                return
            sw = str(w)
            num = getattr(e, "num", 0)
            if num == 4:
                step = -1
            elif num == 5:
                step = 1
            else:
                step = -1 if (getattr(e, "delta", 0) or 0) > 0 else 1
            if sw.startswith(str(self.content_canvas)):
                self.content_canvas.yview_scroll(step * 3, "units")
            elif getattr(self, "nav_canvas", None) is not None and sw.startswith(str(self.nav_canvas)):
                self.nav_canvas.yview_scroll(step * 3, "units")
        except Exception:
            pass

    def _fit_nav(self):
        """يضبط تمرير القائمة الجانبية ويخفي المسطرة لو القائمة كلها بانية أصلاً."""
        try:
            c = self.nav_canvas
            cw, ch = max(1, c.winfo_width()), max(1, c.winfo_height())
            rh = max(1, self.nav_inner.winfo_reqheight())
            c.configure(scrollregion=(0, 0, cw, max(ch, rh)))
            c.itemconfig(self.nav_inner_id, width=cw)
            need = rh > ch + 1
            if need != self._nav_vsb_shown:
                self._nav_vsb_shown = need
                self.nav_vsb.grid() if need else self.nav_vsb.grid_remove()
        except Exception:
            pass

    def _nav_keys(self):
        """ترتيب أزرار القائمة الجانبية الظاهرة للمستخدم الحالي."""
        return [k for k, _ic, _lb in self.NAV if k in self.nav_buttons]

    def nav_move(self, step):
        """ينقل التحديد في القائمة الجانبية بالأسهم، وينزّل القائمة لو الرابط تحت الشاشة."""
        keys = self._nav_keys()
        if not keys:
            return
        cur = self.active_key if self.active_key in keys else keys[0]
        i = keys.index(cur)
        j = max(0, min(len(keys) - 1, i + step))
        k = keys[j]
        if k != self.active_key:
            self.go(k)
        self._nav_reveal(k)

    def _nav_reveal(self, key):
        """ينزّل/يطلّع القائمة الجانبية عشان الزرار المحدد يبان بالكامل."""
        b = self.nav_buttons.get(key)
        if b is None:
            return
        try:
            c = self.nav_canvas
            c.update_idletasks()
            ch = max(1, c.winfo_height())
            rh = max(1, self.nav_inner.winfo_reqheight())
            if rh <= ch:
                return
            top, bot = b.winfo_y(), b.winfo_y() + b.winfo_height()
            vt = c.canvasy(0)
            vb = vt + ch
            if top < vt:
                c.yview_moveto(max(0.0, top / rh))
            elif bot > vb:
                c.yview_moveto(max(0.0, min(1.0, (bot - ch) / rh)))
        except Exception:
            pass

    def _menu_arrow(self, step):
        """الأسهم تتحرّك في القائمة الجانبية — من غير ما تعطّل الجداول والخانات النشطة."""
        try:
            w = self.root.focus_get()
            cls = w.winfo_class() if w is not None else ""
        except Exception:
            cls = ""
        if cls in ("Entry", "TEntry", "Text", "Treeview", "Listbox",
                   "TCombobox", "Spinbox", "TSpinbox"):
            return None
        self.nav_move(step)
        return "break"

    def rebuild_ui(self):
        """يعيد بناء الواجهة بعد تغيير الألوان أو الشكل أو اللغة مع البقاء في نفس الصفحة."""
        cur = self.active_key or "home"
        for w in self.root.winfo_children():
            w.destroy()
        self.nav_buttons = {}
        self.root.configure(bg=BG)
        self._apply_scaling()
        self.setup_style(); self.build_shell(); self.bind_keys(); self.go(cur)

    def _apply_scaling(self):
        """يطبّق حجم الخط العام عبر tk scaling (نسبة مبنية على القيمة الأصلية للشاشة)."""
        try:
            base = getattr(self, "_base_scaling", None)
            if not base or base <= 0:
                base = float(self.root.tk.call("tk", "scaling"))
                if base <= 0:
                    base = 1.333
                self._base_scaling = base
            self.root.tk.call("tk", "scaling", base * FONT_SCALE)
        except Exception:
            pass

    # اختصارات لوحة المفاتيح: F1..F11 -> الأقسام (F12 للمظهر)
    HOTKEYS = [("F1", "home"), ("F2", "notes"), ("F3", "workers"), ("F4", "govstaff"),
               ("F5", "attend"), ("F6", "gov"), ("F7", "visits"), ("F8", "maintenance"),
               ("F9", "purchases"), ("F10", "reports"), ("F11", "compare")]

    def bind_keys(self):
        """يربط اختصارات لوحة المفاتيح: F1..F12 للأقسام + Alt+الأسهم للتنقّل + أسهم/PageUp/Down للتمرير."""
        for seq, key in self.HOTKEYS:
            # مهم: لازم أقواس <> حوالين اسم الزرار، وإلا الربط مايشتغلش أصلاً
            self.root.bind(f"<{seq}>", lambda e, k=key: (self.go(k), "break")[1])
        self.root.bind("<F12>", lambda e: (self.appearance_dialog(), "break")[1])
        self.root.bind("<Shift-F1>", lambda e: (self.shortcuts_dialog(), "break")[1])
        self.root.bind("<Alt-Right>", lambda e: (self.cycle_page(-1), "break")[1])
        self.root.bind("<Alt-Left>", lambda e: (self.cycle_page(1), "break")[1])
        # التمرير بالكيبورد (لو مفيش جدول/خانة مفتوحة بتمتص المفاتيح)
        self.root.bind("<Prior>", lambda e: self._scroll_page(-1))      # PageUp
        self.root.bind("<Next>", lambda e: self._scroll_page(1))        # PageDown
        # الأسهم ↑↓ تتحرّك في القائمة الجانبية وتنزّلها لو أطول من الشاشة
        self.root.bind("<Up>", lambda e: self._menu_arrow(-1))
        self.root.bind("<Down>", lambda e: self._menu_arrow(1))
        self.root.bind("<Control-Home>", lambda e: self.content_canvas.yview_moveto(0))
        self.root.bind("<Control-End>", lambda e: self.content_canvas.yview_moveto(1))
        self.root.bind("<Alt-Home>", lambda e: self.go("home"))
        self.root.bind("<Control-f>", lambda e: (self.toggle_fullscreen(), "break")[1])

    def toggle_fullscreen(self):
        """يكبّر البرنامج على الشاشة كلها أو يرجّعه لحجمه الطبيعي، ويعيد ضبط الواجهة تلقائيًا.

        بعد التكبير/التصغير بنعيد ترتيب القائمة الجانبية والمحتوى على مقاس الشاشة الجديد،
        فتبان الواجهة كاملة على أي مقاس شاشة أو رزوليوشن.
        """
        self._fs = not getattr(self, "_fs", False)
        win = sys.platform.startswith("win")
        try:
            sw, sh = int(self.root.winfo_screenwidth()), int(self.root.winfo_screenheight())
        except Exception:
            sw = sh = 0
        if self._fs:
            ok = False
            if win:
                try:
                    self.root.attributes("-fullscreen", True)
                    ok = True
                except Exception:
                    ok = False
            if not ok and sw and sh:
                try:
                    self.root.geometry("%dx%d+0+0" % (sw, sh))
                except Exception:
                    pass
        else:
            if win:
                try:
                    self.root.attributes("-fullscreen", False)
                except Exception:
                    pass
            self.fit_to_screen()
        # إعادة ضبط الواجهة بعد ما النافذة تستقر (فورًا + بعد لحظة للتأكيد)
        self._relayout()
        try:
            self.root.after(120, self._relayout)
            self.root.after(360, self._relayout)
        except Exception:
            pass

    def _relayout(self):
        """يعيد ترتيب الواجهة (عرض القائمة الجانبية + المحتوى + تمرير القائمة) على المقاس الحالي."""
        try:
            self.root.update_idletasks()
        except Exception:
            pass
        try:
            want = self._side_width()
            self._side_now = want
            self.sidebar.config(width=want)
        except Exception:
            pass
        self._fit_content()
        self._fit_nav()

    def _scroll_page(self, direction):
        try:
            self.content_canvas.yview_scroll(direction, "pages")
        except Exception:
            pass

    def shortcuts_dialog(self):
        """نافذة بكل اختصارات لوحة المفاتيح."""
        win = self._dialog("اختصارات لوحة المفاتيح", "560x600")
        tk.Label(win, text="⌨️ اختصارات لوحة المفاتيح", bg=BG, fg=INK,
                 font=(FONT, 16, "bold")).pack(anchor="e", padx=20, pady=(14, 4))
        tk.Label(win, text="تقدر تدير البرنامج كله من الكيبورد من غير فأرة.", bg=BG, fg=MUTED,
                 font=(FONT, 10)).pack(anchor="e", padx=22, pady=(0, 8))
        rows = [("F1 … F11", "الانتقال السريع للأقسام (الرئيسية، ملاحظات، العاملين، …)"),
                ("F12", "فتح شاشة المظهر والثيمات"),
                ("Alt + →  /  Alt + ←", "التنقّل بين الأقسام واحد واحد"),
                ("Alt + Home", "الرجوع للصفحة الرئيسية"),
                ("Tab  /  Shift + Tab", "التنقّل بين الأزرار والخانات"),
                ("Space  /  Enter", "تشغيل الزرار اللي عليه التركيز"),
                ("↑  /  ↓", "التحرّك في القائمة الجانبية (وتنزّلها لو أطول من الشاشة)"),
                ("↑ ↓ ← →", "التحرّك جوه الجداول والقوائم"),
                ("PageUp / PageDown", "تمرير الصفحة لأعلى / لأسفل"),
                ("Ctrl + Home / Ctrl + End", "أول الصفحة / آخر الصفحة"),
                ("Ctrl + F", "ملء الشاشة / الرجوع لحجمها الطبيعي"),
                ("عجلة الفأرة", "تمرير المحتوى"),
                ("Esc", "إغلاق أي نافذة مفتوحة"),
                ("Shift + F1", "عرض الشاشة دي")]
        box = self.panel(win); box.pack(fill="both", expand=True, padx=16, pady=8)
        for i, (k, d) in enumerate(rows):
            r = tk.Frame(box, bg=WHITE); r.pack(fill="x", padx=12, pady=3)
            tk.Label(r, text=k, bg=WHITE, fg=PRIMARY, font=(FONT, 10, "bold"),
                     width=22, anchor="e").pack(side="right")
            tk.Label(r, text=d, bg=WHITE, fg=TEXT, font=(FONT, 10), anchor="e",
                     justify="right", wraplength=320).pack(side="right", padx=(0, 8))
        self.btn(win, "تمام", win.destroy, PRIMARY).pack(pady=10)

    def visible_pages(self):
        return [k for k, _ic, _lb in self.NAV if k == "home" or self.can(k, "view")]

    def cycle_page(self, step):
        pages = self.visible_pages()
        if not pages:
            return
        cur = self.active_key if self.active_key in pages else pages[0]
        self.go(pages[(pages.index(cur) + step) % len(pages)])

    def toggle_theme(self):
        """يبدّل بين الوضع النهاري وآخر ثيم ليلي استُخدم."""
        cfg = load_settings()
        if THEME in DARK_THEMES:
            new = cfg.get("last_light", "ocean")
            if new in DARK_THEMES:
                new = "ocean"
        else:
            cfg["last_light"] = THEME
            new = "midnight"
        apply_theme(new); cfg["theme"] = new; save_settings(cfg)
        self.rebuild_ui()

    def appearance_dialog(self):
        """شاشة المظهر: 8 ثيمات + نوع الخط + حجم الخط + شكل الأزرار + كثافة العرض.

        المحتوى جوه Canvas بمسطرتين (طولية + عرضية) عشان توصل لآخر الثيمات على أي مقاس شاشة.
        """
        win = self._dialog("المظهر والثيمات", "960x800")
        self._minsize(win, 420, 320)
        tk.Label(win, text="🎨 المظهر والثيمات", bg=BG, fg=INK, font=(FONT, 18, "bold")).pack(anchor="e", padx=24, pady=(14, 2))
        tk.Label(win, text="اختر ثيمًا كاملًا (ألوان + خط + شكل) — التغيير يُطبَّق ويُحفظ فورًا. استخدم المسطرتين للوصول لكل الثيمات.",
                 bg=BG, fg=MUTED, font=(FONT, 10)).pack(anchor="e", padx=26)

        # منطقة قابلة للتمرير (طولي + عرضي) عشان كل الثيمات تبان على أي مقاس شاشة
        wrap = tk.Frame(win, bg=BG)
        wrap.pack(fill="both", expand=True, padx=12, pady=6)
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(1, weight=1)
        canvas = tk.Canvas(wrap, bg=BG, highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(wrap, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.grid(row=0, column=0, sticky="ns")
        canvas.grid(row=0, column=1, sticky="nsew")
        hsb.grid(row=1, column=1, sticky="ew")
        body = tk.Frame(canvas, bg=BG)
        bid = canvas.create_window((0, 0), window=body, anchor="nw")

        def fit(_=None):
            try:
                cw, ch = max(1, canvas.winfo_width()), max(1, canvas.winfo_height())
                rw, rh = max(1, body.winfo_reqwidth()), max(1, body.winfo_reqheight())
                canvas.configure(scrollregion=(0, 0, max(cw, rw), max(ch, rh)))
                canvas.itemconfig(bid, width=max(cw, rw))
                sc_lbl.config(wraplength=max(240, cw - 30))
            except Exception:
                pass

        # عجلة الفأرة جوه النافذة (من غير ما تأثر على الشاشة الرئيسية)
        def wheel(e):
            num = getattr(e, "num", 0)
            if num == 4:
                step = -1
            elif num == 5:
                step = 1
            else:
                step = -1 if (getattr(e, "delta", 0) or 0) > 0 else 1
            canvas.yview_scroll(step * 3, "units")
            return "break"
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            win.bind(seq, wheel, add="+")

        tk.Label(body, text=f"① الثيمات ({len(THEME_ORDER)})", bg=BG, fg=INK, font=(FONT, 13, "bold"), anchor="e").pack(fill="x", pady=(6, 2))
        grid = tk.Frame(body, bg=BG); grid.pack(fill="x")
        for i, key in enumerate(THEME_ORDER):
            pal = THEMES[key]
            on = THEME == key
            card = tk.Frame(grid, bg=pal["WHITE"], highlightbackground=pal["BLUE"] if on else LINE,
                            highlightthickness=3 if on else 1)
            r, c = divmod(i, 4)
            card.grid(row=r, column=3 - c, padx=5, pady=5, sticky="nsew")
            grid.grid_columnconfigure(c, weight=1, uniform="pal")
            sw = tk.Canvas(card, width=150, height=40, highlightthickness=0, bd=0, bg=pal["WHITE"])
            sw.pack(fill="x", padx=6, pady=(6, 2))
            for j, col in enumerate([pal["NAVY_2"], pal["BLUE"], pal["GOLD"], pal["BG"], pal["WHITE"]]):
                sw.create_rectangle(j * 30, 0, (j + 1) * 30, 40, outline="", fill=col)
            tk.Label(card, text=THEME_LABELS.get(key, key), bg=pal["WHITE"], fg=pal["TEXT"],
                     font=(FONT, 10, "bold")).pack(pady=(2, 0))
            tk.Button(card, text="✔ مختار" if on else "تطبيق", command=lambda k=key: self.apply_appearance(theme=k),
                      bg=GREEN if on else pal["BLUE"], fg="white", activebackground=GOLD, activeforeground=INK,
                      relief="flat", bd=0, font=(FONT, 9, "bold"), padx=10, pady=4,
                      cursor="hand2").pack(pady=(2, 6))

        tk.Label(body, text="ملاحظة: ثيم «ذهبي فاخر» بيغيّر شكل الأزرار والخط معاه تلقائيًا (هوية كاملة).",
                 bg=BG, fg=MUTED, font=(FONT, 9), anchor="e").pack(fill="x", pady=(2, 0))

        tk.Label(body, text="② نوع الخط", bg=BG, fg=INK, font=(FONT, 13, "bold"), anchor="e").pack(fill="x", pady=(12, 2))
        frow = tk.Frame(body, bg=WHITE, highlightbackground=LINE, highlightthickness=1); frow.pack(fill="x")
        for key, lbl in FONT_LABELS.items():
            on = FONT_KEY == key
            tk.Button(frow, text=lbl, command=lambda k=key: self.apply_appearance(font=k),
                      bg=GREEN if on else SOFT, fg="white" if on else TEXT, activebackground=BLUE,
                      activeforeground="white", relief="flat", bd=0, font=(UI_FONTS[key], 10, "bold"),
                      padx=12, pady=7, cursor="hand2").pack(side="right", expand=True, fill="x", padx=4, pady=8)

        tk.Label(body, text="③ حجم الخط", bg=BG, fg=INK, font=(FONT, 13, "bold"), anchor="e").pack(fill="x", pady=(10, 2))
        srow = tk.Frame(body, bg=WHITE, highlightbackground=LINE, highlightthickness=1); srow.pack(fill="x")
        for key, lbl in FONT_SIZE_LABELS.items():
            on = FONT_SIZE_KEY == key
            tk.Button(srow, text=lbl, command=lambda k=key: self.apply_appearance(font_size=k),
                      bg=GREEN if on else SOFT, fg="white" if on else TEXT, activebackground=BLUE,
                      activeforeground="white", relief="flat", bd=0, font=(FONT, 10, "bold"),
                      padx=12, pady=7, cursor="hand2").pack(side="right", expand=True, fill="x", padx=4, pady=8)

        tk.Label(body, text="④ شكل الأزرار", bg=BG, fg=INK, font=(FONT, 13, "bold"), anchor="e").pack(fill="x", pady=(10, 2))
        brow = tk.Frame(body, bg=WHITE, highlightbackground=LINE, highlightthickness=1); brow.pack(fill="x")
        for key, lbl in BTN_STYLES.items():
            on = BTN_STYLE == key
            tk.Button(brow, text=lbl, command=lambda k=key: self.apply_appearance(btn_style=k),
                      bg=GREEN if on else SOFT, fg="white" if on else TEXT, activebackground=BLUE,
                      activeforeground="white", relief="flat", bd=0, font=(FONT, 10, "bold"),
                      padx=12, pady=7, cursor="hand2").pack(side="right", expand=True, fill="x", padx=4, pady=8)

        tk.Label(body, text="⑤ كثافة العرض", bg=BG, fg=INK, font=(FONT, 13, "bold"), anchor="e").pack(fill="x", pady=(10, 2))
        drow = tk.Frame(body, bg=WHITE, highlightbackground=LINE, highlightthickness=1); drow.pack(fill="x")
        for key, st in UI_STYLES.items():
            on = STYLE == key
            tk.Button(drow, text=f"{st['label']}  (صفوف {st['ROW_H']}px)", command=lambda k=key: self.apply_appearance(style=k),
                      bg=GREEN if on else SOFT, fg="white" if on else TEXT, activebackground=BLUE,
                      activeforeground="white", relief="flat", bd=0, font=(FONT, 10, "bold"),
                      padx=12, pady=7, cursor="hand2").pack(side="right", expand=True, fill="x", padx=4, pady=8)

        sc_lbl = tk.Label(body, text="⌨️ اختصارات: F1 الرئيسية · F2 ملاحظات · F3 قاعدة بيانات العاملين · "
                                     "F4 قاعدة بيانات موظفي الجهات · F5 حضور وانصراف العاملين · F6 حضور وانصراف الجهات · "
                                     "F7 زيارات المركز · F8 الصيانة · F9 المشتريات · F10 استخراج وطباعة التقارير · "
                                     "F11 مقارنة الفترات · F12 المظهر · Alt+→/← التنقّل بين الأقسام · Esc إغلاق النافذة",
                            bg=BG, fg=MUTED, font=(FONT, 9), anchor="e", justify="right", wraplength=760)
        sc_lbl.pack(fill="x", pady=(10, 6))

        self.btn(win, "إغلاق", win.destroy, MUTED).pack(side="bottom", pady=10)
        body.bind("<Configure>", fit)
        canvas.bind("<Configure>", fit)
        win.after(60, fit)

    def apply_appearance(self, theme=None, style=None, font=None, font_size=None, btn_style=None):
        """يطبّق الثيم/الخط/الحجم/شكل الأزرار، يحفظه في الإعدادات، ويعيد بناء الواجهة فورًا."""
        cfg = load_settings()
        if theme:
            apply_theme(theme); cfg["theme"] = theme
            tdef = THEMES.get(theme, {})
            # بعض الثيمات بتحمل هويتها كاملة (شكل الأزرار + الخط) — تُطبَّق لو المستخدم مش حدّدها بنفسه
            if not btn_style and tdef.get("BTN"):
                apply_btn_style(tdef["BTN"]); cfg["btn_style"] = tdef["BTN"]
            if not font and tdef.get("FONT_KEY"):
                apply_font(tdef["FONT_KEY"]); cfg["font"] = tdef["FONT_KEY"]
        if style:
            apply_style(style); cfg["style"] = style
        if font:
            apply_font(font); cfg["font"] = font
        if font_size:
            apply_font_size(font_size); cfg["font_size"] = font_size
        if btn_style:
            apply_btn_style(btn_style); cfg["btn_style"] = btn_style
        save_settings(cfg)
        self.rebuild_ui()

    def toggle_lang(self):
        """تبديل لغة الواجهة بين العربية والإنجليزية وإعادة بناء الشاشة."""
        new = "ar" if i18n.is_en() else "en"
        i18n.set_lang(new)
        cfg = load_settings(); cfg["lang"] = new; save_settings(cfg)
        self.root.title(T(f"{APP_NAME} - نظام إدارة المركز - {BRANCH}"))
        self.rebuild_ui()

    def restore_backup(self, path=None):
        """إدراج (استرجاع) نسخة احتياطية من ملف .db أو نسخة كاملة .zip.

        لو path=None يفتح نافذة اختيار الملف. يرجع True عند نجاح الاسترجاع.
        """
        if not self.need_admin():
            return False
        if netmod and NET_MODE == netmod.MODE_CLIENT:
            messagebox.showinfo("إدراج نسخة احتياطية",
                                "إدراج النسخ الاحتياطية يتم على الجهاز الرئيسي (السيرفر) فقط،\n"
                                "لأن قاعدة البيانات كلها هناك.")
            return False
        if path is None:
            path = filedialog.askopenfilename(
                title="اختر ملف النسخة الاحتياطية (.db أو .zip)",
                initialdir=BACKUP_DIR if os.path.isdir(BACKUP_DIR) else BASE,
                filetypes=[("نسخة احتياطية", "*.db *.zip"), ("قاعدة بيانات", "*.db"),
                           ("نسخة كاملة (قاعدة + صور)", "*.zip"), ("كل الملفات", "*.*")])
        if not path:
            return False
        if os.path.abspath(path) == os.path.abspath(DB_PATH):
            messagebox.showwarning("نفس الملف", "هذا هو ملف قاعدة البيانات الحالي نفسه. اختر ملف نسخة احتياطية آخر.")
            return False
        tmpdir, photos_src, db_file = None, None, path
        try:
            try:
                db_file, photos_src, tmpdir = open_backup_archive(path)
            except zipfile.BadZipFile as ex:
                messagebox.showerror("ملف غير صالح", f"تعذر قراءة الملف المضغوط:\n{ex}")
                return False
            except ValueError as ex:
                messagebox.showerror("ملف غير صالح", str(ex))
                return False
            ok, msg, nrec = validate_backup_db(db_file)
            if not ok:
                messagebox.showerror("ملف غير صالح", msg)
                return False
            when = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d/%m/%Y %I:%M %p")
            extra = "\nالصور: سيتم دمج صور النسخة مع الصور الحالية." if photos_src else ""
            if not messagebox.askyesno(
                    "تأكيد إدراج النسخة الاحتياطية",
                    f"سيتم استبدال كل البيانات الحالية ببيانات هذه النسخة:\n{os.path.basename(path)}\n"
                    f"تاريخ الملف: {when}   |   عدد السجلات: {nrec}{extra}\n\n"
                    "سيحفظ البرنامج نسخة من بياناتك الحالية تلقائيًا قبل الاستبدال.\nهل تريد المتابعة؟", icon="warning"):
                return False
            try:
                safety = do_restore(db_file, photos_src)
            except Exception as ex:
                messagebox.showerror("خطأ", f"تعذر إدراج النسخة (تم الإبقاء على بياناتك الحالية):\n{ex}")
                return False
            messagebox.showinfo("تم", "تم إدراج النسخة الاحتياطية بنجاح.\n"
                                      f"نسخة بياناتك السابقة محفوظة في:\n{safety}\n\n"
                                      + ("تم أيضًا دمج صور العاملين والفواتير." if photos_src
                                         else "ملاحظة: الصور في فولدر photos لا تتغير بهذه العملية."))
            self.go("home")
            return True
        finally:
            if tmpdir:
                shutil.rmtree(tmpdir, ignore_errors=True)

    def logout(self):
        if not messagebox.askyesno("تسجيل الخروج", "هل تريد تسجيل الخروج؟"):
            return
        for w in self.root.winfo_children():
            w.destroy()
        self.root.withdraw()
        u = login_dialog(self.root)
        if not u:
            self.root.destroy(); return
        self.user = u; self.nav_buttons = {}
        self.root.deiconify(); self.root.configure(bg=BG)
        self.setup_style(); self.build_shell(); self.go("home")

    def change_password_dialog(self):
        win = self._dialog("تغيير كلمة المرور", "440x360")
        tk.Label(win, text="🔑 تغيير كلمة المرور", bg=BG, fg=INK, font=(FONT, 15, "bold")).pack(anchor="e", padx=20, pady=(14, 6))
        vs = [tk.StringVar() for _ in range(3)]
        for lab, v in zip(("كلمة المرور الحالية", "كلمة المرور الجديدة", "تأكيد كلمة المرور الجديدة"), vs):
            tk.Label(win, text=lab, bg=BG, fg=TEXT, font=(FONT, 10, "bold"), anchor="e").pack(fill="x", padx=22, pady=(6, 2))
            ttk.Entry(win, textvariable=v, show="●", justify="right", width=40).pack(padx=22)

        def ok():
            old, new, conf = (v.get() for v in vs)
            with db() as con:
                r = con.execute("SELECT pw_hash,salt FROM users WHERE id=?", (self.user["id"],)).fetchone()
            if not r or not hmac.compare_digest(hash_pw(old, r[1]), r[0]):
                messagebox.showerror("كلمة المرور", "كلمة المرور الحالية غير صحيحة.", parent=win); return
            if len(new) < 4:
                messagebox.showwarning("كلمة المرور", "كلمة المرور الجديدة لا تقل عن 4 خانات.", parent=win); return
            err = validate_password(new, conf)
            if err:
                messagebox.showwarning("كلمة المرور", err, parent=win); return
            salt = secrets.token_hex(16)
            with db() as con:
                con.execute("UPDATE users SET pw_hash=?, salt=? WHERE id=?", (hash_pw(new, salt), salt, self.user["id"]))
            win.destroy()
            messagebox.showinfo("تم", "تم تغيير كلمة المرور بنجاح.")
        row = tk.Frame(win, bg=BG); row.pack(pady=16)
        self.btn(row, "💾 حفظ", ok, GREEN).pack(side="right", padx=6)
        self.btn(row, "إلغاء", win.destroy, MUTED).pack(side="right", padx=6)

    def users_dialog(self):
        if not self.need_admin():
            return
        win = self._dialog("المستخدمون والصلاحيات", "820x520")
        self._minsize(win, 700, 420)
        bx = tk.Frame(win, bg=BG); bx.pack(side="bottom", fill="x", pady=10)
        tk.Label(win, text="👥 المستخدمون والصلاحيات", bg=BG, fg=INK, font=(FONT, 15, "bold")).pack(anchor="e", padx=20, pady=(12, 2))
        tk.Label(win, text="أضف مستخدمًا وحدد لكل قسم: ممنوع / عرض فقط / عرض + إضافة / تعديل وحذف كامل.",
                 bg=BG, fg=MUTED, font=(FONT, 9)).pack(anchor="e", padx=22)
        tree = self.make_tree(win, ("name", "role", "state", "pages"), ("اسم المستخدم", "الدور", "الحالة", "الأقسام المتاحة"),
                              (170, 90, 90, 380), 10)

        def load():
            tree.delete(*tree.get_children())
            with db() as con:
                rows = con.execute("SELECT id,username,is_admin,active,perms FROM users ORDER BY is_admin DESC, username").fetchall()
            for n, (uid, name, adm, act, perms) in enumerate(rows):
                try:
                    pm = json.loads(perms or "{}")
                except ValueError:
                    pm = {}
                if adm:
                    pages = "كل الأقسام + إدارة المستخدمين والنسخ الاحتياطي"
                else:
                    pages = "، ".join(f"{lbl} ({LEVEL_LABEL[pm[k]]})" for k, lbl in PERM_PAGES if pm.get(k, "none") != "none") or "لا شيء"
                tree.insert("", "end", iid=str(uid), tags=("odd",) if n % 2 else (),
                            values=(name, "مدير" if adm else "مستخدم", "مفعّل" if act else "موقوف", pages))
        load()

        def selected():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("اختر مستخدمًا", "حدد المستخدم من الجدول أولًا.", parent=win); return None
            return int(sel[0])

        def edit():
            uid = selected()
            if uid: self.user_edit_dialog(uid, load)

        def delete():
            uid = selected()
            if not uid: return
            if uid == self.user["id"]:
                messagebox.showwarning("غير ممكن", "لا يمكنك حذف حسابك الحالي.", parent=win); return
            with db() as con:
                r = con.execute("SELECT username,is_admin,active FROM users WHERE id=?", (uid,)).fetchone()
                others = con.execute("SELECT COUNT(*) FROM users WHERE is_admin=1 AND active=1 AND id<>?", (uid,)).fetchone()[0]
            if r and r[1] and r[2] and not others:
                messagebox.showwarning("غير ممكن", "لازم يفضل مدير واحد مفعّل على الأقل.", parent=win); return
            if not messagebox.askyesno("تأكيد الحذف", f"حذف المستخدم «{r[0]}» نهائيًا؟", parent=win):
                return
            with db() as con:
                con.execute("DELETE FROM users WHERE id=?", (uid,))
            load()
        tree.bind("<Double-1>", lambda e: edit())
        inner = tk.Frame(bx, bg=BG); inner.pack()
        self.btn(inner, "➕ إضافة مستخدم", lambda: self.user_edit_dialog(None, load), GREEN).pack(side="right", padx=6)
        self.btn(inner, "✏️ تعديل المحدد", edit, BLUE).pack(side="right", padx=6)
        self.btn(inner, "حذف المحدد", delete, RED).pack(side="right", padx=6)
        self.btn(inner, "إغلاق", win.destroy, MUTED).pack(side="right", padx=6)

    def user_edit_dialog(self, uid, refresh):
        ex = None
        if uid:
            with db() as con:
                ex = con.execute("SELECT username,is_admin,active,perms FROM users WHERE id=?", (uid,)).fetchone()
            if not ex:
                return
        try:
            cur = json.loads(ex[3] or "{}") if ex else {}
        except ValueError:
            cur = {}
        is_self = bool(uid and uid == self.user["id"])
        win = self._dialog("تعديل مستخدم" if uid else "إضافة مستخدم", "800x680")
        self._minsize(win, 740, 560)
        bx = tk.Frame(win, bg=BG); bx.pack(side="bottom", fill="x", pady=8)
        tk.Label(win, text=("✏️ تعديل مستخدم" if uid else "➕ إضافة مستخدم جديد"), bg=BG, fg=INK,
                 font=(FONT, 15, "bold")).pack(anchor="e", padx=20, pady=(10, 0))
        top = self.panel(win); top.pack(fill="x", padx=16, pady=6)
        name_var, pw_var = tk.StringVar(value=ex[0] if ex else ""), tk.StringVar()
        self.add_field(top, "اسم المستخدم", name_var, 0, 1, 220)
        pe = self.add_field(top, "كلمة المرور" + (" (اتركها فارغة لعدم التغيير)" if uid else " (4 خانات على الأقل)"), pw_var, 0, 0, 240)
        pe.configure(show="●")
        adm_var = tk.BooleanVar(value=bool(ex[1]) if ex else False)
        act_var = tk.BooleanVar(value=bool(ex[2]) if ex else True)
        ck = tk.Frame(top, bg=WHITE); ck.grid(row=1, column=0, columnspan=2, sticky="e", padx=14, pady=(0, 8))
        c_adm = tk.Checkbutton(ck, text="مدير (كل الصلاحيات + إدارة المستخدمين والنسخ الاحتياطي)", variable=adm_var,
                               bg=WHITE, fg=TEXT, selectcolor=ENTRY, activebackground=WHITE, activeforeground=TEXT, font=(FONT, 10))
        c_act = tk.Checkbutton(ck, text="الحساب مفعّل (إلغاء التفعيل يمنع الدخول)", variable=act_var,
                               bg=WHITE, fg=TEXT, selectcolor=ENTRY, activebackground=WHITE, activeforeground=TEXT, font=(FONT, 10))
        c_adm.pack(side="right", padx=10); c_act.pack(side="right", padx=10)
        if is_self:
            c_adm.configure(state="disabled"); c_act.configure(state="disabled")

        pm_outer = self.panel(win); pm_outer.pack(fill="both", expand=True, padx=16, pady=6)
        # مسطرة جانبية عشان كل الصلاحيات تظهر لأي عدد أقسام
        pm_canvas = tk.Canvas(pm_outer, bg=WHITE, highlightthickness=0, bd=0)
        pm_vsb = ttk.Scrollbar(pm_outer, orient="vertical", command=pm_canvas.yview)
        pm_canvas.configure(yscrollcommand=pm_vsb.set)
        pm_vsb.pack(side="left", fill="y")
        pm_canvas.pack(side="right", fill="both", expand=True)
        pm = tk.Frame(pm_canvas, bg=WHITE)
        pm_win = pm_canvas.create_window((0, 0), window=pm, anchor="nw")
        def _pm_fit(_e=None):
            try:
                pm_canvas.configure(scrollregion=pm_canvas.bbox("all"))
                pm_canvas.itemconfig(pm_win, width=pm_canvas.winfo_width())
            except Exception:
                pass
        pm.bind("<Configure>", _pm_fit)
        pm_canvas.bind("<Configure>", _pm_fit)
        def _pm_wheel(e):
            try:
                if getattr(e, "num", 0) == 4:
                    pm_canvas.yview_scroll(-3, "units")
                elif getattr(e, "num", 0) == 5:
                    pm_canvas.yview_scroll(3, "units")
                else:
                    pm_canvas.yview_scroll(-3 if getattr(e, "delta", 0) > 0 else 3, "units")
            except Exception:
                pass
            return "break"
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            win.bind(seq, _pm_wheel, add="+")
        tk.Label(pm, text="الصلاحيات حسب القسم", bg=WHITE, fg=INK, font=(FONT, 11, "bold")).grid(row=0, column=0, columnspan=5, sticky="e", padx=14, pady=(8, 2))
        for col, lv in zip((3, 2, 1, 0), LEVELS):
            tk.Label(pm, text=LEVEL_LABEL[lv], bg=WHITE, fg=MUTED, font=(FONT, 9, "bold")).grid(row=1, column=col, padx=8, pady=2)
        pvars, radios = {}, []
        for i, (key, label) in enumerate(PERM_PAGES, 2):
            tk.Label(pm, text=label, bg=WHITE, fg=TEXT, font=(FONT, 10, "bold"), anchor="e").grid(row=i, column=4, sticky="e", padx=12, pady=3)
            v = tk.StringVar(value=cur.get(key, "none") if cur.get(key, "none") in LEVELS else "none")
            if key in VIEW_ONLY_PAGES and v.get() in ("add", "edit"):
                v.set("view")
            pvars[key] = v
            for col, lv in zip((3, 2, 1, 0), LEVELS):
                if key in VIEW_ONLY_PAGES and lv in ("add", "edit"):
                    continue
                rb = ttk.Radiobutton(pm, value=lv, variable=v); rb.grid(row=i, column=col, padx=8, pady=3)
                radios.append(rb)
        for c in range(5):
            pm.grid_columnconfigure(c, weight=1)

        def preset(lv):
            for k, v in pvars.items():
                v.set("view" if (k in VIEW_ONLY_PAGES and lv in ("add", "edit")) else lv)
        pr = tk.Frame(win, bg=BG); pr.pack(fill="x", padx=16)
        tk.Label(pr, text="تعبئة سريعة:", bg=BG, fg=TEXT, font=(FONT, 9, "bold")).pack(side="right", padx=6)
        for lv in ("none", "view", "add", "edit"):
            tk.Button(pr, text=f"الكل: {LEVEL_LABEL[lv]}", command=lambda l=lv: preset(l), bg=SOFT, fg=TEXT, relief="flat",
                      font=(FONT, 9, "bold"), padx=10, pady=3, cursor="hand2").pack(side="right", padx=3)

        def sync_admin(*_):
            for rb in radios:
                rb.configure(state="disabled" if adm_var.get() else "normal")
        adm_var.trace_add("write", sync_admin); sync_admin()

        def save():
            name, pw = name_var.get().strip(), pw_var.get()
            if not name:
                messagebox.showwarning("بيانات ناقصة", "اكتب اسم المستخدم.", parent=win); return
            if (not uid and len(pw) < 4) or (pw and len(pw) < 4):
                messagebox.showwarning("كلمة المرور", "كلمة المرور لا تقل عن 4 خانات.", parent=win); return
            is_ad, act = adm_var.get(), act_var.get()
            perms = {k: v.get() for k, v in pvars.items() if v.get() != "none"}
            with db() as con:
                if con.execute("SELECT 1 FROM users WHERE username=? AND id<>?", (name, uid or 0)).fetchone():
                    messagebox.showwarning("اسم مكرر", "يوجد مستخدم بنفس الاسم.", parent=win); return
                if uid:
                    others = con.execute("SELECT COUNT(*) FROM users WHERE is_admin=1 AND active=1 AND id<>?", (uid,)).fetchone()[0]
                    if not others and not (is_ad and act):
                        messagebox.showwarning("غير ممكن", "لازم يفضل مدير واحد مفعّل على الأقل.", parent=win); return
                    if pw:
                        salt = secrets.token_hex(16)
                        con.execute("UPDATE users SET username=?,pw_hash=?,salt=?,is_admin=?,active=?,perms=? WHERE id=?",
                                    (name, hash_pw(pw, salt), salt, int(is_ad), int(act), json.dumps(perms), uid))
                    else:
                        con.execute("UPDATE users SET username=?,is_admin=?,active=?,perms=? WHERE id=?",
                                    (name, int(is_ad), int(act), json.dumps(perms), uid))
                else:
                    salt = secrets.token_hex(16)
                    con.execute("INSERT INTO users(username,pw_hash,salt,is_admin,active,perms,created_at) VALUES(?,?,?,?,?,?,?)",
                                (name, hash_pw(pw, salt), salt, int(is_ad), int(act), json.dumps(perms), now_stamp()))
            if is_self:
                self.user["username"] = name
            win.destroy(); refresh()
        inner = tk.Frame(bx, bg=BG); inner.pack()
        self.btn(inner, "💾 حفظ", save, GREEN).pack(side="right", padx=6)
        self.btn(inner, "إلغاء", win.destroy, MUTED).pack(side="right", padx=6)

    def manual_backup(self, full=None):
        """إنشاء نسخة احتياطية الآن. full=None يسأل: كاملة (قاعدة + صور) أو سريعة (قاعدة فقط)."""
        if not self.need_admin():
            return
        if full is None:
            ans = messagebox.askyesnocancel(
                "نسخة احتياطية",
                "اختر نوع النسخة:\n\n"
                "• «نعم» = نسخة كاملة (قاعدة البيانات + صور العاملين والفواتير)\n"
                "• «لا» = نسخة سريعة (قاعدة البيانات فقط)\n\n"
                "النسخة الكاملة بتاخد وقت وحجم أكبر، بس بتأمّن الصور كمان.")
            if ans is None:
                return
            full = bool(ans)
        if netmod and NET_MODE == netmod.MODE_CLIENT:
            try:
                res = netmod.request_backup(NET_CFG, full=full)
                if res.get("ok"):
                    messagebox.showinfo("نسخة احتياطية",
                                        f"تم إنشاء النسخة على الجهاز الرئيسي:\n{res.get('path')}")
                else:
                    messagebox.showerror("خطأ", res.get("error") or "تعذر إنشاء النسخة")
            except Exception as ex:
                messagebox.showerror("خطأ", f"تعذر الاتصال بالجهاز الرئيسي:\n{ex}")
            return
        try:
            dest = backup_full_now() if full else backup_now()
            messagebox.showinfo("نسخة احتياطية",
                                ("تم حفظ نسخة كاملة (قاعدة + صور):\n" if full
                                 else "تم حفظ نسخة من قاعدة البيانات:\n") + dest)
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر إنشاء النسخة:\n{ex}")

    def backups_dialog(self):
        """شاشة إدارة النسخ الاحتياطية: عرض + إنشاء (سريعة/كاملة) + استرجاع + حذف + فتح المجلد."""
        if not self.need_admin():
            return
        is_client = bool(netmod and NET_MODE == netmod.MODE_CLIENT)
        win = self._dialog("النسخ الاحتياطية", "1000x680")
        self._minsize(win, 880, 560)

        head = tk.Frame(win, bg=BG); head.pack(fill="x", padx=22, pady=(14, 0))
        tk.Label(head, text="🗂️ النسخ الاحتياطية", bg=BG, fg=INK, font=(FONT, 16, "bold")).pack(anchor="e")
        tk.Label(head, text="نسخة كاملة = قاعدة البيانات + صور العاملين والفواتير   |   نسخة سريعة = قاعدة البيانات فقط",
                 bg=BG, fg=MUTED, font=(FONT, 9)).pack(anchor="e", pady=(2, 0))
        if is_client:
            tk.Label(head, text="⚠️ النسخ محفوظة على الجهاز الرئيسي — الاسترجاع يتم عليه فقط.",
                     bg=BG, fg=RED, font=(FONT, 9, "bold")).pack(anchor="e", pady=(2, 0))

        info = tk.Label(win, text="", bg=BG, fg=TEXT, font=(FONT, 10, "bold"))
        info.pack(anchor="e", padx=24, pady=(6, 0))

        wrap = tk.Frame(win, bg=BG); wrap.pack(fill="both", expand=True, padx=20, pady=8)
        cols, heads, widths = ("kind", "when", "size", "name"), ("النوع", "التاريخ والوقت", "الحجم", "الملف"), (150, 190, 110, 420)
        tree = ttk.Treeview(wrap, columns=cols, show="headings", height=13)
        for c, h, w in zip(cols, heads, widths):
            tree.heading(c, text=h); tree.column(c, width=w, anchor="e")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y"); tree.pack(side="right", fill="both", expand=True)

        by_iid = {}

        def human(n):
            return ("%.1f ميجا" % (n / 1048576) if n >= 1048576
                    else ("%.0f كيلوبايت" % (n / 1024) if n >= 1024 else "%d بايت" % n))

        def refresh():
            tree.delete(*tree.get_children()); by_iid.clear()
            try:
                rows = netmod.list_backups(NET_CFG) if is_client else list_backups()
            except Exception as ex:
                info.config(text=f"تعذر الاتصال بالجهاز الرئيسي: {ex}", fg=RED); return
            total = 0
            for r in rows:
                when = datetime.fromtimestamp(r["mtime"]).strftime("%d/%m/%Y %I:%M %p")
                iid = tree.insert("", "end", values=(r["kind"], when, human(r["size"]), r["name"]))
                by_iid[iid] = r; total += r["size"]
            info.config(text="عدد النسخ: %d   |   الحجم الإجمالي: %s   |   المجلد: %s"
                             % (len(rows), human(total), BACKUP_DIR if not is_client else "على الجهاز الرئيسي"), fg=TEXT)

        def sel():
            s = tree.selection()
            return by_iid.get(s[0]) if s else None

        def do_quick():
            try:
                if is_client:
                    res = netmod.request_backup(NET_CFG, full=False)
                    if not res.get("ok"):
                        raise RuntimeError(res.get("error") or "تعذر إنشاء النسخة")
                    messagebox.showinfo("نسخة سريعة", f"تم إنشاء نسخة على الجهاز الرئيسي:\n{res.get('path')}")
                else:
                    dest = backup_now()
                    messagebox.showinfo("نسخة سريعة", f"تم حفظ نسخة من قاعدة البيانات:\n{dest}")
                refresh()
            except Exception as ex:
                messagebox.showerror("خطأ", f"تعذر إنشاء النسخة:\n{ex}")

        def do_full():
            try:
                if is_client:
                    res = netmod.request_backup(NET_CFG, full=True)
                    if not res.get("ok"):
                        raise RuntimeError(res.get("error") or "تعذر إنشاء النسخة")
                    messagebox.showinfo("نسخة كاملة", f"تم إنشاء نسخة كاملة على الجهاز الرئيسي:\n{res.get('path')}")
                else:
                    dest = backup_full_now()
                    messagebox.showinfo("نسخة كاملة", f"تم حفظ نسخة كاملة (قاعدة + صور):\n{dest}")
                refresh()
            except Exception as ex:
                messagebox.showerror("خطأ", f"تعذر إنشاء النسخة:\n{ex}")

        def do_restore():
            r = sel()
            if not r:
                messagebox.showinfo("استرجاع", "اختر نسخة من القائمة أولًا."); return
            if is_client:
                messagebox.showinfo("استرجاع", "الاسترجاع يتم على الجهاز الرئيسي (السيرفر) فقط."); return
            self.restore_backup(r["path"]); refresh()

        def do_delete():
            r = sel()
            if not r:
                messagebox.showinfo("حذف", "اختر نسخة من القائمة أولًا."); return
            if not messagebox.askyesno("حذف نسخة", f"حذف هذه النسخة نهائيًا؟\n\n{r['name']}", icon="warning"):
                return
            try:
                os.remove(r["path"]); refresh()
            except OSError as ex:
                messagebox.showerror("خطأ", f"تعذر حذف الملف:\n{ex}")

        def open_dir():
            try:
                if is_client:
                    messagebox.showinfo("المجلد", "مجلد النسخ موجود على الجهاز الرئيسي.")
                    return
                os.makedirs(BACKUP_DIR, exist_ok=True)
                if sys.platform == "win32":
                    os.startfile(BACKUP_DIR)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", BACKUP_DIR])
                else:
                    subprocess.Popen(["xdg-open", BACKUP_DIR])
            except Exception as ex:
                messagebox.showerror("خطأ", f"تعذر فتح المجلد:\n{ex}")

        bar = tk.Frame(win, bg=BG); bar.pack(fill="x", padx=20, pady=(0, 14))
        self.btn(bar, "🗜️ نسخة كاملة (قاعدة + صور)", do_full, GREEN).pack(side="right", padx=5)
        self.btn(bar, "💾 نسخة سريعة", do_quick, PRIMARY).pack(side="right", padx=5)
        self.btn(bar, "♻️ استرجاع المختار", do_restore, GOLD).pack(side="right", padx=5)
        self.btn(bar, "🗑️ حذف المختار", do_delete, RED).pack(side="right", padx=5)
        self.btn(bar, "🔄 تحديث", refresh, MUTED).pack(side="left", padx=5)
        self.btn(bar, "📂 فتح المجلد", open_dir, MUTED).pack(side="left", padx=5)
        self.btn(bar, "إغلاق", win.destroy, MUTED).pack(side="left", padx=5)
        refresh()

    # ---------- الشبكة (LAN) ----------
    def footer_text(self):
        base = f"تصميم وتطوير: {DESIGNER}   |   {APP_NAME} © {date.today().year}"
        st = self.net_status_text()
        return base + ("      " + st if st else "")

    def net_status_text(self):
        if not netmod or NET_MODE == netmod.MODE_STANDALONE:
            return ""
        if NET_MODE == netmod.MODE_SERVER:
            return (f"🟢 الجهاز الرئيسي (سيرفر) — {netmod.local_ip()}:{NET_CFG.get('port')}"
                    f"   |   محطات متصلة: {NET_STATUS.get('stations', 0)}")
        host = NET_CFG.get("server_host") or "?"
        return (f"🟢 متصل بالجهاز الرئيسي: {host}") if NET_STATUS.get("ok") else (f"🔴 غير متصل بالجهاز الرئيسي: {host}")

    def refresh_net_status(self):
        if not netmod or NET_MODE == netmod.MODE_STANDALONE:
            return
        if NET_MODE == netmod.MODE_CLIENT:
            try:
                info = netmod.ping(NET_CFG, timeout=3)
                NET_STATUS["ok"] = True
                NET_STATUS["stations"] = info.get("stations", 0)
            except Exception:
                NET_STATUS["ok"] = False
        else:
            NET_STATUS["ok"] = True
            NET_STATUS["stations"] = len(NET_SERVER.stations) if NET_SERVER else 0
        try:
            if getattr(self, "footer_lbl", None):
                self.footer_lbl.config(text=self.footer_text(),
                                       fg=GOLD if NET_STATUS.get("ok") else "#ff9b9b")
        except Exception:
            pass

    def _can_auto_refresh(self):
        """ما نعملش تحديث تلقائي وإحنا في نص إدخال أو نافذة مفتوحة (حتى لا تضيع البيانات)."""
        try:
            if self.root.grab_current() is not None:
                return False
            f = self.root.focus_get()
            if f is not None and f.winfo_class() in ("Entry", "TEntry", "TCombobox", "Text", "Spinbox"):
                return False
        except Exception:
            return False
        return True

    def auto_refresh(self):
        """في وضع المحطة: يعيد تحميل الصفحة لو عدّل جهاز آخر البيانات على السيرفر."""
        try:
            secs = int(NET_CFG.get("auto_refresh", 20) or 0)
        except Exception:
            secs = 0
        if netmod and NET_MODE == netmod.MODE_CLIENT and secs > 0:
            try:
                info = netmod.ping(NET_CFG, timeout=3)
                NET_STATUS["ok"] = True
                v = info.get("version")
                if getattr(self, "_net_version", None) is None:
                    self._net_version = v
                elif v != self._net_version:
                    self._net_version = v
                    if self._can_auto_refresh():
                        self.go(self.active_key or "home")
                    else:
                        # حتى لو المستخدم بيكتب، نحدّث شريط الملاحظات الجديدة بس
                        try: self.build_notes_banner()
                        except Exception: pass
            except Exception:
                NET_STATUS["ok"] = False
        if getattr(self, "_auto_job", None):
            try: self.root.after_cancel(self._auto_job)
            except Exception: pass
        self._auto_job = self.root.after(max(5, secs or 20) * 1000, self.auto_refresh)

    def net_dialog(self):
        if not self.need_admin():
            return
        if not netmod:
            messagebox.showerror("الشبكة", "ملف net.py غير موجود بجوار البرنامج."); return
        win = tk.Toplevel(self.root)
        win.title("الشبكة المحلية (LAN)"); win.configure(bg=BG); win.resizable(False, False)
        w, h = 580, 600
        try:
            h = min(h, max(380, int(win.winfo_screenheight()) - 120))
        except Exception:
            pass
        win.geometry(f"{w}x{h}+{(win.winfo_screenwidth() - w) // 2}+{max(0, (win.winfo_screenheight() - h) // 2 - 30)}")
        tk.Label(win, text="🌐  إعداد الشبكة المحلية", bg=BG, fg=INK, font=(FONT, 15, "bold")).pack(pady=(16, 2))
        tk.Label(win, text="خلّي جهازًا واحدًا «رئيسي (سيرفر)» والباقي «محطة»، واكتب نفس المنفذ وكلمة الربط في الكل.",
                 bg=BG, fg=MUTED, font=(FONT, 9), wraplength=520, justify="center").pack()

        mode = tk.StringVar(value=NET_CFG.get("mode", "standalone"))
        host = tk.StringVar(value=NET_CFG.get("server_host", ""))
        port = tk.StringVar(value=str(NET_CFG.get("port", 8765)))
        token = tk.StringVar(value=NET_CFG.get("token", ""))
        station = tk.StringVar(value=NET_STATION or "")
        refresh = tk.StringVar(value=str(NET_CFG.get("auto_refresh", 20)))

        body = tk.Frame(win, bg=BG); body.pack(fill="x", padx=24, pady=8)
        tk.Label(body, text="وضع الجهاز", bg=BG, fg=TEXT, font=(FONT, 10, "bold"), anchor="e").pack(fill="x", pady=(4, 2))
        for val, lbl in (("standalone", "مستقل (قاعدة بيانات محلية على نفس الجهاز)"),
                         ("server", "جهاز رئيسي / سيرفر (يملك قاعدة البيانات)"),
                         ("client", "محطة (تبعث البيانات للجهاز الرئيسي)")):
            tk.Radiobutton(body, text=lbl, value=val, variable=mode, bg=BG, fg=TEXT, anchor="e",
                           selectcolor=WHITE, activebackground=BG, font=(FONT, 10, "bold")).pack(fill="x", padx=30)

        def row(label, var, show=None):
            fr = tk.Frame(body, bg=BG); fr.pack(fill="x", pady=4)
            tk.Label(fr, text=label, bg=BG, fg=TEXT, font=(FONT, 10, "bold"), width=22, anchor="e").pack(side="right", padx=6)
            e = tk.Entry(fr, textvariable=var, font=(FONT, 11), justify="right", bg=ENTRY, fg=TEXT, relief="solid", bd=1)
            if show: e.config(show=show)
            e.pack(side="right", fill="x", expand=True, padx=6)
            return e

        row("عنوان الجهاز الرئيسي", host)
        row("المنفذ (Port)", port)
        row("كلمة الربط (Token)", token)
        row("اسم المحطة", station)
        row("تحديث تلقائي (ثواني)", refresh)

        info = tk.Label(win, text="", bg=BG, fg=PRIMARY, font=(FONT, 9, "bold"), wraplength=520, justify="center")
        info.pack(pady=4)

        def show_info():
            if mode.get() == "server":
                info.config(text=f"عنوان هذا الجهاز على الشبكة: {netmod.local_ip()}   (المنفذ {port.get()})\n"
                                 "اكتب نفس المنفذ وكلمة الربط في كل محطة، واسمح للمنفذ في جدار الحماية (Windows Firewall).",
                            fg=GREEN)
            elif mode.get() == "client":
                info.config(text="اكتب عنوان الجهاز الرئيسي (مثال: 192.168.1.10) ونفس كلمة الربط.", fg=PRIMARY)
            else:
                info.config(text="الوضع المستقل: كل البيانات محفوظة على هذا الجهاز فقط.", fg=PRIMARY)
        mode.trace_add("write", lambda *_: show_info()); show_info()

        def gen_token():
            token.set(netmod.new_token()); show_info()
        def test():
            try:
                cfg = dict(NET_CFG); cfg.update({"mode": "client", "server_host": host.get().strip(),
                                                 "port": int(port.get() or 8765), "token": token.get().strip()})
                res = netmod.ping(cfg, timeout=5)
                info.config(text=f"✅ تم الاتصال بالجهاز الرئيسي بنجاح (الإصدار {res.get('proto')}).", fg=GREEN)
            except Exception as ex:
                info.config(text=f"❌ فشل الاتصال: {ex}", fg=RED)
        def save():
            try:
                p = int(port.get())
            except ValueError:
                messagebox.showwarning("خطأ", "المنفذ لازم يكون رقم.", parent=win); return
            try:
                r = int(refresh.get() or 0)
            except ValueError:
                r = 20
            NET_CFG.update({"mode": mode.get(), "server_host": host.get().strip(), "port": p,
                            "token": token.get().strip(), "station": station.get().strip(),
                            "auto_refresh": r})
            try:
                netmod.save_config(BASE, NET_CFG)
            except Exception as ex:
                messagebox.showerror("خطأ", f"تعذر حفظ الإعدادات:\n{ex}", parent=win); return
            messagebox.showinfo("تم", "تم حفظ إعدادات الشبكة.\nأعد تشغيل البرنامج لتطبيقها.", parent=win)
            win.destroy()

        btns = tk.Frame(win, bg=BG); btns.pack(pady=10)
        self.btn(btns, "💾 حفظ", save, GREEN).pack(side="right", padx=5)
        self.btn(btns, "🔑 توليد كلمة ربط", gen_token, BLUE).pack(side="right", padx=5)
        self.btn(btns, "🔌 اختبار الاتصال", test, PRIMARY).pack(side="right", padx=5)
        self.btn(btns, "إلغاء", win.destroy, MUTED).pack(side="right", padx=5)
        tk.Label(win, text=f"الوضع الحالي: {NET_CFG.get('mode')}   |   محطة هذا الجهاز: {NET_STATION}",
                 bg=BG, fg=MUTED, font=(FONT, 9)).pack(side="bottom", pady=8)
        win.bind("<Escape>", lambda e: win.destroy())
        try: win.grab_set()
        except Exception: pass

    def audit_dialog(self):
        if not self.need_admin():
            return
        if netmod and NET_MODE == netmod.MODE_CLIENT:
            try:
                rows = netmod.fetch_audit(NET_CFG)
            except Exception as ex:
                messagebox.showerror("سجل العمليات", f"تعذر جلب السجل من الجهاز الرئيسي:\n{ex}"); return
        else:
            rows = netmod.audit_rows(DB_PATH) if netmod else []
        win = tk.Toplevel(self.root)
        win.title("سجل العمليات"); win.configure(bg=BG)
        w, h = 1000, 620
        try:
            sw, sh = int(win.winfo_screenwidth()), int(win.winfo_screenheight())
            w, h = min(w, max(420, sw - 40)), min(h, max(320, sh - 120))
        except Exception:
            pass
        win.geometry(f"{w}x{h}+{max(0, (win.winfo_screenwidth() - w) // 2)}+{max(0, (win.winfo_screenheight() - h) // 2 - 30)}")
        tk.Label(win, text="📋  سجل العمليات (مين أضاف/عدّل/حذف ومن أنهي جهاز)", bg=BG, fg=INK,
                 font=(FONT, 14, "bold")).pack(pady=10)
        cols = ("ts", "username", "station", "action", "table_name", "detail")
        heads = ("الوقت", "المستخدم", "الجهاز", "العملية", "الجدول", "تفاصيل")
        frame = tk.Frame(win, bg=BG); frame.pack(fill="both", expand=True, padx=12, pady=6)
        tv = ttk.Treeview(frame, columns=cols, show="headings", height=20)
        for c, hd in zip(cols, heads):
            tv.heading(c, text=hd)
            tv.column(c, width=140 if c != "detail" else 330, anchor="center")
        vs = ttk.Scrollbar(frame, orient="vertical", command=tv.yview); tv.configure(yscrollcommand=vs.set)
        tv.pack(side="right", fill="both", expand=True); vs.pack(side="left", fill="y")
        act_ar = {"INSERT": "إضافة", "UPDATE": "تعديل", "DELETE": "حذف"}
        for r in rows:
            tv.insert("", "end", values=(r[0], r[1], r[2], act_ar.get(r[3], r[3]), r[4], r[5]))
        tk.Button(win, text="إغلاق", command=win.destroy, bg=MUTED, fg="white", relief="flat", bd=0,
                  font=(FONT, 10, "bold"), padx=24, pady=7, cursor="hand2").pack(pady=8)
        try: win.grab_set()
        except Exception: pass

    def tick(self):
        days = ["الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]
        n = datetime.now()
        try:
            self.hdr.itemconfig(self.clock_id, text=f"{days[n.weekday()]}  {n.strftime('%d/%m/%Y')}   {n.strftime('%I:%M %p')}")
        except Exception:
            pass
        if getattr(self, "_tick_job", None):
            try: self.root.after_cancel(self._tick_job)
            except Exception: pass
        self._tick_job = self.root.after(30000, self.tick)
        if netmod and NET_MODE != netmod.MODE_STANDALONE:
            self.refresh_net_status()
            if not getattr(self, "_auto_job", None):
                self.auto_refresh()

    def nav_hover(self, key, on):
        if key != self.active_key:
            self.nav_buttons[key].config(bg=BLUE if on else NAVY_2)

    def go(self, key):
        if key != "home" and not self.can(key, "view"):
            self.need(key, "view"); return
        self.active_key = key
        for k, b in self.nav_buttons.items():
            b.config(bg=GOLD if k == key else NAVY_2, fg=NAVY if k == key else "white")
        self.clear_content()
        self.build_notes_banner()
        if key == "home": self.show_home()
        elif key == "notes": self.show_notes()
        elif key == "workers": self.show_workers()
        elif key == "visits": self.show_visits()
        elif key == "maintenance": self.show_maintenance()
        elif key == "employees": self.show_employees()
        elif key == "attend": self.show_attendance()
        elif key == "gov": self.show_gov()
        elif key == "govstaff": self.show_govstaff()
        elif key == "compare": self.show_compare()
        elif key == "reports": self.show_reports()
        elif key == "purchases": self.show_purchases()
        else: self.show_inventory(key)
        lv = self.level(key) if key != "home" else "edit"
        if lv in ("view", "add") and key not in VIEW_ONLY_PAGES:
            txt = ("🔒 صلاحيتك في هذا القسم: عرض فقط (بدون إضافة أو تعديل أو حذف)" if lv == "view"
                   else "🔒 صلاحيتك في هذا القسم: عرض وإضافة فقط (بدون تعديل أو حذف)")
            kids = self.content.winfo_children()
            lbl = tk.Label(self.content, text=txt, bg=SOFT, fg=RED, font=(FONT, 10, "bold"), anchor="e", padx=12, pady=4)
            if kids: lbl.pack(fill="x", padx=24, pady=(8, 0), before=kids[0])
            else: lbl.pack(fill="x", padx=24, pady=(8, 0))
        # ارجع لأول الصفحة وأعيد حساب التمرير بعد بناء القسم
        try:
            self.content_canvas.yview_moveto(0)
            self.content_canvas.xview_moveto(0)
            self.content.update_idletasks()
            self._last_size = None
            self._fit_content()
        except Exception:
            pass

    def clear_content(self):
        for w in self.content.winfo_children():
            w.destroy()
        self._last_size = None

    def page_title(self, title, subtitle=""):
        tk.Label(self.content, text=title, bg=BG, fg=INK, font=(FONT, 21, "bold"),
                 anchor="e").pack(fill="x", padx=26, pady=(18, 0))
        tk.Frame(self.content, bg=GOLD, height=3, width=70).pack(anchor="e", padx=28, pady=(3, 4))
        if subtitle:
            tk.Label(self.content, text=subtitle, bg=BG, fg=MUTED, font=(FONT, 10),
                     anchor="e").pack(fill="x", padx=27, pady=(0, 10))

    def panel(self, parent=None):
        return tk.Frame(parent or self.content, bg=WHITE, highlightbackground=LINE, highlightthickness=1)

    def btn(self, parent, text, cmd, color=None, **kw):
        color = color or PRIMARY
        if BTN_STYLE == "outline":
            cfg = dict(bg=WHITE, fg=color, activebackground=color, activeforeground="white",
                       highlightbackground=color, highlightcolor=color, highlightthickness=1, relief="flat")
        elif BTN_STYLE == "soft":
            cfg = dict(bg=lerp_color(WHITE, color, 0.16), fg=color, activebackground=color,
                       activeforeground="white", relief="flat")
        elif BTN_STYLE == "gold":
            # ذهبي فاخر: تعبئة داكنة + إطار ذهبي + نص ذهبي (عند المرور يظهر لون الزر الأصلي)
            cfg = dict(bg=WHITE, fg=GOLD, activebackground=color, activeforeground=NAVY_2,
                       highlightbackground=GOLD, highlightcolor=GOLD, highlightthickness=1, relief="flat")
        else:
            cfg = dict(bg=color, fg="white", activebackground=BLUE, activeforeground="white", relief="flat")
        cfg.update(kw)
        return tk.Button(parent, text=text, command=cmd, font=(FONT, 10, "bold"),
                         padx=BTN_PADX, pady=BTN_PADY, cursor="hand2", **cfg)

    def add_field(self, parent, label, var, row, col, width):
        box = tk.Frame(parent, bg=WHITE)
        box.grid(row=row, column=col, padx=14, pady=8, sticky="e")
        tk.Label(box, text=label, bg=WHITE, fg=TEXT, font=(FONT, 10, "bold"), anchor="e").pack(fill="x", pady=(0, 4))
        ent = ttk.Entry(box, textvariable=var, width=max(22, int(width / 9)), justify="right")
        ent.pack()
        return ent

    def add_date(self, parent, label, var, row, col, optional=False):
        """تاريخ بثلاث قوائم منسدلة (يوم / شهر / سنة) مرتبطة بمتغير نصي بصيغة يوم/شهر/سنة."""
        box = tk.Frame(parent, bg=WHITE)
        box.grid(row=row, column=col, padx=14, pady=8, sticky="e")
        tk.Label(box, text=label, bg=WHITE, fg=TEXT, font=(FONT, 10, "bold"), anchor="e").pack(fill="x", pady=(0, 4))
        line = tk.Frame(box, bg=WHITE); line.pack(anchor="e")
        blank = [""] if optional else []
        y0 = date.today().year
        d_var, m_var, y_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        cd = ttk.Combobox(line, textvariable=d_var, values=blank + [f"{i:02d}" for i in range(1, 32)],
                          state="readonly", width=4, justify="center")
        cmo = ttk.Combobox(line, textvariable=m_var, values=blank + i18n.MONTHS, state="readonly", width=9, justify="center")
        cy = ttk.Combobox(line, textvariable=y_var, values=blank + [str(y) for y in range(y0 - 10, y0 + 3)],
                          state="readonly", width=6, justify="center")
        for w in (cd, cmo, cy):
            w.pack(side="right", padx=2)
        busy = {"v": False}

        def from_var(*_):
            if busy["v"]:
                return
            try:
                dt = parse_date(var.get())
                d_var.set(f"{dt.day:02d}"); m_var.set(i18n.MONTHS[dt.month - 1]); y_var.set(str(dt.year))
                if str(dt.year) not in list(cy["values"]):
                    cy.configure(values=blank + [str(y) for y in range(min(dt.year, y0 - 10), max(dt.year, y0 + 2) + 1)])
            except ValueError:
                if optional:
                    d_var.set(""); m_var.set(""); y_var.set("")
                else:
                    t = date.today()
                    d_var.set(f"{t.day:02d}"); m_var.set(i18n.MONTHS[t.month - 1]); y_var.set(str(t.year))
            except tk.TclError:
                pass

        def to_var(_=None):
            d, m, y = d_var.get(), m_var.get(), y_var.get()
            busy["v"] = True
            try:
                if not (d and m and y):
                    var.set("")
                    return
                mi, yi = i18n.month_index(m), int(y)
                di = min(int(d), calendar.monthrange(yi, mi)[1])      # لا يسمح بـ 31 فبراير
                if di != int(d):
                    d_var.set(f"{di:02d}")
                var.set(f"{di:02d}/{mi:02d}/{yi}")
            finally:
                busy["v"] = False
        for w in (cd, cmo, cy):
            w.bind("<<ComboboxSelected>>", to_var)
        var.trace_add("write", from_var)
        from_var()
        return box

    def add_combo(self, parent, label, var, values, row, col, readonly=True):
        box = tk.Frame(parent, bg=WHITE)
        box.grid(row=row, column=col, padx=14, pady=8, sticky="e")
        tk.Label(box, text=label, bg=WHITE, fg=TEXT, font=(FONT, 10, "bold"), anchor="e").pack(fill="x", pady=(0, 4))
        cb = ttk.Combobox(box, textvariable=var, values=values, state="readonly" if readonly else "normal", justify="right", width=22)
        cb.pack()
        return cb

    def make_tree(self, parent, columns, headings, widths, height=10):
        frame = tk.Frame(parent, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        frame.pack(fill="both", expand=True, padx=24, pady=(6, 10))
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=height, selectmode="browse")
        for col, head, w in zip(columns, headings, widths):
            tree.heading(col, text=head, anchor="center")
            tree.column(col, width=w, anchor="center")
        tree.tag_configure("odd", background=ODD)
        tree.tag_configure("bad", foreground=RED)
        tree.tag_configure("good", foreground=GREEN)
        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="right", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        return tree

    # ---------- الرئيسية ----------
    def month_start(self):
        return date.today().replace(day=1).isoformat()

    STAT_ICONS = {"visits": "visits_today", "maint": "maint_month", "notes": "notes_month", "errors": "errors_month"}

    def stat_card(self, parent, label, value, icon, c1, c2, qkey=None):
        H = 112
        cv = tk.Canvas(parent, height=H, highlightthickness=0, bd=0, bg=BG)
        hover = {"on": False}

        def draw(_=None):
            w = max(cv.winfo_width(), 120)
            paint_gradient(cv, w, H, c1, c2)
            cv.delete("txt")
            cv.create_text(w - 20, 34, text=str(value), anchor="e", fill="white", font=(FONT, 30, "bold"), tags="txt")
            cv.create_text(w - 20, 70, text=label, anchor="e", fill="#f2f6ff", font=(FONT, 10, "bold"), tags="txt")
            ic = load_icon(self.STAT_ICONS.get(qkey, ""), 54) if qkey else None
            if ic:
                cv.create_image(46, 40, image=ic, tags="txt")
            else:
                cv.create_text(46, 40, text=icon, fill="white", font=(EMOJI_FONT, 28), tags="txt")
            if qkey:
                cv.create_text(w - 20, 95, text="اضغط للتفاصيل ›", anchor="e", fill="#e3efff",
                               font=(FONT, 8, "bold"), tags="txt")
            if hover["on"]:
                cv.create_rectangle(2, 2, w - 2, H - 2, outline="white", width=2, tags="txt")
        cv.bind("<Configure>", draw)
        if qkey:
            cv.config(cursor="hand2")
            cv.bind("<Enter>", lambda e: (hover.update(on=True), draw()))
            cv.bind("<Leave>", lambda e: (hover.update(on=False), draw()))
            cv.bind("<Button-1>", lambda e: self.quick_info(qkey))
        return cv

    def quick_info(self, kind):
        """نافذة بيانات سريعة تفتح عند الضغط على بطاقة من بطاقات الصفحة الرئيسية."""
        today, ms = date.today().isoformat(), self.month_start()
        n_sql = ("SELECT n.employee, COALESCE(e.department,'-'), {sev}n.details, n.action, n.created_at "
                 "FROM employee_notes n LEFT JOIN employees e ON e.name=n.employee "
                 "WHERE n.kind=? AND date(n.created_at)>=? ORDER BY n.created_at DESC LIMIT 200")
        specs = {
            "visits": ("زيارات اليوم", "visits",
                       "SELECT person, entity, reason, created_at FROM visits WHERE date(created_at)=? ORDER BY created_at DESC LIMIT 200",
                       (today,), ("اسم الشخص", "الجهة", "سبب الزيارة", "تاريخ ووقت الزيارة"), (170, 150, 250, 165), "اليوم: "),
            "maint": ("صيانة هذا الشهر", "maintenance",
                      "SELECT person, entity, notes, created_at FROM maintenance WHERE date(created_at)>=? ORDER BY created_at DESC LIMIT 200",
                      (ms,), ("الفني / المسؤول", "الجهة / الشركة", "ملاحظات الصيانة", "تاريخ ووقت التسجيل"), (170, 150, 250, 165),
                      "من بداية الشهر: "),
            "notes": ("ملاحظات العاملين هذا الشهر", "employees", n_sql.format(sev=""), ("ملاحظة", ms),
                      ("اسم العامل", "القسم", "التفاصيل", "الإجراء المتخذ", "التاريخ والوقت"), (140, 80, 240, 130, 160),
                      "من بداية الشهر: "),
            "errors": ("أخطاء العاملين هذا الشهر", "employees", n_sql.format(sev="n.severity, "), ("خطأ", ms),
                       ("اسم العامل", "القسم", "الدرجة", "التفاصيل", "الإجراء المتخذ", "التاريخ والوقت"),
                       (130, 70, 70, 210, 120, 160), "من بداية الشهر: "),
        }
        title, section, sql, args, heads, widths, when = specs[kind]
        with db() as con:
            rows = con.execute(sql, args).fetchall()
        win = self._dialog("معلومات سريعة - " + title, "860x560")
        self._minsize(win, 640, 420)
        bx = tk.Frame(win, bg=BG); bx.pack(side="bottom", fill="x", pady=10)
        inner = tk.Frame(bx, bg=BG); inner.pack()
        if self.can(section, "view"):
            self.btn(inner, "فتح القسم كاملًا", lambda: (win.destroy(), self.go(section)), BLUE).pack(side="right", padx=6)
        self.btn(inner, "إغلاق", win.destroy, MUTED).pack(side="right", padx=6)
        tk.Label(win, text=title, bg=BG, fg=INK, font=(FONT, 16, "bold")).pack(anchor="e", padx=22, pady=(14, 0))
        info = f"{when}{date.today().strftime('%d/%m/%Y')}   |   العدد: {len(rows)}"
        if kind in ("notes", "errors") and rows:
            cnt = {}
            for r in rows:
                cnt[r[0]] = cnt.get(r[0], 0) + 1
            top = max(cnt.items(), key=lambda kv: kv[1])
            info += f"   |   أكثر العاملين {'أخطاءً' if kind == 'errors' else 'ملاحظات'}: {top[0]} ({top[1]})"
        tk.Label(win, text=info, bg=BG, fg=MUTED, font=(FONT, 10)).pack(anchor="e", padx=24, pady=(2, 4))
        if not rows:
            tk.Label(win, text="لا توجد سجلات", bg=BG, fg=MUTED, font=(FONT, 12, "bold")).pack(pady=30)
            return
        tree = self.make_tree(win, [f"c{i}" for i in range(len(heads))], heads, widths, 14)
        for n, r in enumerate(rows):
            vals = list(r)
            vals[-1] = fmt_stamp(vals[-1])
            if kind in ("notes", "errors"):
                vals[1] = DEPT_LABEL.get(vals[1], vals[1])
            tree.insert("", "end", values=[v if v not in (None, "") else "-" for v in vals],
                        tags=["odd"] if n % 2 else [])

    def mod_card(self, parent, key, icon, title, desc, color):
        card = tk.Frame(parent, bg=WHITE, highlightbackground=LINE, highlightthickness=1, cursor="hand2")
        ic = tk.Canvas(card, width=46, height=46, bg=WHITE, highlightthickness=0)
        ic.create_oval(2, 2, 44, 44, fill=color, outline="")
        img = load_icon(key, 34)
        if img:
            ic.create_image(23, 24, image=img)
        else:
            ic.create_text(23, 24, text=icon, fill="white", font=(EMOJI_FONT, 15))
        ic.pack(side="right", padx=10, pady=12, anchor="n")
        txt = tk.Frame(card, bg=WHITE); txt.pack(side="right", fill="both", expand=True, pady=10)
        tk.Label(txt, text=title, bg=WHITE, fg=INK, font=(FONT, 11, "bold"), anchor="e", justify="right",
                 wraplength=150).pack(fill="x")
        tk.Label(txt, text=desc, bg=WHITE, fg=MUTED, font=(FONT, 9), anchor="e", justify="right",
                 wraplength=150).pack(fill="x")
        tk.Label(txt, text="فتح القسم  ←", bg=WHITE, fg=color, font=(FONT, 9, "bold"), anchor="e").pack(fill="x", pady=(4, 0))

        def on(_=None): card.config(highlightbackground=color, highlightthickness=2)
        def off(_=None): card.config(highlightbackground=LINE, highlightthickness=1)

        def wire(w):
            w.bind("<Enter>", on); w.bind("<Leave>", off); w.bind("<Button-1>", lambda e: self.go(key))
            for c in w.winfo_children():
                wire(c)
        wire(card)
        return card

    def insight_box(self, parent, title, lines):
        box = tk.Frame(parent, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        tk.Frame(box, bg=PURPLE, width=6).pack(side="right", fill="y")
        inner = tk.Frame(box, bg=WHITE); inner.pack(side="right", fill="x", expand=True, padx=12, pady=8)
        tk.Label(inner, text=title, bg=WHITE, fg=INK, font=(FONT, 11, "bold"), anchor="e").pack(fill="x")
        for t in lines:
            tk.Label(inner, text=t, bg=WHITE, fg=TEXT, font=(FONT, 10), anchor="e", justify="right").pack(fill="x", pady=1)
        return box

    def show_home(self):
        self.page_title("الصفحة الرئيسية", f"مرحبًا بك في {APP_NAME} - اختر القسم المطلوب")
        today, ms = date.today().isoformat(), self.month_start()
        with db() as con:
            def one(sql, *a):
                try:
                    return con.execute(sql, a).fetchone()[0]
                except Exception:
                    return 0        # المستخدم المحدود مايشوفش كل الجداول
            stats = [
                ("زيارات اليوم", one("SELECT COUNT(*) FROM visits WHERE date(created_at)=?", today), "👥", "#2563eb", "#38bdf8", "visits"),
                ("صيانة هذا الشهر", one("SELECT COUNT(*) FROM maintenance WHERE date(created_at)>=?", ms), "🔧", "#f97316", "#fbbf24", "maint"),
                ("ملاحظات العاملين هذا الشهر", one("SELECT COUNT(*) FROM employee_notes WHERE kind='ملاحظة' AND date(created_at)>=?", ms), "📝", "#7c3aed", "#c084fc", "notes"),
                ("أخطاء العاملين هذا الشهر", one("SELECT COUNT(*) FROM employee_notes WHERE kind='خطأ' AND date(created_at)>=?", ms), "⚠️", "#e11d48", "#fb7185", "errors"),
            ]
        stat_pages = ["visits", "maintenance", "employees", "employees"]
        stats = [st for st, pg in zip(stats, stat_pages) if self.can(pg, "view")]
        row = tk.Frame(self.content, bg=BG)
        row.pack(fill="x", padx=24, pady=(0, 6))
        for i, st in enumerate(stats):
            self.stat_card(row, *st).grid(row=0, column=i, sticky="nsew", padx=7, pady=4)
            row.grid_columnconfigure(i, weight=1, uniform="st")

        MC = {"workers": TEAL, "visits": BLUE, "maintenance": ORANGE, "employees": PURPLE, "cleaning": CYAN,
              "paper": AMBER, "toner": PINK, "purchases": "#16a34a", "attend": "#0ea5e9", "gov": "#be123c", "govstaff": "#0f766e", "compare": "#6366f1", "reports": GREEN}
        cards = [
            ("workers", "👷", "قاعدة بيانات العاملين", "بيانات كاملة بالصور وبداية العمل والحالة (أمن - خدمات - صيانة)"),
            ("visits", "👥", "زيارات المركز", "تسجيل الزائر والجهة وسبب الزيارة بالتاريخ والوقت"),
            ("maintenance", "🔧", "الصيانة", "أعمال الصيانة وملاحظاتها وتاريخ تنفيذها"),
            ("employees", "📝", "ملاحظات وأخطاء العاملين", "ملاحظات وأخطاء وتنبيهات وإشادات كل عامل"),
            ("cleaning", "🧴", "جرد خامات النظافة", "رصيد الخامات وجرد كل أسبوع"),
            ("paper", "📄", "جرد الأدوات الورقية", "أرصدة الورق والأدوات المكتبية"),
            ("toner", "🖨️", "جرد الأحبار", "الأحبار والتونرات حسب الطابعة"),
            ("purchases", "🛒", "المشتريات", "ورقيات وأحبار وأدوات - إضافة يدوية مع صورة الفاتورة"),
            ("attend", "🕒", "حضور وانصراف العاملين", "حضور وانصراف يدوي + إذنان شهريًا + الغياب والتأخير"),
            ("govstaff", "🗂️", "قاعدة بيانات موظفي الجهات", "موظفو الجهات الثماني: البيانات والوظيفة والهاتف والحالة"),
            ("gov", "🏛️", "حضور وانصراف الجهات", "اختر الجهة ثم الموظف وسجّل دخول أو خروج أو ملاحظة"),
            ("compare", "📈", "مقارنة الفترات", "قارن شهرين أو أسبوعين مع رسم بياني وملخص ذكي"),
            ("reports", "📊", "استخراج وطباعة التقارير", "Excel أو طباعة جاهزة لأي فترة"),
        ]
        cards = [c for c in cards if self.can(c[0], "view")]
        _order = ["workers", "govstaff", "attend", "gov", "visits", "maintenance", "cleaning", "paper", "toner",
                  "purchases", "employees", "compare", "reports"]
        cards.sort(key=lambda c: _order.index(c[0]) if c[0] in _order else 99)
        grid = tk.Frame(self.content, bg=BG)
        grid.pack(fill="both", expand=True, padx=24, pady=4)
        for i, (key, icon, title, desc) in enumerate(cards):
            r, c = divmod(i, 4)
            self.mod_card(grid, key, icon, title, desc, MC[key]).grid(row=r, column=3 - c, sticky="nsew", padx=7, pady=6)
            grid.grid_columnconfigure(c, weight=1, uniform="md")
            grid.grid_rowconfigure(r, weight=1)
        size, total, pmb = db_info()
        tk.Label(self.content, text=f"قاعدة البيانات: {size:.2f} ميجابايت   |   إجمالي السجلات: {total:,}   |   الصور: {pmb:.1f} ميجابايت   |   نسخة احتياطية يومية تلقائية في مجلد backups",
                 bg=BG, fg=MUTED, font=(FONT, 9)).pack(pady=(0, 6))

    # ---------- الزيارات ----------
    def show_visits(self):
        self.page_title("زيارات المركز", "أدخل بيانات الزائر، وسيتم تسجيل التاريخ والوقت تلقائيًا")
        form = self.panel(); form.pack(fill="x", padx=24, pady=6)
        self.visit_person = tk.StringVar(); self.visit_entity = tk.StringVar(); self.visit_reason = tk.StringVar()
        self.add_field(form, "اسم الشخص", self.visit_person, 0, 0, 250)
        self.add_field(form, "الجهة", self.visit_entity, 0, 1, 250)
        self.add_field(form, "سبب الزيارة", self.visit_reason, 1, 0, 560)
        bx = tk.Frame(form, bg=WHITE); bx.grid(row=1, column=1, padx=14, pady=10, sticky="e")
        self.btn(bx, "حفظ الزيارة", self.save_visit).pack(side="right", padx=4)
        self.btn(bx, "✏️ تعديل المحدد", lambda: self.edit_record("visits", self.visits_tree), BLUE).pack(side="right", padx=4)
        self.btn(bx, "حذف المحدد", lambda: self.delete_selected("visits", self.visits_tree), RED).pack(side="right", padx=4)
        self.btn(bx, "🖨️ طباعة", lambda: self.print_table("visits"), AMBER).pack(side="right", padx=4)
        self.btn(bx, "📊 استخراج Excel", lambda: self.export_table("visits"), GREEN).pack(side="right", padx=4)
        self.visits_tree = self.make_tree(self.content, ("person", "entity", "reason", "time"),
                                          ("اسم الشخص", "الجهة", "سبب الزيارة", "تاريخ ووقت الزيارة"), (190, 180, 330, 190))
        self.visits_tree.bind("<Double-1>", lambda e: self.edit_record("visits", self.visits_tree))
        self.load_recent("visits", self.visits_tree)

    def save_visit(self):
        person = self.visit_person.get().strip()
        if not person:
            messagebox.showwarning("بيانات ناقصة", "اكتب اسم الشخص أولًا."); return
        with db() as con:
            con.execute("INSERT INTO visits(person,entity,reason,created_at) VALUES(?,?,?,?)",
                        (person, self.visit_entity.get().strip(), self.visit_reason.get().strip(), now_stamp()))
        for v in (self.visit_person, self.visit_entity, self.visit_reason): v.set("")
        self.load_recent("visits", self.visits_tree)
        messagebox.showinfo("تم الحفظ", "تم تسجيل الزيارة بالتاريخ والوقت.")

    # ---------- الملاحظات وتعليمات الوحدات ----------
    def _station_name(self):
        try:
            return ((NET_CFG.get("station") if NET_CFG else "") or NET_STATION or "").strip()
        except Exception:
            return ""

    def _note_for_me(self, target):
        tgt = (target or "").strip()
        if tgt in ("", NOTE_TARGET_ALL, "الكل", "الجميع"):
            return True
        st = self._station_name().lower()
        return bool(st) and tgt.lower() == st

    def _pending_notes(self, mine=True):
        try:
            with db() as con:
                rows = con.execute(
                    "SELECT id, body, target, priority, created_by, created_at "
                    "FROM notes WHERE status='new' ORDER BY id DESC LIMIT 200").fetchall()
        except Exception:
            return []
        if not mine:
            return rows
        return [r for r in rows if self._note_for_me(r[2])]

    def build_notes_banner(self):
        """شريط أعلى الشاشة: الملاحظات الجديدة الموجهة لهذه الوحدة مع زر «قدمها»."""
        old = getattr(self, "_notes_banner", None)
        if old is not None:
            try: old.destroy()
            except Exception: pass
            self._notes_banner = None
        pending = self._pending_notes(True)
        # ---- تنبيه بصوت عند وصول ملاحظة جديدة (مش أول تحميل، ومش عند التنقّل) ----
        ids = {r[0] for r in pending}
        fresh = ids - getattr(self, "_notes_seen", set())
        first = not getattr(self, "_notes_init", False)
        self._notes_seen = ids
        self._notes_init = True
        if fresh and not first:
            urgent = any((r[3] or "") == "عاجلة" for r in pending if r[0] in fresh)
            self._play_note_sound(urgent)
        if not pending:
            return
        kids = self.content.winfo_children()
        bar = tk.Frame(self.content, bg=AMBER)
        if kids:
            bar.pack(fill="x", padx=24, pady=(10, 2), before=kids[0])
        else:
            bar.pack(fill="x", padx=24, pady=(10, 2))
        head = tk.Frame(bar, bg=AMBER); head.pack(fill="x", padx=12, pady=(6, 2))
        tk.Label(head, text=f"📢  ملاحظات وتعليمات جديدة ({len(pending)})", bg=AMBER, fg=NAVY,
                 font=(FONT, 12, "bold"), anchor="e").pack(side="right")
        for r in pending:
            row = tk.Frame(bar, bg=AMBER); row.pack(fill="x", padx=12, pady=3)
            tk.Button(row, text="✅ قدمها", command=lambda i=r[0]: self.mark_note_done(i),
                      bg=GREEN, fg="white", activebackground=NAVY, activeforeground="white",
                      relief="flat", bd=0, font=(FONT, 9, "bold"), padx=12, pady=3,
                      cursor="hand2").pack(side="left", padx=(0, 10))
            txt = (r[1] or "")
            if (r[3] or "") == "عاجلة":
                txt = "⚠️ " + txt
            tk.Label(row, text=txt, bg=AMBER, fg=INK, font=(FONT, 11, "bold"),
                     anchor="e", justify="right", wraplength=980).pack(side="right", fill="x", expand=True)
        tk.Frame(bar, bg=AMBER, height=6).pack()
        self._notes_banner = bar

    def _sound_on(self):
        """هل التنبيه الصوتي للملاحظات مفعّل؟ (محفوظ في settings.json)."""
        try:
            return bool(load_settings().get("notes_sound", True))
        except Exception:
            return True

    def _beep(self, urgent=False):
        """صوت تنبيه: على ويندوز صوت النظام، وعلى لينكس جرس Tk."""
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONHAND if urgent else winsound.MB_ICONEXCLAMATION)
            return
        except Exception:
            pass
        try:
            self.root.bell()
        except Exception:
            pass

    def _play_note_sound(self, urgent=False):
        """يشتغل لما توصل ملاحظة جديدة (العاجلة بجرسين)."""
        if not self._sound_on():
            return
        self._beep(urgent)
        if urgent:
            try: self.root.after(450, lambda: self._beep(True))
            except Exception: pass

    def toggle_note_sound(self):
        cfg = load_settings()
        cfg["notes_sound"] = bool(self.note_sound.get())
        save_settings(cfg)

    def mark_note_done(self, note_id):
        """يعلّم الملاحظة كـ «تم» (من أي وحدة)."""
        try:
            if netmod and NET_MODE == netmod.MODE_CLIENT:
                res = netmod.mark_note_done(NET_CFG, note_id)
                if not res.get("ok"):
                    messagebox.showerror("خطأ", res.get("error") or "تعذر تحديث الملاحظة")
                    return
            else:
                with db() as con:
                    con.execute("UPDATE notes SET status='done', done_by=?, done_at=? WHERE id=?",
                                (self.user["username"], now_stamp(), int(note_id)))
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تحديث الملاحظة:\n{ex}"); return
        if getattr(self, "active_key", None) == "notes" and getattr(self, "notes_tree", None):
            self.load_notes()
        self.build_notes_banner()

    def note_targets(self):
        vals = [NOTE_TARGET_ALL]
        st = self._station_name()
        if st and st not in vals:
            vals.append(st)
        try:
            if NET_SERVER and getattr(NET_SERVER, "stations", None):
                for s in NET_SERVER.stations:
                    if s and s not in vals:
                        vals.append(s)
        except Exception:
            pass
        for s in NOTE_DEFAULT_TARGETS:
            if s not in vals:
                vals.append(s)
        return vals

    def show_notes(self):
        self.page_title("📢 ملاحظات وتعليمات الوحدات",
                        "اكتب ملاحظة وحدد الوحدة — تظهر عندهم فورًا في شريط أعلى الشاشة مع زر «قدمها»، وتتحول هنا إلى «تم»")
        lv = self.level("notes")
        if lv in ("add", "edit"):
            form = self.panel(); form.pack(fill="x", padx=24, pady=6)
            self.note_body = tk.StringVar()
            self.note_target = tk.StringVar(value=NOTE_TARGET_ALL)
            self.note_prio = tk.StringVar(value=NOTE_PRIORITIES[0])
            self.add_combo(form, "الوحدة المستهدفة", self.note_target, self.note_targets(), 0, 0, readonly=False)
            self.add_combo(form, "الأولوية", self.note_prio, NOTE_PRIORITIES, 0, 1)
            self.add_field(form, "نص الملاحظة", self.note_body, 1, 0, 620)
            bx = tk.Frame(form, bg=WHITE); bx.grid(row=1, column=1, padx=14, pady=10, sticky="e")
            self.btn(bx, "📢 إرسال الملاحظة", self.save_note).pack(side="right", padx=4)
            self.btn(bx, "حذف المحدد", self.delete_note, RED).pack(side="right", padx=4)
        sbar = tk.Frame(self.content, bg=BG); sbar.pack(fill="x", padx=24, pady=(4, 0))
        self.note_filter = tk.StringVar(value="الكل")
        tk.Label(sbar, text="عرض:", bg=BG, fg=TEXT, font=(FONT, 10, "bold")).pack(side="right", padx=6)
        cb = ttk.Combobox(sbar, textvariable=self.note_filter, values=["الكل", "جديد", "تم"],
                          state="readonly", width=10, justify="right")
        cb.pack(side="right")
        cb.bind("<<ComboboxSelected>>", lambda e: self.load_notes())
        self.btn(sbar, "✅ قدّم المحدد", self.mark_note_done_selected, GREEN).pack(side="left", padx=6)
        self.btn(sbar, "🔄 تحديث", self.load_notes, MUTED).pack(side="left", padx=6)
        self.note_sound = tk.BooleanVar(value=self._sound_on())
        tk.Checkbutton(sbar, text="🔔 تنبيه بصوت عند وصول ملاحظة جديدة", variable=self.note_sound,
                       command=self.toggle_note_sound, bg=BG, fg=TEXT, activebackground=BG,
                       activeforeground=TEXT, selectcolor=WHITE, font=(FONT, 10, "bold"),
                       cursor="hand2").pack(side="left", padx=10)
        self.notes_tree = self.make_tree(self.content, ("body", "target", "prio", "status", "by", "time"),
                                         ("نص الملاحظة", "الوحدة", "الأولوية", "الحالة", "قدّمها", "التاريخ والوقت"),
                                         (400, 120, 80, 90, 120, 150), 11)
        self.load_notes()

    def load_notes(self):
        if not getattr(self, "notes_tree", None):
            return
        try:
            with db() as con:
                rows = con.execute(
                    "SELECT id, body, target, priority, created_by, created_at, status, done_by, done_at "
                    "FROM notes ORDER BY id DESC LIMIT 500").fetchall()
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تحميل الملاحظات:\n{ex}"); return
        flt = self.note_filter.get() if getattr(self, "note_filter", None) else "الكل"
        self.notes_tree.delete(*self.notes_tree.get_children())
        for n, r in enumerate(rows):
            rid, body, target, prio, by, at, status, dby, dat = r
            if flt == "جديد" and status == "done":
                continue
            if flt == "تم" and status != "done":
                continue
            done = (status == "done")
            tags = ["odd"] if n % 2 else []
            tags.append("good" if done else "bad")
            self.notes_tree.insert("", "end", iid=str(rid), tags=tags, values=(
                body or "", target or NOTE_TARGET_ALL, prio or NOTE_PRIORITIES[0],
                "تم ✅" if done else "جديد",
                (dby or "") if done else "",
                fmt_stamp(dat) if done else fmt_stamp(at)))

    def save_note(self):
        body = self.note_body.get().strip()
        if not body:
            messagebox.showwarning("بيانات ناقصة", "اكتب نص الملاحظة أولًا."); return
        target = self.note_target.get().strip()
        if target == NOTE_TARGET_ALL:
            target = ""
        with db() as con:
            con.execute("INSERT INTO notes(body,target,priority,created_by,created_at,status) VALUES(?,?,?,?,?,'new')",
                        (body, target, self.note_prio.get(), self.user["username"], now_stamp()))
        self.note_body.set("")
        self.load_notes()
        self.build_notes_banner()
        messagebox.showinfo("تم الإرسال", "تم إرسال الملاحظة للوحدة المحددة.\nهتظهر عندهم فورًا مع زر «قدمها».")

    def mark_note_done_selected(self):
        sel = self.notes_tree.selection() if getattr(self, "notes_tree", None) else ()
        if not sel:
            messagebox.showinfo("تقديم الملاحظة", "اختر ملاحظة من الجدول أولًا."); return
        self.mark_note_done(int(sel[0]))

    def delete_note(self):
        sel = self.notes_tree.selection() if getattr(self, "notes_tree", None) else ()
        if not sel:
            messagebox.showinfo("حذف الملاحظة", "اختر ملاحظة من الجدول أولًا."); return
        if not messagebox.askyesno("حذف الملاحظة", "حذف الملاحظة المحددة نهائيًا؟", icon="warning"):
            return
        with db() as con:
            con.execute("DELETE FROM notes WHERE id=?", (int(sel[0]),))
        self.load_notes()
        self.build_notes_banner()

    # ---------- الصيانة ----------
    def show_maintenance(self):
        self.page_title("سجل الصيانة", "سجل أعمال الصيانة مع الجهة أو الفني والملاحظات، والتاريخ والوقت تلقائيًا")
        form = self.panel(); form.pack(fill="x", padx=24, pady=6)
        self.maint_person = tk.StringVar(); self.maint_entity = tk.StringVar(); self.maint_notes = tk.StringVar()
        self.add_field(form, "اسم الفني / المسؤول", self.maint_person, 0, 0, 250)
        self.add_field(form, "الجهة / الشركة", self.maint_entity, 0, 1, 250)
        self.add_field(form, "ملاحظات الصيانة", self.maint_notes, 1, 0, 560)
        bx = tk.Frame(form, bg=WHITE); bx.grid(row=1, column=1, padx=14, pady=10, sticky="e")
        self.btn(bx, "حفظ عملية الصيانة", self.save_maintenance).pack(side="right", padx=4)
        self.btn(bx, "✏️ تعديل المحدد", lambda: self.edit_record("maintenance", self.maint_tree), BLUE).pack(side="right", padx=4)
        self.btn(bx, "حذف المحدد", lambda: self.delete_selected("maintenance", self.maint_tree), RED).pack(side="right", padx=4)
        self.btn(bx, "🖨️ طباعة", lambda: self.print_table("maintenance"), AMBER).pack(side="right", padx=4)
        self.btn(bx, "📊 استخراج Excel", lambda: self.export_table("maintenance"), GREEN).pack(side="right", padx=4)
        self.maint_tree = self.make_tree(self.content, ("person", "entity", "notes", "time"),
                                         ("الفني / المسؤول", "الجهة / الشركة", "ملاحظات الصيانة", "تاريخ ووقت التسجيل"), (180, 180, 350, 190))
        self.maint_tree.bind("<Double-1>", lambda e: self.edit_record("maintenance", self.maint_tree))
        self.load_recent("maintenance", self.maint_tree)

    def save_maintenance(self):
        notes = self.maint_notes.get().strip()
        if not notes:
            messagebox.showwarning("بيانات ناقصة", "اكتب ملاحظات الصيانة أولًا."); return
        with db() as con:
            con.execute("INSERT INTO maintenance(person,entity,notes,created_at) VALUES(?,?,?,?)",
                        (self.maint_person.get().strip(), self.maint_entity.get().strip(), notes, now_stamp()))
        for v in (self.maint_person, self.maint_entity, self.maint_notes): v.set("")
        self.load_recent("maintenance", self.maint_tree)
        messagebox.showinfo("تم الحفظ", "تم تسجيل عملية الصيانة بالتاريخ والوقت.")

    # ---------- ملاحظات وأخطاء العاملين ----------
    def show_employees(self):
        self.page_title("📝 ملاحظات وأخطاء العاملين",
                        "سجّل الملاحظة أو الخطأ أو التنبيه أو الإشادة لكل عامل، ثم استخرج تقريرًا بالفترة من صفحة التقارير")
        form = self.panel(); form.pack(fill="x", padx=24, pady=6)
        self.emp_name = tk.StringVar(); self.emp_kind = tk.StringVar(value=NOTE_KINDS[0])
        self.emp_sev = tk.StringVar(value=SEVERITIES[0]); self.emp_action = tk.StringVar()
        self.add_combo(form, "اسم العامل (اختر أو اكتب)", self.emp_name, self.worker_names(), 0, 0, readonly=False)
        self.add_combo(form, "النوع", self.emp_kind, NOTE_KINDS, 0, 1)
        self.add_combo(form, "الدرجة", self.emp_sev, SEVERITIES, 0, 2)
        box = tk.Frame(form, bg=WHITE); box.grid(row=1, column=0, columnspan=2, padx=14, pady=8, sticky="e")
        tk.Label(box, text="تفاصيل الملاحظة / الخطأ", bg=WHITE, fg=TEXT, font=(FONT, 10, "bold"), anchor="e").pack(fill="x", pady=(0, 4))
        self.emp_details = tk.Text(box, height=3, width=62, font=(FONT, 10), relief="solid", bd=1, wrap="word",
                                   bg=ENTRY, fg=TEXT, insertbackground=TEXT)
        self.emp_details.tag_configure("r", justify="right")
        self.emp_details.bind("<KeyRelease>", lambda e: self.emp_details.tag_add("r", "1.0", "end"))
        self.emp_details.pack()
        self.add_field(form, "الإجراء المتخذ (اختياري)", self.emp_action, 1, 2, 250)
        bx = tk.Frame(form, bg=WHITE); bx.grid(row=2, column=0, columnspan=3, padx=14, pady=(2, 12), sticky="e")
        self.btn(bx, "حفظ", self.save_employee_note).pack(side="right", padx=4)
        self.btn(bx, "✏️ تعديل المحدد", lambda: self.edit_record("employee_notes", self.emp_tree), BLUE).pack(side="right", padx=4)
        self.btn(bx, "حذف المحدد", lambda: self.delete_selected("employee_notes", self.emp_tree), RED).pack(side="right", padx=4)
        self.btn(bx, "🖨️ طباعة", lambda: self.print_table("employee_notes"), AMBER).pack(side="right", padx=4)
        self.btn(bx, "📊 استخراج Excel", lambda: self.export_table("employee_notes"), GREEN).pack(side="right", padx=4)

        sbar = tk.Frame(self.content, bg=BG); sbar.pack(fill="x", padx=24)
        self.emp_search = tk.StringVar()
        tk.Label(sbar, text="بحث باسم العامل:", bg=BG, fg=TEXT, font=(FONT, 10, "bold")).pack(side="right", padx=6)
        e = ttk.Entry(sbar, textvariable=self.emp_search, width=26, justify="right"); e.pack(side="right")
        e.bind("<KeyRelease>", lambda ev: self.load_recent("employee_notes", self.emp_tree, self.emp_search.get().strip()))
        self.emp_tree = self.make_tree(self.content, ("emp", "dept", "kind", "sev", "details", "action", "time"),
                                       ("اسم العامل", "القسم", "النوع", "الدرجة", "التفاصيل", "الإجراء المتخذ", "التاريخ والوقت"),
                                       (130, 70, 85, 70, 250, 150, 150), 8)
        self.emp_tree.bind("<Double-1>", lambda e: self.edit_record("employee_notes", self.emp_tree))
        self.load_recent("employee_notes", self.emp_tree)

    def save_employee_note(self):
        name = self.emp_name.get().strip()
        details = self.emp_details.get("1.0", "end").strip()
        if not name or not details:
            messagebox.showwarning("بيانات ناقصة", "اكتب اسم العامل وتفاصيل الملاحظة أو الخطأ."); return
        sev = self.emp_sev.get() if self.emp_kind.get() in ("خطأ", "تنبيه شفهي") else ""
        with db() as con:
            con.execute("""INSERT INTO employee_notes(employee,kind,severity,details,action,created_at)
                VALUES(?,?,?,?,?,?)""", (name, self.emp_kind.get(), sev, details, self.emp_action.get().strip(), now_stamp()))
        self.emp_name.set(""); self.emp_action.set(""); self.emp_details.delete("1.0", "end")
        self.load_recent("employee_notes", self.emp_tree, self.emp_search.get().strip())
        messagebox.showinfo("تم الحفظ", "تم تسجيل البيان بالتاريخ والوقت.")

    # ---------- قاعدة بيانات العاملين ----------
    def worker_names(self):
        with db() as con:
            return [r[0] for r in con.execute("SELECT name FROM employees WHERE status<>? ORDER BY name", (STATUS_OFF,))]

    def show_workers(self):
        self.page_title("👷 قاعدة بيانات العاملين",
                        "بيانات كل عامل كاملة: القسم والصور وبداية العمل وحالته (على رأس العمل / ترك العمل) - اضغط مرتين على العامل لعرض بياناته وتعديلها")
        bar = tk.Frame(self.content, bg=BG); bar.pack(fill="x", padx=24, pady=(0, 6))
        self.btn(bar, "➕ إضافة عامل", lambda: self.worker_dialog()).pack(side="right", padx=4)
        self.btn(bar, "✏️ عرض / تعديل المحدد", self.edit_selected_worker, BLUE).pack(side="right", padx=4)
        self.btn(bar, "حذف المحدد", self.delete_worker, RED).pack(side="right", padx=4)
        self.btn(bar, "🖨️ طباعة", self.print_workers, AMBER).pack(side="right", padx=4)
        self.btn(bar, "📊 استخراج Excel", self.export_workers, GREEN).pack(side="right", padx=4)

        fbar = tk.Frame(self.content, bg=BG); fbar.pack(fill="x", padx=24, pady=2)
        self.w_filter = tk.StringVar(value="الكل"); self.w_sfilter = tk.StringVar(value="الكل"); self.w_search = tk.StringVar()
        tk.Label(fbar, text="القسم:", bg=BG, fg=TEXT, font=(FONT, 10, "bold")).pack(side="right", padx=6)
        cb = ttk.Combobox(fbar, textvariable=self.w_filter, values=["الكل"] + DEPARTMENTS,
                          state="readonly", width=10, justify="right")
        cb.pack(side="right"); cb.bind("<<ComboboxSelected>>", lambda e: self.load_workers())
        tk.Label(fbar, text="الحالة:", bg=BG, fg=TEXT, font=(FONT, 10, "bold")).pack(side="right", padx=(14, 6))
        cb2 = ttk.Combobox(fbar, textvariable=self.w_sfilter, values=["الكل", STATUS_ON, STATUS_OFF],
                           state="readonly", width=14, justify="right")
        cb2.pack(side="right"); cb2.bind("<<ComboboxSelected>>", lambda e: self.load_workers())
        tk.Label(fbar, text="بحث (اسم / هاتف / رقم قومي):", bg=BG, fg=TEXT, font=(FONT, 10, "bold")).pack(side="right", padx=(14, 6))
        en = ttk.Entry(fbar, textvariable=self.w_search, width=22, justify="right")
        en.pack(side="right"); en.bind("<KeyRelease>", lambda e: self.load_workers())
        self.w_count = tk.Label(self.content, text="", bg=BG, fg=INK, font=(FONT, 10, "bold"), anchor="e")
        self.w_count.pack(fill="x", padx=26, pady=(4, 0))
        self.w_tree = self.make_tree(self.content, ("name", "dept", "title", "status", "start", "end", "phone"),
                                     ("اسم العامل", "القسم", "الوظيفة", "الحالة", "بدء العمل", "انتهاء العمل", "الهاتف"),
                                     (190, 70, 120, 110, 100, 100, 110), 12)
        self.w_tree.tag_configure("off", foreground=MUTED)
        self.w_tree.bind("<Double-1>", lambda e: self.edit_selected_worker())
        self.load_workers()

    def load_workers(self):
        sql, args = ("SELECT id,name,department,job_title,status,start_date,end_date,phone "
                     "FROM employees WHERE 1=1"), []
        if self.w_filter.get() != "الكل":
            sql += " AND department=?"; args.append(self.w_filter.get())
        if self.w_sfilter.get() != "الكل":
            sql += " AND status=?"; args.append(self.w_sfilter.get())
        key = self.w_search.get().strip()
        if key:
            sql += " AND (name LIKE ? OR phone LIKE ? OR national_id LIKE ?)"; args += [f"%{key}%"] * 3
        with db() as con:
            rows = con.execute(sql + " ORDER BY (status=?), department, name", args + [STATUS_OFF]).fetchall()
            stat = con.execute("SELECT department, status, COUNT(*) FROM employees GROUP BY department, status").fetchall()
        self.w_tree.delete(*self.w_tree.get_children())
        for n, (rid, name, dept, title, status, sd, ed, phone) in enumerate(rows):
            tags = ["odd"] if n % 2 else []
            if status == STATUS_OFF: tags.append("off")
            self.w_tree.insert("", "end", iid=str(rid), tags=tags,
                               values=(name, dept, title or "-", status, fmt_iso_date(sd) or "-", fmt_iso_date(ed) or "-", phone or "-"))
        on = sum(c for _, st, c in stat if st != STATUS_OFF); off = sum(c for _, st, c in stat if st == STATUS_OFF)
        per = "   |   ".join(f"{d}: {sum(c for dd, st, c in stat if dd == d and st != STATUS_OFF)}" for d in DEPARTMENTS)
        self.w_count.config(text=f"الإجمالي: {on + off}   |   على رأس العمل: {on}   |   تركوا العمل: {off}      (على رأس العمل بالقسم -> {per})")

    def _workers_secs(self):
        return workers_sections(self.w_filter.get(), self.w_sfilter.get(), self.w_search.get().strip())

    def print_workers(self):
        """يطبع قاعدة العاملين حسب الفلاتر الظاهرة (القسم / الحالة / البحث)."""
        try:
            open_html(sections_html("قاعدة بيانات العاملين", self._workers_secs()), "workers")
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تجهيز الطباعة:\n{ex}")

    def export_workers(self):
        try:
            secs = self._workers_secs()
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تجهيز الملف:\n{ex}"); return
        self.export_sections_excel(secs, f"قاعدة_العاملين_{date.today():%Y-%m-%d}.xlsx", "قاعدة العاملين")

    def edit_selected_worker(self):
        sel = self.w_tree.selection()
        if not sel:
            messagebox.showwarning("اختر عاملًا", "حدد العامل من الجدول أولًا."); return
        self.worker_dialog(int(sel[0]))

    def load_thumb(self, name, size):
        if not name:
            return None
        try:
            path = make_thumb(name, size)
            if path:
                return tk.PhotoImage(file=path)
            src = ensure_photo(name) or os.path.join(PHOTOS_DIR, name)
            if src.lower().endswith((".png", ".gif")):
                img = tk.PhotoImage(file=src)
                f = max(1, -(-max(img.width(), img.height()) // size))
                return img.subsample(f, f) if f > 1 else img
        except Exception:
            pass
        return None

    def worker_dialog(self, wid=None):
        if wid is None and not self.need("workers", "add"):
            return
        row = None
        if wid:
            with db() as con:
                row = con.execute("""SELECT name,department,job_title,national_id,phone,birth_date,address,
                    start_date,status,end_date,notes,photo_file,id_file FROM employees WHERE id=?""", (wid,)).fetchone()
        r = list(row) if row else ["", DEPARTMENTS[0], "", "", "", "", "", date.today().isoformat(), STATUS_ON, "", "", "", ""]
        old_name = r[0]
        win = tk.Toplevel(self.root)
        win.title("بيانات العامل" if wid else "إضافة عامل جديد")
        _sz = self._clamp_size("830x700", 520, 400)
        win.configure(bg=BG); win.geometry(_sz)
        try:
            win.transient(self.root)
            win.geometry(f"{_sz}+{self.root.winfo_rootx() + 120}+{max(0, self.root.winfo_rooty() + 20)}")
            win.wait_visibility(); win.grab_set()
        except Exception:
            pass
        keys = ["name", "dept", "title", "nid", "phone", "birth", "address", "start", "status", "end", "notes"]
        init = [r[0], r[1], r[2], r[3], r[4], fmt_iso_date(r[5]), r[6], fmt_iso_date(r[7]), r[8] or STATUS_ON, fmt_iso_date(r[9]), r[10]]
        v = {k: tk.StringVar(value=x or "") for k, x in zip(keys, init)}

        tk.Label(win, text="بيانات العامل" if wid else "إضافة عامل جديد", bg=BG, fg=INK,
                 font=(FONT, 16, "bold")).pack(anchor="e", padx=20, pady=(12, 4))
        frm = self.panel(win); frm.pack(fill="x", padx=16, pady=4)
        self.add_field(frm, "اسم العامل (ثلاثي)", v["name"], 0, 2, 260)
        self.add_combo(frm, "القسم", v["dept"], DEPARTMENTS, 0, 1)
        self.add_field(frm, "الوظيفة / المسمى", v["title"], 0, 0, 200)
        self.add_field(frm, "الرقم القومي (14 رقم)", v["nid"], 1, 2, 260)
        self.add_field(frm, "رقم الهاتف", v["phone"], 1, 1, 200)
        self.add_date(frm, "تاريخ الميلاد", v["birth"], 1, 0, optional=True)
        self.add_field(frm, "العنوان", v["address"], 2, 2, 260)
        self.add_field(frm, "ملاحظات", v["notes"], 2, 1, 200)
        self.add_date(frm, "تاريخ بدء العمل", v["start"], 3, 2, optional=True)
        cb = self.add_combo(frm, "الحالة", v["status"], [STATUS_ON, STATUS_OFF], 3, 1)
        self.add_date(frm, "تاريخ انتهاء العمل (إن ترك العمل)", v["end"], 3, 0, optional=True)

        def on_status(_=None):
            if v["status"].get() == STATUS_OFF and not v["end"].get().strip():
                v["end"].set(today_text())
            elif v["status"].get() == STATUS_ON:
                v["end"].set("")
        cb.bind("<<ComboboxSelected>>", on_status)

        photos = {"photo": r[11] or "", "id": r[12] or ""}
        labels, refs = {}, {}
        pframe = self.panel(win); pframe.pack(fill="x", padx=16, pady=6)

        def refresh(key):
            name, lbl = photos[key], labels[key]
            img = self.load_thumb(name, 150) if name else None
            if img:
                refs[key] = img; lbl.config(image=img, text="")
            else:
                refs.pop(key, None)
                lbl.config(image="", text=("تم حفظ الصورة\n(اضغط عرض للمشاهدة)" if name else "لا توجد صورة"))

        def choose(key):
            path = filedialog.askopenfilename(parent=win, title="اختر الصورة",
                                              filetypes=[("صور", "*.jpg *.jpeg *.png *.bmp *.gif"), ("كل الملفات", "*.*")])
            if not path:
                return
            try:
                photos[key] = save_photo(path, key)
            except Exception as ex:
                messagebox.showerror("خطأ", f"تعذر نسخ الصورة:\n{ex}", parent=win); return
            refresh(key)

        def view(key):
            if not photos[key]:
                return
            p = ensure_photo(photos[key]) or os.path.join(PHOTOS_DIR, photos[key])
            try:
                os.startfile(p)
            except Exception:
                messagebox.showinfo("الصورة", f"مسار الصورة:\n{p}", parent=win)

        def clear(key):
            photos[key] = ""; refresh(key)

        def photo_box(title, key, col):
            box = tk.Frame(pframe, bg=WHITE); box.grid(row=0, column=col, padx=22, pady=8)
            tk.Label(box, text=title, bg=WHITE, fg=TEXT, font=(FONT, 10, "bold")).pack()
            cont = tk.Frame(box, width=190, height=150, bg=SOFT, highlightbackground=LINE, highlightthickness=1)
            cont.pack(pady=4); cont.pack_propagate(False)
            lbl = tk.Label(cont, bg=SOFT, fg=MUTED, font=(FONT, 9), text="لا توجد صورة")
            lbl.pack(expand=True, fill="both"); labels[key] = lbl
            bb = tk.Frame(box, bg=WHITE); bb.pack()
            for txt, fn, color in (("اختيار", lambda: choose(key), BLUE), ("عرض", lambda: view(key), PRIMARY), ("إزالة", lambda: clear(key), RED)):
                tk.Button(bb, text=txt, command=fn, bg=color, fg="white", relief="flat", font=(FONT, 9, "bold"),
                          padx=10, pady=3, cursor="hand2").pack(side="right", padx=2)
            refresh(key)
        photo_box("صورة العامل", "photo", 1)
        photo_box("صورة البطاقة", "id", 0)

        def save():
            name = " ".join(v["name"].get().split())
            if not name:
                messagebox.showwarning("بيانات ناقصة", "اكتب اسم العامل.", parent=win); return

            def dt(key, label):
                s_ = v[key].get().strip()
                if not s_:
                    return ""
                try:
                    return parse_date(s_).isoformat()
                except ValueError:
                    raise ValueError(label)
            try:
                birth = dt("birth", "تاريخ الميلاد"); start = dt("start", "تاريخ بدء العمل"); end = dt("end", "تاريخ انتهاء العمل")
            except ValueError as ex:
                messagebox.showerror("التاريخ", f"{ex} غير صحيح، اكتبه بالشكل يوم/شهر/سنة.", parent=win); return
            if wid and not self.need("workers", "edit"):
                return
            status = v["status"].get()
            if status == STATUS_OFF and not end:
                end = date.today().isoformat()
            if status == STATUS_ON:
                end = ""
            if start and end and end < start:
                messagebox.showerror("التاريخ", "تاريخ انتهاء العمل قبل تاريخ البدء.", parent=win); return
            nid = v["nid"].get().strip()
            if nid and not (nid.isdigit() and len(nid) == 14):
                if not messagebox.askyesno("الرقم القومي", "الرقم القومي ليس 14 رقمًا. هل تريد الحفظ على أي حال؟", parent=win):
                    return
            vals = (name, v["dept"].get(), v["title"].get().strip(), nid, v["phone"].get().strip(), birth,
                    v["address"].get().strip(), start, status, end, v["notes"].get().strip(), photos["photo"], photos["id"])
            try:
                with db() as con:
                    if wid:
                        con.execute("""UPDATE employees SET name=?,department=?,job_title=?,national_id=?,phone=?,
                            birth_date=?,address=?,start_date=?,status=?,end_date=?,notes=?,photo_file=?,id_file=?
                            WHERE id=?""", vals + (wid,))
                        if name != old_name:   # يبقى سجل الملاحظات مرتبطًا بالاسم الجديد
                            for tbl in ("employee_notes", "permits", "absences"):
                                con.execute(f"UPDATE {_ident(tbl)} SET employee=? WHERE employee=?", (name, old_name))
                    else:
                        con.execute("""INSERT INTO employees(name,department,job_title,national_id,phone,birth_date,
                            address,start_date,status,end_date,notes,photo_file,id_file,created_at)
                            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", vals + (now_stamp(),))
            except sqlite3.IntegrityError:
                messagebox.showwarning("اسم مكرر", "يوجد عامل بنفس الاسم بالفعل.", parent=win); return
            win.destroy()
            self.load_workers()

        bx = tk.Frame(win, bg=BG); bx.pack(pady=10)
        self.btn(bx, "💾 حفظ", save, GREEN).pack(side="right", padx=6)
        self.btn(bx, "إلغاء", win.destroy, MUTED).pack(side="right", padx=6)

    def delete_worker(self):
        sel = self.w_tree.selection()
        if not sel:
            messagebox.showwarning("اختر عاملًا", "حدد العامل المطلوب حذفه من الجدول أولًا."); return
        if not messagebox.askyesno("تأكيد الحذف", "حذف العامل نهائيًا من القاعدة؟\n(الأفضل لمن ترك العمل تغيير حالته إلى «ترك العمل» بدل الحذف)\nملاحظاته السابقة تظل محفوظة في السجل."):
            return
        with db() as con:
            con.execute("DELETE FROM employees WHERE id=?", (int(sel[0]),))
        self.load_workers()

    # ---------- تحميل وحذف ----------
    LOADERS = {
        "visits": ("SELECT id,person,entity,reason,created_at FROM visits", None),
        "maintenance": ("SELECT id,person,entity,notes,created_at FROM maintenance", None),
        "employee_notes": ("SELECT n.id,n.employee,COALESCE(e.department,'-'),n.kind,n.severity,n.details,n.action,n.created_at "
                           "FROM employee_notes n LEFT JOIN employees e ON e.name=n.employee", "n.employee"),
    }

    def load_recent(self, table, tree, search=""):
        sql, search_col = self.LOADERS[table]
        args = ()
        if search and search_col:
            sql += f" WHERE {search_col} LIKE ?"; args = (f"%{search}%",)
        with db() as con:
            rows = con.execute(sql + " ORDER BY 1 DESC LIMIT 200", args).fetchall()
        tree.delete(*tree.get_children())
        for n, row in enumerate(rows):
            rid, vals = row[0], list(row[1:])
            vals[-1] = fmt_stamp(vals[-1])
            tags = ["odd"] if n % 2 else []
            if table == "employee_notes":
                if vals[2] == "خطأ": tags.append("bad")
                elif vals[2] == "إشادة": tags.append("good")
            tree.insert("", "end", iid=str(rid), values=vals, tags=tags)

    # (العمود، العنوان، النوع، القيم، إجباري؟)
    EDIT_SPECS = {
        "visits": ("تعديل بيانات الزيارة", [
            ("person", "اسم الشخص", "entry", None, True), ("entity", "الجهة", "entry", None, False),
            ("reason", "سبب الزيارة", "entry", None, False)]),
        "maintenance": ("تعديل عملية الصيانة", [
            ("person", "اسم الفني / المسؤول", "entry", None, False), ("entity", "الجهة / الشركة", "entry", None, False),
            ("notes", "ملاحظات الصيانة", "text", None, True)]),
        "employee_notes": ("تعديل ملاحظة / خطأ عامل", [
            ("employee", "اسم العامل (اختر أو اكتب)", "combo_free", "workers", True),
            ("kind", "النوع", "combo", NOTE_KINDS, False), ("severity", "الدرجة (للأخطاء والتنبيهات)", "combo", SEVERITIES, False),
            ("details", "التفاصيل", "text", None, True), ("action", "الإجراء المتخذ", "entry", None, False)]),
        "gov_staff": ("تعديل بيانات موظف بالجهة", [
            ("entity", "الجهة", "combo", GOV_ENTITIES, True), ("name", "الاسم", "entry", None, True),
            ("job_title", "الوظيفة", "entry", None, False), ("phone", "الهاتف", "entry", None, False),
            ("national_id", "الرقم القومي (14 رقم)", "entry", None, False),
            ("status", "الحالة", "combo", [STATUS_ON, STATUS_OFF], False), ("notes", "ملاحظات", "text", None, False)]),
        "gov_attendance": ("تعديل حضور وانصراف", [
            ("person", "الاسم", "entry", None, True), ("in_time", "وقت الحضور (مثل 7:30)", "entry", None, False),
            ("out_time", "وقت الانصراف (مثل 4:00 م)", "entry", None, False), ("notes", "ملاحظات", "text", None, False)]),
    }

    TABLE_KIND = {"visits": ("visits", "زيارات المركز"), "maintenance": ("maintenance", "سجل الصيانة"),
                  "employee_notes": ("employees", "ملاحظات وأخطاء العاملين")}

    def _ask_period(self, title, table):
        """نافذة اختيار الفترة (الافتراضي = من أول سجل حتى اليوم أي كل السجلات). ترجع (start, end) أو None."""
        win = self._dialog(title, "560x250")
        res = {"v": None}
        tk.Label(win, text=title, bg=BG, fg=INK, font=(FONT, 14, "bold")).pack(anchor="e", padx=20, pady=(14, 2))
        tk.Label(win, text="الفترة الافتراضية تشمل كل السجلات، وتقدر تغيّرها.", bg=BG, fg=MUTED,
                 font=(FONT, 9)).pack(anchor="e", padx=22)
        f1 = tk.StringVar(value=earliest_record_date(table).strftime("%d/%m/%Y")); f2 = tk.StringVar(value=today_text())
        box = self.panel(win); box.pack(fill="x", padx=16, pady=8)
        self.add_date(box, "من تاريخ", f1, 0, 1)
        self.add_date(box, "إلى تاريخ", f2, 0, 0)

        def ok():
            try:
                d1, d2 = parse_date(f1.get()), parse_date(f2.get())
                if d2 < d1: raise ValueError
            except ValueError:
                messagebox.showerror("التاريخ", "راجع الفترة (يوم/شهر/سنة).", parent=win); return
            res["v"] = (d1.isoformat(), d2.isoformat()); win.destroy()
        row = tk.Frame(win, bg=BG); row.pack(pady=8)
        self.btn(row, "متابعة", ok, GREEN).pack(side="right", padx=6)
        self.btn(row, "إلغاء", win.destroy, MUTED).pack(side="right", padx=6)
        self.root.wait_window(win)
        return res["v"]

    def _table_sections(self, table):
        kind, title = self.TABLE_KIND[table]
        p = self._ask_period(title, table)
        if not p:
            return None
        secs = report_sections(kind, *p)
        if table == "employee_notes":
            key = self.emp_search.get().strip() if hasattr(self, "emp_search") else ""
            if key:                                      # نفس بحث الشاشة (جزء من اسم العامل)
                for x in secs:
                    x["rows"] = [r for r in x["rows"] if r and key in str(r[0])]
                    x["period"] += f"   |   بحث: {key}"
        return title, secs

    def print_table(self, table):
        r = self._table_sections(table)
        if not r: return
        title, secs = r
        try:
            open_html(sections_html(title, secs), table)
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تجهيز الطباعة:\n{ex}")

    def export_table(self, table):
        r = self._table_sections(table)
        if not r: return
        title, secs = r
        self.export_sections_excel(secs, f"{title}_{date.today():%Y-%m-%d}.xlsx", title)

    def _reload(self, table, tree):
        if table == "gov_staff":
            self.govstaff_reload(); return
        if table == "gov_attendance":
            self.gov_reload(); return
        search = self.emp_search.get().strip() if table == "employee_notes" else ""
        self.load_recent(table, tree, search)

    def edit_record(self, table, tree):
        """تعديل السجل المحدد (الزيارات / الصيانة / ملاحظات العاملين) مع الإبقاء على تاريخ ووقت التسجيل الأصلي."""
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("اختر سجلًا", "حدد السجل المطلوب تعديله من الجدول أولًا."); return
        rid = int(sel[0])
        title, fields = self.EDIT_SPECS[table]
        with db() as con:
            row = con.execute(f"SELECT {','.join(_ident(f[0]) for f in fields)} FROM {_ident(table)} WHERE id=?", (rid,)).fetchone()
        if not row:
            messagebox.showerror("خطأ", "السجل غير موجود."); self._reload(table, tree); return
        win = self._dialog(title, "640x600")
        self._minsize(win, 560, 420)
        bx = tk.Frame(win, bg=BG); bx.pack(side="bottom", fill="x", pady=12)
        tk.Label(win, text="✏️ " + title, bg=BG, fg=INK, font=(FONT, 15, "bold")).pack(anchor="e", padx=22, pady=(14, 4))
        getters = {}
        for (col, label, kind, vals, req), val in zip(fields, row):
            box = tk.Frame(win, bg=BG); box.pack(fill="x", padx=22, pady=5)
            tk.Label(box, text=label + (" *" if req else ""), bg=BG, fg=TEXT, font=(FONT, 10, "bold"),
                     anchor="e").pack(fill="x", pady=(0, 3))
            if kind == "text":
                w = tk.Text(box, height=5, font=(FONT, 10), relief="solid", bd=1, wrap="word",
                            bg=ENTRY, fg=TEXT, insertbackground=TEXT)
                w.tag_configure("r", justify="right")
                w.insert("1.0", val or ""); w.tag_add("r", "1.0", "end")
                w.bind("<KeyRelease>", lambda e, w=w: w.tag_add("r", "1.0", "end"))
                w.pack(fill="x")
                getters[col] = lambda w=w: w.get("1.0", "end").strip()
            elif kind.startswith("combo"):
                var = tk.StringVar(value=val or "")
                values = self.worker_names() if vals == "workers" else list(vals)
                ttk.Combobox(box, textvariable=var, values=values, justify="right",
                             state="readonly" if kind == "combo" else "normal").pack(fill="x")
                getters[col] = lambda var=var: var.get().strip()
            else:
                var = tk.StringVar(value=val or "")
                ttk.Entry(box, textvariable=var, justify="right").pack(fill="x")
                getters[col] = lambda var=var: var.get().strip()

        def save():
            data = {c: g() for c, g in getters.items()}
            for col, label, kind, vals, req in fields:
                if req and not data[col]:
                    messagebox.showwarning("بيانات ناقصة", f"املأ حقل «{label}».", parent=win); return
            if table == "gov_staff":
                nid = data["national_id"].translate(_AR_DIGITS)
                if nid and not (nid.isdigit() and len(nid) == 14):
                    messagebox.showwarning("الرقم القومي", "الرقم القومي لازم يكون 14 رقمًا.", parent=win); return
                data["national_id"] = nid
            if table == "gov_attendance":
                try:
                    i_ = parse_time_flex(data["in_time"]); o_ = parse_time_flex(data["out_time"], True, i_ or None)
                except ValueError as ex:
                    messagebox.showerror("الوقت", str(ex), parent=win); return
                if not i_ and not o_:
                    messagebox.showwarning("بيانات ناقصة", "اكتب وقت الحضور أو الانصراف.", parent=win); return
                if i_ and o_ and o_ <= i_:
                    messagebox.showwarning("الوقت", "وقت الانصراف لازم يكون بعد الحضور.", parent=win); return
                data["in_time"], data["out_time"] = i_, o_
            if table == "employee_notes" and data["kind"] not in ("خطأ", "تنبيه شفهي"):
                data["severity"] = ""
            with db() as con:
                con.execute(f"UPDATE {_ident(table)} SET " + ",".join(f"{_ident(c)}=?" for c in data) + " WHERE id=?",
                            (*data.values(), rid))
            win.destroy()
            self._reload(table, tree)
            if tree.exists(str(rid)):
                tree.selection_set(str(rid)); tree.see(str(rid))
            messagebox.showinfo("تم التعديل", "تم حفظ التعديل (تاريخ ووقت التسجيل الأصلي كما هو).")
        inner = tk.Frame(bx, bg=BG); inner.pack()
        self.btn(inner, "💾 حفظ التعديل", save, GREEN).pack(side="right", padx=6)
        self.btn(inner, "إلغاء", win.destroy, MUTED).pack(side="right", padx=6)

    def delete_selected(self, table, tree):
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("اختر سجلًا", "حدد السجل المطلوب حذفه من الجدول أولًا."); return
        if not messagebox.askyesno("تأكيد الحذف", "هل تريد حذف السجل المحدد نهائيًا؟"):
            return
        with db() as con:
            con.execute(f"DELETE FROM {_ident(table)} WHERE id=?", (int(sel[0]),))
        self._reload(table, tree)

    # ---------- الجرد ----------
    def show_inventory(self, category):
        labels = {"cleaning": "جرد خامات النظافة", "paper": "جرد الأدوات الورقية", "toner": "جرد الأحبار"}
        self.page_title(labels[category], "سجّل الوارد الشهري فيُضاف على الرصيد تلقائيًا، ثم أدخل الجرد الفعلي ليظهر الاستهلاك في التقرير")
        top = self.panel(); top.pack(fill="x", padx=24, pady=6)
        self.inv_from = tk.StringVar(value=today_text()); self.inv_to = tk.StringVar(value=today_text())
        self.add_date(top, "من تاريخ", self.inv_from, 0, 1)
        self.add_date(top, "إلى تاريخ", self.inv_to, 0, 0)
        self.btn(top, "💾 حفظ الجرد", lambda: self.save_inventory(category), GREEN).grid(row=0, column=2, padx=18, pady=22)
        bx = tk.Frame(top, bg=WHITE); bx.grid(row=1, column=0, columnspan=3, sticky="e", padx=14, pady=(0, 10))
        self.btn(bx, "📥 تسجيل وارد جديد", lambda: self.stock_in_dialog(category), BLUE).pack(side="right", padx=4)
        self.btn(bx, "📋 سجل الوارد", lambda: self.stock_log_dialog(category), PRIMARY).pack(side="right", padx=4)
        self.btn(bx, "⬇️ تعبئة الجرد بالرصيد المتوقع", self.fill_expected, PURPLE).pack(side="right", padx=4)
        self.btn(bx, "🖨️ طباعة الجرد الحالي", lambda: self.print_inventory(category), AMBER).pack(side="right", padx=4)
        self.btn(bx, "📊 استخراج الجرد الحالي Excel", lambda: self.export_inventory(category), GREEN).pack(side="right", padx=4)

        bar = tk.Frame(self.content, bg=BG); bar.pack(side="bottom", fill="x", pady=(0, 8))
        bb = tk.Frame(bar, bg=BG); bb.pack(pady=4)
        self.btn(bb, f"💾 حفظ {INV_NAMES[category]}", lambda: self.save_inventory(category), GREEN).pack(side="right", padx=6)
        self.btn(bb, "🖨️ طباعة الجرد الحالي", lambda: self.print_inventory(category), AMBER).pack(side="right", padx=6)
        self.btn(bb, "📊 استخراج الجرد الحالي Excel", lambda: self.export_inventory(category), GREEN).pack(side="right", padx=6)
        tk.Label(bar, text="الرصيد المتوقع = كمية آخر جرد + الوارد المسجل بعد تاريخه.   اضغط مرتين على «الجرد الفعلي» أو «ملاحظات» للتعديل (Enter حفظ - Esc إلغاء).",
                 bg=BG, fg=MUTED, font=(FONT, 9)).pack()
        wrap = tk.Frame(self.content, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        wrap.pack(fill="both", expand=True, padx=24, pady=8)
        cols = ("item", "qty", "notes", "last", "inc", "exp")
        self.inv_tree = ttk.Treeview(wrap, columns=cols, show="headings", height=13)
        for col, head, w in [("item", "الصنف", 230), ("qty", "الجرد الفعلي الجديد", 130), ("notes", "ملاحظات", 170),
                             ("last", "كمية آخر جرد", 100), ("inc", "وارد بعد آخر جرد", 120), ("exp", "الرصيد المتوقع", 110)]:
            self.inv_tree.heading(col, text=head)
            self.inv_tree.column(col, width=w, anchor="e" if col in ("item", "notes") else "center")
        self.inv_tree.tag_configure("odd", background=ODD)
        self.inv_tree.tag_configure("bad", foreground=RED)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.inv_tree.yview)
        self.inv_tree.configure(yscrollcommand=sb.set)
        self.inv_tree.pack(side="right", fill="both", expand=True); sb.pack(side="left", fill="y")
        self.inv_stat = stock_status(category)
        for n, item in enumerate(INV_ITEMS[category]):
            lq, wt, after, exp = self.inv_stat[item]
            tags = ["odd"] if n % 2 else []
            if exp is not None and exp <= 0: tags.append("bad")
            self.inv_tree.insert("", "end", values=(item, "", "", fmt_num(lq), fmt_num(after) if after else "-", fmt_num(exp)), tags=tags)
        editable = {"#2": "qty", "#3": "notes"}
        self.inv_tree.bind("<Double-1>", lambda e: self.edit_cell(self.inv_tree, e, editable))

    def _typed_inventory(self):
        typed = {}
        for iid in self.inv_tree.get_children():
            v = self.inv_tree.item(iid, "values")
            if str(v[1]).strip():
                typed[v[0]] = str(v[1]).strip()
        return typed

    def export_sections_excel(self, sections, default_name, what="التقرير"):
        """يسأل عن مكان الحفظ ثم ينشئ ملف Excel ويعرض فتحه."""
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name,
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        self.root.config(cursor="watch"); self.root.update_idletasks()
        try:
            res = build_sections_excel(path, sections)
        except PermissionError:
            messagebox.showerror("الملف مفتوح", "الملف مفتوح في Excel. أغلقه أو اختر اسمًا آخر ثم أعد المحاولة."); return
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر إنشاء الملف:\n{ex}"); return
        finally:
            self.root.config(cursor="")
        total = sum(res.values())
        msg = f"تم استخراج {what} بنجاح ({total} سجل):\n{path}"
        if messagebox.askyesno("تم الاستخراج", msg + "\n\nهل تريد فتح الملف الآن؟"):
            try: os.startfile(path)
            except Exception: pass

    def export_inventory(self, category):
        try:
            secs = inventory_excel_sections(category, self._typed_inventory())
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تجهيز الجرد:\n{ex}"); return
        self.export_sections_excel(secs, f"{INV_NAMES[category]}_الحالي_{date.today():%Y-%m-%d}.xlsx", "الجرد الحالي")

    def print_inventory(self, category):
        """يطبع الموجود فعليًا حاليًا بالمخزن (وأي كميات مكتوبة بالجدول ولم تُحفظ تظهر في عمود الجرد الفعلي)."""
        typed = {}
        try:
            for iid in self.inv_tree.get_children():
                v = self.inv_tree.item(iid, "values")
                if str(v[1]).strip():
                    typed[v[0]] = str(v[1]).strip()
            open_html(inventory_print_html(category, typed), f"inventory_{category}")
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تجهيز الطباعة:\n{ex}")

    def fill_expected(self):
        n = 0
        for iid in self.inv_tree.get_children():
            item = self.inv_tree.item(iid, "values")[0]
            exp = self.inv_stat.get(item, (None, None, 0, None))[3]
            if exp is not None:
                self.inv_tree.set(iid, "qty", fmt_num(exp)); n += 1
        if not n:
            messagebox.showinfo("لا يوجد رصيد", "لا توجد أرصدة سابقة أو وارد مسجل بعد. أدخل الجرد يدويًا أو سجّل الوارد أولًا.")

    def edit_cell(self, tree, event, editable):
        """تعديل خانة داخل جدول بالضغط المزدوج. editable = {'#2': 'qty', '#3': 'notes'}"""
        iid = tree.identify_row(event.y); col = tree.identify_column(event.x)
        if not iid or col not in editable:
            return
        x, y, w, h = tree.bbox(iid, col)
        key = editable[col]
        ent = ttk.Entry(tree, justify="center" if key == "qty" else "right")
        ent.place(x=x, y=y, width=w, height=h)
        ent.insert(0, tree.set(iid, key)); ent.focus_set()
        state = {"done": False}

        def finish(_=None):
            if state["done"]:          # يمنع تنفيذ Enter ثم FocusOut مرتين
                return
            state["done"] = True
            val = ent.get().strip()
            if key == "qty" and val:
                try:
                    if float(val) < 0: raise ValueError
                except ValueError:
                    ent.destroy()
                    messagebox.showerror("قيمة غير صحيحة", "أدخل كمية رقمية صفر أو أكبر."); return
            tree.set(iid, key, val); ent.destroy()

        def cancel(_=None):
            state["done"] = True; ent.destroy()
        ent.bind("<Return>", finish); ent.bind("<FocusOut>", finish); ent.bind("<Escape>", cancel)

    def save_inventory(self, category):
        try:
            d1 = parse_date(self.inv_from.get()); d2 = parse_date(self.inv_to.get())
            if d2 < d1: raise ValueError
        except ValueError:
            messagebox.showerror("التاريخ", "اكتب التاريخ بالشكل يوم/شهر/سنة وتأكد من صحة الفترة."); return
        records = []
        for iid in self.inv_tree.get_children():
            vals = self.inv_tree.item(iid, "values")
            item, qty, notes = vals[0], vals[1], vals[2]
            if str(qty).strip():
                records.append((d1.isoformat(), d2.isoformat(), item, float(qty), notes, now_stamp(), category))
        if not records:
            messagebox.showwarning("لا توجد بيانات", "أدخل كمية صنف واحد على الأقل."); return
        with db() as con:
            old = con.execute("SELECT COUNT(*) FROM inventory WHERE category=? AND week_from=? AND week_to=?",
                              (category, d1.isoformat(), d2.isoformat())).fetchone()[0]
            if old:
                if not messagebox.askyesno("جرد موجود", "يوجد جرد محفوظ لنفس الفترة. هل تريد استبداله؟\n(لتجنب تكرار الأصناف في التقرير)"):
                    return
                con.execute("DELETE FROM inventory WHERE category=? AND week_from=? AND week_to=?",
                            (category, d1.isoformat(), d2.isoformat()))
            con.executemany("""INSERT INTO inventory(week_from,week_to,item,quantity,notes,saved_at,category)
                VALUES(?,?,?,?,?,?,?)""", records)
        messagebox.showinfo("تم الحفظ", f"تم حفظ {len(records)} صنفًا في الجرد.")
        self.go(category)

    def _dialog(self, title, size):
        win = tk.Toplevel(self.root)
        size = self._clamp_size(size)
        win.title(title); win.configure(bg=BG); win.geometry(size)
        try:
            win.transient(self.root)
            px = self.root.winfo_rootx() + 120
            py = max(0, self.root.winfo_rooty() + 20)
            try:
                sw, sh = int(win.winfo_screenwidth()), int(win.winfo_screenheight())
                dw, dh = (int(v) for v in size.lower().split("x")[:2])
                px = max(0, min(px, sw - dw - 10))
                py = max(0, min(py, sh - dh - 40))
            except Exception:
                pass
            win.geometry(f"{size}+{px}+{py}")
            win.bind("<Escape>", lambda e: win.destroy())   # Esc يغلق أي نافذة
            win.wait_visibility(); win.grab_set()
        except Exception:
            pass
        return win

    def stock_in_dialog(self, category):
        win = self._dialog(f"تسجيل وارد - {INV_NAMES[category]}", "760x740")
        self._minsize(win, 640, 520)
        bx = tk.Frame(win, bg=BG); bx.pack(side="bottom", fill="x", pady=10)      # أزرار الحفظ أولًا لتظل ظاهرة دائمًا
        tk.Label(win, text=f"📥 تسجيل وارد جديد - {INV_NAMES[category]}", bg=BG, fg=INK,
                 font=(FONT, 15, "bold")).pack(anchor="e", padx=20, pady=(12, 2))
        tk.Label(win, text="يُضاف الوارد على الرصيد الموجود. اكتب الكمية المستلمة لكل صنف فقط واترك الباقي فارغًا.",
                 bg=BG, fg=MUTED, font=(FONT, 9)).pack(anchor="e", padx=22)
        top = self.panel(win); top.pack(fill="x", padx=16, pady=6)
        d_var, n_var = tk.StringVar(value=today_text()), tk.StringVar()
        self.add_date(top, "تاريخ الاستلام", d_var, 0, 1)
        self.add_field(top, "المورد / رقم إذن الاستلام (اختياري)", n_var, 0, 0, 300)
        inv = {"src": ""}
        ibox = tk.Frame(top, bg=WHITE); ibox.grid(row=1, column=0, columnspan=2, sticky="e", padx=14, pady=(0, 8))
        inv_lbl = tk.Label(ibox, text="لم تُرفق صورة فاتورة", bg=WHITE, fg=MUTED, font=(FONT, 9))

        def pick_invoice():
            path = filedialog.askopenfilename(parent=win, title="اختر صورة الفاتورة", filetypes=[
                ("صور الفاتورة", "*.jpg *.jpeg *.png *.gif *.bmp *.webp *.pdf"), ("كل الملفات", "*.*")])
            if path:
                inv["src"] = path; inv_lbl.config(text="✓ " + os.path.basename(path), fg=GREEN)

        def clear_invoice():
            inv["src"] = ""; inv_lbl.config(text="لم تُرفق صورة فاتورة", fg=MUTED)
        self.btn(ibox, "📎 إدراج فاتورة من صورة", pick_invoice, PURPLE).pack(side="right")
        self.btn(ibox, "✖", clear_invoice, MUTED).pack(side="right", padx=4)
        inv_lbl.pack(side="right", padx=8)
        wrap = tk.Frame(win, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        wrap.pack(fill="both", expand=True, padx=16, pady=6)
        tree = ttk.Treeview(wrap, columns=("item", "qty", "notes"), show="headings", height=14)
        for col, head, w in (("item", "الصنف", 290), ("qty", "الكمية الواردة", 130), ("notes", "ملاحظة الصنف", 230)):
            tree.heading(col, text=head); tree.column(col, width=w, anchor="e" if col != "qty" else "center")
        tree.tag_configure("odd", background=ODD)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="right", fill="both", expand=True); sb.pack(side="left", fill="y")
        for n, item in enumerate(INV_ITEMS[category]):
            tree.insert("", "end", values=(item, "", ""), tags=("odd",) if n % 2 else ())
        editable = {"#2": "qty", "#3": "notes"}
        tree.bind("<Double-1>", lambda e: self.edit_cell(tree, e, editable))

        def save():
            try:
                iso = parse_date(d_var.get()).isoformat()
            except ValueError:
                messagebox.showerror("التاريخ", "راجع تاريخ الاستلام.", parent=win); return
            recs = []
            for iid in tree.get_children():
                item, qty, note = tree.item(iid, "values")[:3]
                try:
                    q = float(str(qty).strip()) if str(qty).strip() else 0
                except ValueError:
                    continue
                if q > 0:
                    recs.append((category, item, q, iso, " - ".join(x for x in (n_var.get().strip(), str(note).strip()) if x), now_stamp()))
            if not recs:
                messagebox.showwarning("لا توجد بيانات", "اكتب كمية صنف واحد على الأقل (أكبر من صفر).", parent=win); return
            fname = ""
            if inv["src"]:
                try:
                    fname = save_photo(inv["src"], "stockin")
                except Exception as ex:
                    messagebox.showerror("الفاتورة", f"تعذر حفظ صورة الفاتورة:\n{ex}", parent=win); return
            with db() as con:
                con.executemany("INSERT INTO stock_in(category,item,qty,in_date,notes,created_at,invoice_file) VALUES(?,?,?,?,?,?,?)",
                                [r + (fname,) for r in recs])
            win.destroy()
            messagebox.showinfo("تم الحفظ", f"تم تسجيل وارد {len(recs)} صنفًا وإضافته على الرصيد.")
            self.go(category)
        inner = tk.Frame(bx, bg=BG); inner.pack()
        self.btn(inner, "💾 حفظ الوارد", save, GREEN).pack(side="right", padx=6)
        self.btn(inner, "إلغاء", win.destroy, MUTED).pack(side="right", padx=6)

    def stock_log_dialog(self, category):
        win = self._dialog(f"سجل الوارد - {INV_NAMES[category]}", "800x600")
        self._minsize(win, 640, 460)
        bx = tk.Frame(win, bg=BG); bx.pack(side="bottom", fill="x", pady=8)       # الأزرار أولًا لتظل ظاهرة
        tk.Label(win, text=f"📋 سجل الوارد - {INV_NAMES[category]}", bg=BG, fg=INK,
                 font=(FONT, 15, "bold")).pack(anchor="e", padx=20, pady=(12, 2))
        tk.Label(win, text="اضغط مرتين على «الكمية» أو «ملاحظات» للتعديل، ثم اضغط «حفظ التعديلات».",
                 bg=BG, fg=MUTED, font=(FONT, 9)).pack(anchor="e", padx=22)
        tree = self.make_tree(win, ("date", "item", "qty", "notes", "inv"), ("تاريخ الاستلام", "الصنف", "الكمية", "ملاحظات", "الفاتورة"),
                              (105, 210, 70, 220, 70), 14)
        orig = {}

        def load():
            with db() as con:
                rows = con.execute("""SELECT id, in_date, item, qty, notes, invoice_file FROM stock_in WHERE category=?
                    ORDER BY in_date DESC, id DESC LIMIT 300""", (category,)).fetchall()
            tree.delete(*tree.get_children()); orig.clear()
            for n, (rid, d, item, q, note, invf) in enumerate(rows):
                tree.insert("", "end", iid=str(rid), tags=("odd",) if n % 2 else (),
                            values=(fmt_iso_date(d), item, fmt_num(q), note or "", "🧾 ✓" if invf else "-"))
                orig[str(rid)] = (fmt_num(q), note or "")
        load()
        tree.bind("<Double-1>", lambda e: self.edit_cell(tree, e, {"#3": "qty", "#4": "notes"}))

        def save_edits():
            if not self.need(category, "edit"):
                return
            changes = []
            for iid in tree.get_children():
                vals = tree.item(iid, "values")
                q, note = str(vals[2]).strip(), str(vals[3]).strip()
                if (q, note) == orig.get(iid):
                    continue
                try:
                    qv = float(q)
                    if qv <= 0: raise ValueError
                except ValueError:
                    messagebox.showerror("كمية غير صحيحة", f"الكمية في صنف «{vals[1]}» لازم تكون رقم أكبر من صفر.", parent=win); return
                changes.append((qv, note, int(iid)))
            if not changes:
                messagebox.showinfo("لا تغييرات", "لم يتم تعديل أي سجل.", parent=win); return
            with db() as con:
                con.executemany("UPDATE stock_in SET qty=?, notes=? WHERE id=?", changes)
            messagebox.showinfo("تم الحفظ", f"تم حفظ تعديل {len(changes)} سجل.", parent=win)
            load(); self.go(category)

        def delete():
            if not self.need(category, "edit"):
                return
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("اختر سجلًا", "حدد سجل الوارد المطلوب حذفه.", parent=win); return
            if not messagebox.askyesno("تأكيد", "حذف سجل الوارد المحدد؟ (سيُخصم من الرصيد المتوقع)", parent=win):
                return
            with db() as con:
                con.execute("DELETE FROM stock_in WHERE id=?", (int(sel[0]),))
            load(); self.go(category)
        inner = tk.Frame(bx, bg=BG); inner.pack()
        def view_inv():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("اختر سجلًا", "حدد سجل الوارد لعرض فاتورته.", parent=win); return
            with db() as con:
                r = con.execute("SELECT invoice_file FROM stock_in WHERE id=?", (int(sel[0]),)).fetchone()
            if not r or not r[0]:
                messagebox.showinfo("لا توجد فاتورة", "هذا السجل لم تُرفق له صورة فاتورة.", parent=win); return
            self.open_invoice(r[0])

        def attach_inv():
            if not self.need(category, "edit"):
                return
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("اختر سجلًا", "حدد سجل الوارد لإدراج فاتورته.", parent=win); return
            path = filedialog.askopenfilename(parent=win, title="اختر صورة الفاتورة", filetypes=[
                ("صور الفاتورة", "*.jpg *.jpeg *.png *.gif *.bmp *.webp *.pdf"), ("كل الملفات", "*.*")])
            if not path:
                return
            try:
                fname = save_photo(path, "stockin")
            except Exception as ex:
                messagebox.showerror("الفاتورة", f"تعذر حفظ الصورة:\n{ex}", parent=win); return
            with db() as con:
                con.execute("UPDATE stock_in SET invoice_file=? WHERE id=?", (fname, int(sel[0])))
            load()
            messagebox.showinfo("تم", "تم إدراج صورة الفاتورة للسجل المحدد.", parent=win)
        self.btn(inner, "💾 حفظ التعديلات", save_edits, GREEN).pack(side="right", padx=6)
        self.btn(inner, "📎 إدراج فاتورة", attach_inv, PURPLE).pack(side="right", padx=6)
        self.btn(inner, "🧾 عرض الفاتورة", view_inv, BLUE).pack(side="right", padx=6)
        self.btn(inner, "حذف المحدد", delete, RED).pack(side="right", padx=6)
        self.btn(inner, "إغلاق", win.destroy, MUTED).pack(side="right", padx=6)

    # ---------- المشتريات ----------
    def show_purchases(self):
        self.page_title("المشتريات", "سجّل مشتريات الورقيات والأحبار والأدوات يدويًا مع إرفاق صورة الفاتورة")
        top = self.panel(); top.pack(fill="x", padx=24, pady=6)
        self.pur_from = tk.StringVar(value=date.today().replace(month=1, day=1).strftime("%d/%m/%Y"))
        self.pur_to = tk.StringVar(value=today_text())
        self.pur_cat = tk.StringVar(value="all")
        self.add_date(top, "من تاريخ", self.pur_from, 0, 2)
        self.add_date(top, "إلى تاريخ", self.pur_to, 0, 1)
        fb = tk.Frame(top, bg=WHITE); fb.grid(row=0, column=0, padx=14, sticky="e")
        tk.Label(fb, text="النوع", bg=WHITE, fg=TEXT, font=(FONT, 10, "bold")).pack(side="right", padx=6)
        for val, label in [("all", "الكل")] + list(PURCH_CATS.items()):
            ttk.Radiobutton(fb, text=label, value=val, variable=self.pur_cat, command=self.load_purchases).pack(side="right", padx=5)
        bx = tk.Frame(top, bg=WHITE); bx.grid(row=1, column=0, columnspan=3, sticky="e", padx=14, pady=(0, 10))
        self.btn(bx, "➕ تسجيل مشتريات جديدة", lambda: self.purchase_dialog(
            self.pur_cat.get() if self.pur_cat.get() in PURCH_CATS else "paper"), GREEN).pack(side="right", padx=4)
        self.btn(bx, "🧾 عرض صورة الفاتورة", self.view_invoice, BLUE).pack(side="right", padx=4)
        self.btn(bx, "📎 إدراج فاتورة من صورة", self.attach_invoice, PURPLE).pack(side="right", padx=4)
        self.btn(bx, "👤 تعديل «تم السداد بواسطة»", self.edit_paid_by, TEAL).pack(side="right", padx=4)
        self.btn(bx, "🔍 عرض", self.load_purchases, PRIMARY).pack(side="right", padx=4)
        self.btn(bx, "🖨️ طباعة", self.print_purchases, AMBER).pack(side="right", padx=4)
        self.btn(bx, "📊 استخراج Excel بالمدة", self.export_purchases, GREEN).pack(side="right", padx=4)
        self.btn(bx, "🗑️ حذف المحدد", self.delete_purchase, RED).pack(side="right", padx=4)
        self.pur_total = tk.Label(self.content, text="", bg=BG, fg=INK, font=(FONT, 11, "bold"))
        self.pur_total.pack(side="bottom", pady=(0, 8))
        self.pur_tree = self.make_tree(
            self.content, ("date", "cat", "item", "qty", "price", "total", "sup", "no", "paid", "inv"),
            ("التاريخ", "النوع", "الصنف", "الكمية", "سعر الوحدة", "الإجمالي", "المورد", "رقم الفاتورة", "سُدد بواسطة", "الفاتورة"),
            (90, 65, 170, 60, 80, 90, 120, 85, 120, 65), 12)
        self.pur_tree.bind("<Double-1>", lambda e: self.view_invoice())
        self.load_purchases()

    def _pur_period(self):
        try:
            d1, d2 = parse_date(self.pur_from.get()), parse_date(self.pur_to.get())
            if d2 < d1: raise ValueError
            return d1.isoformat(), d2.isoformat()
        except ValueError:
            messagebox.showerror("التاريخ", "اكتب التاريخ بالشكل يوم/شهر/سنة وتأكد من صحة الفترة.")
            return None

    def load_purchases(self):
        p = self._pur_period()
        if not p: return
        rows = purchases_rows(self.pur_cat.get(), *p)
        tree = self.pur_tree
        tree.delete(*tree.get_children())
        for n, r in enumerate(rows[:500]):
            tree.insert("", "end", iid=str(r[0]), tags=("odd",) if n % 2 else (),
                        values=(fmt_iso_date(r[1]), PURCH_CATS.get(r[2], r[2]), r[3], fmt_num(r[4]), fmt_money(r[5]),
                                fmt_money(r[6]), r[7], r[8], r[10] or "-", "🧾 ✓" if r[9] else "-"))
        total = sum(r[6] or 0 for r in rows)
        extra = f"   (معروض أحدث 500 من {len(rows)})" if len(rows) > 500 else ""
        self.pur_total.config(text=f"عدد البنود: {len(rows)}   |   إجمالي المشتريات: {fmt_money(total)} جنيه{extra}")

    def _invoice_of(self, rid):
        with db() as con:
            r = con.execute("SELECT invoice_file FROM purchases WHERE id=?", (rid,)).fetchone()
        return r[0] if r else ""

    def open_invoice(self, name):
        path = ensure_photo(name) or os.path.join(PHOTOS_DIR, name)
        if not os.path.exists(path):
            messagebox.showwarning("الفاتورة", "ملف الفاتورة غير موجود في مجلد photos."); return
        try:
            os.startfile(path)
        except Exception:
            webbrowser.open(Path(path).as_uri())

    def view_invoice(self):
        sel = self.pur_tree.selection()
        if not sel:
            messagebox.showwarning("اختر سجلًا", "حدد بند المشتريات لعرض فاتورته."); return
        f = self._invoice_of(int(sel[0]))
        if not f:
            messagebox.showinfo("لا توجد فاتورة", "هذا البند لم تُرفق له صورة فاتورة."); return
        self.open_invoice(f)

    def attach_invoice(self):
        """إدراج (أو استبدال) صورة فاتورة لبند مشتريات مسجّل مسبقًا."""
        sel = self.pur_tree.selection()
        if not sel:
            messagebox.showwarning("اختر سجلًا", "حدد بند المشتريات المراد إدراج فاتورته من الجدول أولًا."); return
        rid = int(sel[0])
        with db() as con:
            r = con.execute("SELECT p_date,supplier,invoice_no,invoice_file FROM purchases WHERE id=?", (rid,)).fetchone()
            if not r:
                return
            same = [x[0] for x in con.execute(
                "SELECT id FROM purchases WHERE p_date=? AND supplier=? AND invoice_no=? AND invoice_file=?", r)]
        if r[3] and not messagebox.askyesno("استبدال الفاتورة", "هذا البند له فاتورة مرفقة بالفعل. هل تريد استبدالها بصورة جديدة؟"):
            return
        path = filedialog.askopenfilename(title="اختر صورة الفاتورة", filetypes=[
            ("صور الفاتورة", "*.jpg *.jpeg *.png *.gif *.bmp *.webp *.pdf"), ("كل الملفات", "*.*")])
        if not path:
            return
        ids = [rid]
        if len(same) > 1 and messagebox.askyesno(
                "بنود نفس الفاتورة", f"يوجد {len(same)} بنود بنفس التاريخ والمورد ورقم الفاتورة.\nهل تريد إدراج هذه الصورة لها كلها؟"):
            ids = same
        try:
            fname = save_photo(path, "invoice")
        except Exception as ex:
            messagebox.showerror("الفاتورة", f"تعذر حفظ صورة الفاتورة:\n{ex}"); return
        with db() as con:
            for i in ids:
                con.execute("UPDATE purchases SET invoice_file=? WHERE id=?", (fname, i))
            old = r[3]
            if old and not con.execute("SELECT 1 FROM purchases WHERE invoice_file=?", (old,)).fetchone():
                try: os.remove(os.path.join(PHOTOS_DIR, old))
                except OSError: pass
        self.load_purchases()
        messagebox.showinfo("تم", f"تم إدراج صورة الفاتورة لعدد {len(ids)} بند.")

    def edit_paid_by(self):
        """يكتب أو يعدل اسم الشخص الذي سدّد قيمة بند مشتريات مسجّل (ولبنود نفس الفاتورة لو أردت)."""
        sel = self.pur_tree.selection()
        if not sel:
            messagebox.showwarning("اختر سجلًا", "حدد بند المشتريات من الجدول أولًا."); return
        rid = int(sel[0])
        with db() as con:
            r = con.execute("SELECT p_date,supplier,invoice_no,paid_by FROM purchases WHERE id=?", (rid,)).fetchone()
            if not r:
                return
            same = [x[0] for x in con.execute(
                "SELECT id FROM purchases WHERE p_date=? AND supplier=? AND invoice_no=?", r[:3])]
        name = self._ask_text("تم السداد بواسطة", "اكتب اسم الشخص الذي سدّد قيمة المشتريات:", r[3] or "")
        if name is None:
            return
        ids = [rid]
        if len(same) > 1 and messagebox.askyesno(
                "بنود نفس الفاتورة", f"يوجد {len(same)} بنود بنفس التاريخ والمورد ورقم الفاتورة.\nهل تريد تطبيق الاسم عليها كلها؟"):
            ids = same
        with db() as con:
            for i in ids:
                con.execute("UPDATE purchases SET paid_by=? WHERE id=?", (name, i))
        self.load_purchases()
        messagebox.showinfo("تم", f"تم حفظ «تم السداد بواسطة» لعدد {len(ids)} بند.")

    def _ask_text(self, title, prompt, initial=""):
        """نافذة صغيرة لكتابة نص واحد. ترجع النص أو None عند الإلغاء."""
        win = self._dialog(title, "460x190")
        res = {"v": None}
        tk.Label(win, text=prompt, bg=BG, fg=TEXT, font=(FONT, 10, "bold"), anchor="e").pack(fill="x", padx=20, pady=(18, 6))
        var = tk.StringVar(value=initial)
        ent = ttk.Entry(win, textvariable=var, justify="right", width=48); ent.pack(padx=20)
        ent.focus_set()

        def ok(_=None):
            res["v"] = var.get().strip(); win.destroy()
        ent.bind("<Return>", ok)
        row = tk.Frame(win, bg=BG); row.pack(pady=16)
        self.btn(row, "💾 حفظ", ok, GREEN).pack(side="right", padx=6)
        self.btn(row, "إلغاء", win.destroy, MUTED).pack(side="right", padx=6)
        self.root.wait_window(win)
        return res["v"]

    def export_purchases(self):
        p = self._pur_period()
        if not p: return
        try:
            secs = purchases_excel_sections(self.pur_cat.get(), *p)
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تجهيز المشتريات:\n{ex}"); return
        self.export_sections_excel(secs, f"المشتريات_{p[0]}_إلى_{p[1]}.xlsx", "المشتريات")

    def delete_purchase(self):
        sel = self.pur_tree.selection()
        if not sel:
            messagebox.showwarning("اختر سجلًا", "حدد بند المشتريات المطلوب حذفه."); return
        if not messagebox.askyesno("تأكيد الحذف", "حذف بند المشتريات المحدد؟\n(لو كان مضافًا للمخزون كوارد فسيُخصم من الرصيد أيضًا)"):
            return
        rid = int(sel[0])
        with db() as con:
            r = con.execute("SELECT stock_in_id, invoice_file FROM purchases WHERE id=?", (rid,)).fetchone()
            con.execute("DELETE FROM purchases WHERE id=?", (rid,))
            if r and r[0]:
                con.execute("DELETE FROM stock_in WHERE id=?", (r[0],))
            if r and r[1] and not con.execute("SELECT 1 FROM purchases WHERE invoice_file=?", (r[1],)).fetchone():
                try: os.remove(os.path.join(PHOTOS_DIR, r[1]))
                except OSError: pass
        self.load_purchases()

    def print_purchases(self):
        p = self._pur_period()
        if not p: return
        try:
            open_html(purchases_html(self.pur_cat.get(), *p), "purchases")
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تجهيز الطباعة:\n{ex}")

    def purchase_dialog(self, default_cat="paper"):
        win = self._dialog("تسجيل مشتريات جديدة", "900x840")
        self._minsize(win, 820, 640)
        bx = tk.Frame(win, bg=BG); bx.pack(side="bottom", fill="x", pady=10)      # أزرار الحفظ أولًا لتظل ظاهرة
        tk.Label(win, text="🛒 تسجيل مشتريات جديدة", bg=BG, fg=INK, font=(FONT, 15, "bold")).pack(anchor="e", padx=20, pady=(12, 2))
        tk.Label(win, text="أضف الأصناف واحدًا واحدًا (اختر من القائمة أو اكتب الاسم يدويًا) ثم أرفق صورة الفاتورة واحفظ.",
                 bg=BG, fg=MUTED, font=(FONT, 9)).pack(anchor="e", padx=22)
        name_to_key = {v: k for k, v in PURCH_CATS.items()}
        cat_var = tk.StringVar(value=PURCH_CATS[default_cat])
        d_var, sup_var, no_var, paid_var = tk.StringVar(value=today_text()), tk.StringVar(), tk.StringVar(), tk.StringVar()
        top = self.panel(win); top.pack(fill="x", padx=16, pady=6)
        cat_cb = self.add_combo(top, "نوع المشتريات", cat_var, list(PURCH_CATS.values()), 0, 2)
        self.add_date(top, "تاريخ الفاتورة", d_var, 0, 1)
        self.add_field(top, "المورد / المحل", sup_var, 0, 0, 220)
        self.add_field(top, "رقم الفاتورة", no_var, 1, 2, 200)
        self.add_field(top, "تم السداد بواسطة (اسم الشخص)", paid_var, 1, 1, 220)
        inv = {"src": ""}
        ibox = tk.Frame(top, bg=WHITE); ibox.grid(row=2, column=0, columnspan=3, sticky="e", padx=14, pady=8)
        inv_lbl = tk.Label(ibox, text="لم تُرفق صورة فاتورة", bg=WHITE, fg=MUTED, font=(FONT, 9))

        def pick_invoice():
            path = filedialog.askopenfilename(parent=win, title="اختر صورة الفاتورة", filetypes=[
                ("صور الفاتورة", "*.jpg *.jpeg *.png *.gif *.bmp *.webp *.pdf"), ("كل الملفات", "*.*")])
            if path:
                inv["src"] = path; inv_lbl.config(text="✓ " + os.path.basename(path), fg=GREEN)

        def clear_invoice():
            inv["src"] = ""; inv_lbl.config(text="لم تُرفق صورة فاتورة", fg=MUTED)
        self.btn(ibox, "📎 إرفاق صورة الفاتورة", pick_invoice, BLUE).pack(side="right")
        self.btn(ibox, "✖", clear_invoice, MUTED).pack(side="right", padx=4)
        inv_lbl.pack(side="right", padx=8)

        add = self.panel(win); add.pack(fill="x", padx=16, pady=6)
        it_var, q_var, p_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        item_cb = self.add_combo(add, "الصنف (اختر أو اكتب يدويًا)", it_var, PURCH_ITEMS[default_cat], 0, 3, readonly=False)
        self.add_field(add, "الكمية", q_var, 0, 2, 100)
        self.add_field(add, "سعر الوحدة (جنيه)", p_var, 0, 1, 100)
        cat_cb.bind("<<ComboboxSelected>>", lambda e: (item_cb.configure(values=PURCH_ITEMS[name_to_key[cat_var.get()]]),
                                                       it_var.set(""), refresh_stock_box()))
        lines = {}
        tree = self.make_tree(win, ("item", "qty", "price", "total"), ("الصنف", "الكمية", "سعر الوحدة", "الإجمالي"),
                              (330, 100, 120, 130), 7)
        tot_lbl = tk.Label(win, text="", bg=BG, fg=INK, font=(FONT, 11, "bold"))
        stock_var = tk.BooleanVar(value=True)
        stock_cb = tk.Checkbutton(win, text="إضافة الكميات إلى المخزون كوارد (للأصناف الموجودة في قائمة الجرد فقط)",
                                  variable=stock_var, bg=BG, fg=TEXT, selectcolor=ENTRY, activebackground=BG,
                                  activeforeground=TEXT, font=(FONT, 10), anchor="e")

        def refresh_stock_box():
            if name_to_key[cat_var.get()] in INV_ITEMS:
                stock_cb.pack(anchor="e", padx=22, pady=2, before=tot_lbl)
            else:
                stock_cb.pack_forget()

        def refresh_total():
            tot_lbl.config(text=f"إجمالي الفاتورة: {fmt_money(sum(q * p for _, q, p in lines.values()))} جنيه")

        def add_line():
            item = it_var.get().strip()
            if not item:
                messagebox.showwarning("الصنف", "اختر الصنف أو اكتب اسمه.", parent=win); return
            try:
                q = float(q_var.get().strip().replace(",", "."))
                if q <= 0: raise ValueError
            except ValueError:
                messagebox.showerror("الكمية", "اكتب كمية رقمية أكبر من صفر.", parent=win); return
            try:
                p = float(p_var.get().strip().replace(",", ".") or 0)
                if p < 0: raise ValueError
            except ValueError:
                messagebox.showerror("السعر", "اكتب سعرًا رقميًا (أو اتركه فارغًا).", parent=win); return
            iid = tree.insert("", "end", values=(item, fmt_num(q), fmt_money(p), fmt_money(q * p)),
                              tags=("odd",) if len(lines) % 2 else ())
            lines[iid] = (item, q, p)
            it_var.set(""); q_var.set(""); p_var.set(""); refresh_total()

        def del_line():
            for iid in tree.selection():
                lines.pop(iid, None); tree.delete(iid)
            refresh_total()
        tk.Button(add, text="➕ إضافة للقائمة", command=add_line, bg=GOLD, fg=NAVY, relief="flat", bd=0,
                  font=(FONT, 10, "bold"), padx=14, pady=8, cursor="hand2").grid(row=0, column=0, padx=14, pady=(26, 8))
        tot_lbl.pack(anchor="e", padx=24, pady=(0, 4))
        refresh_stock_box(); refresh_total()

        def save():
            try:
                iso = parse_date(d_var.get()).isoformat()
            except ValueError:
                messagebox.showerror("التاريخ", "راجع تاريخ الفاتورة.", parent=win); return
            if not lines:
                messagebox.showwarning("لا توجد أصناف", "أضف صنفًا واحدًا على الأقل للقائمة (زر «إضافة للقائمة»).", parent=win); return
            cat = name_to_key[cat_var.get()]
            fname = ""
            if inv["src"]:
                try:
                    fname = save_photo(inv["src"], "invoice")
                except Exception as ex:
                    messagebox.showerror("الفاتورة", f"تعذر حفظ صورة الفاتورة:\n{ex}", parent=win); return
            added = 0
            with db() as con:
                for item, q, p in lines.values():
                    sid = None
                    if stock_var.get() and cat in INV_ITEMS and item in INV_ITEMS[cat]:
                        note = " - ".join(x for x in ("مشتريات", sup_var.get().strip(), no_var.get().strip()) if x)
                        sid = con.execute("INSERT INTO stock_in(category,item,qty,in_date,notes,created_at) VALUES(?,?,?,?,?,?)",
                                          (cat, item, q, iso, note, now_stamp())).lastrowid
                        added += 1
                    con.execute("""INSERT INTO purchases(category,item,qty,unit_price,total,supplier,invoice_no,p_date,
                        invoice_file,stock_in_id,created_at,paid_by) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                                (cat, item, q, p, round(q * p, 2), sup_var.get().strip(), no_var.get().strip(), iso,
                                 fname, sid, now_stamp(), paid_var.get().strip()))
            win.destroy()
            msg = f"تم حفظ {len(lines)} بند في المشتريات."
            if added: msg += f"\nوتمت إضافة {added} صنف إلى المخزون كوارد."
            messagebox.showinfo("تم الحفظ", msg)
            self.go("purchases")
        inner = tk.Frame(bx, bg=BG); inner.pack()
        self.btn(inner, "💾 حفظ المشتريات", save, GREEN).pack(side="right", padx=6)
        self.btn(inner, "حذف الصنف المحدد", del_line, RED).pack(side="right", padx=6)
        self.btn(inner, "إلغاء", win.destroy, MUTED).pack(side="right", padx=6)

    # ---------- التقارير ----------
    def show_reports(self):
        self.page_title("استخراج وطباعة التقارير", "حدد الفترة ونوع التقرير، راجع الأعداد ثم استخرج ملف Excel أو اطبعه")
        panel = self.panel(); panel.pack(fill="x", padx=24, pady=8)
        self.rep_from = tk.StringVar(value=today_text()); self.rep_to = tk.StringVar(value=today_text())
        self.add_date(panel, "من تاريخ", self.rep_from, 0, 1)
        self.add_date(panel, "إلى تاريخ", self.rep_to, 0, 0)
        quick = tk.Frame(panel, bg=WHITE); quick.grid(row=0, column=2, padx=14, sticky="e")
        for label, fn in [("اليوم", self.q_today), ("هذا الشهر", self.q_month), ("هذه السنة", self.q_year)]:
            self.btn(quick, label, fn, BLUE).pack(side="right", padx=3)
        self.report_type = tk.StringVar(value="all")
        box = tk.Frame(panel, bg=WHITE); box.grid(row=1, column=0, columnspan=3, sticky="e", padx=14, pady=8)
        opts = [("all", "كل السجلات"), ("visits", "الزيارات"), ("maintenance", "الصيانة"),
                ("cleaning", "جرد النظافة"), ("paper", "الجرد الورقي"), ("toner", "جرد الأحبار"),
                ("stockin", "الوارد للمخازن"), ("employees", "ملاحظات وأخطاء العاملين"), ("permits", "إذونات الخروج"),
                ("attendance", "الحضور والانصراف"), ("absences", "الغياب والتأخير"), ("govatt", "حضور موظفي الجهات"),
                ("govstaff", "قاعدة بيانات موظفي الجهات"), ("workers", "قاعدة العاملين")]
        need_pg = {"visits": ("visits",), "maintenance": ("maintenance",), "cleaning": ("cleaning",), "paper": ("paper",),
                   "toner": ("toner",), "stockin": ("cleaning", "paper", "toner"), "employees": ("employees",),
                   "permits": ("attend",), "attendance": ("attend",), "absences": ("attend",),
                   "govatt": ("gov",), "govstaff": ("govstaff",), "workers": ("workers",)}
        opts = [o for o in opts if (o[0] == "all" and self.all_data_visible())
                or (o[0] != "all" and any(self.can(pg, "view") for pg in need_pg.get(o[0], ())))]
        self.report_type.set(opts[0][0] if opts else "none")
        tk.Label(box, text="نوع التقرير", bg=WHITE, fg=TEXT, font=(FONT, 10, "bold")).grid(row=0, column=5, padx=8, sticky="e")
        for i, (val, label) in enumerate(opts):
            ttk.Radiobutton(box, text=label, value=val, variable=self.report_type).grid(
                row=i // 5, column=4 - i % 5, padx=8, pady=2, sticky="e")
        bx = tk.Frame(panel, bg=WHITE); bx.grid(row=2, column=0, columnspan=3, pady=12)
        self.btn(bx, "معاينة الأعداد", self.preview_counts, BLUE).pack(side="right", padx=6)
        self.btn(bx, "استخراج التقرير إلى Excel", self.export_report, GREEN).pack(side="right", padx=6)
        self.btn(bx, "🖨️ طباعة التقرير", self.print_report, PURPLE).pack(side="right", padx=6)

        self.preview_box = self.panel(); self.preview_box.pack(fill="both", expand=True, padx=24, pady=10)
        tk.Label(self.preview_box, text="محتويات التقرير", bg=WHITE, fg=INK, font=(FONT, 14, "bold"),
                 anchor="e").pack(fill="x", padx=20, pady=(16, 6))
        for line in ["• الزيارات: اسم الشخص، الجهة، سبب الزيارة، التاريخ والوقت.",
                     "• الصيانة: الفني أو المسؤول، الشركة، ملاحظات الصيانة، التاريخ والوقت.",
                     "• الجرد: الصنف، الكمية، الرصيد السابق، الوارد، والاستهلاك (السابق + الوارد - الجرد الجديد). والوارد للمخازن له تقرير مستقل.",
                     "• ملاحظات وأخطاء العاملين: تفاصيل كل بيان + ورقة ملخص بعدد ملاحظات وأخطاء كل عامل.",
                     "• إذونات الخروج والغياب والتأخير: تفاصيل كل بيان + ملخص بعدد إذونات وغياب وتأخير كل عامل.",
                     "• حضور موظفي الجهات: تفاصيل كل دخول وانصراف + ملخص لكل موظف + ملخص لكل جهة (للفترة المحددة).",
                     "• قاعدة بيانات موظفي الجهات: كل موظفي الجهات الثماني بالوظيفة والهاتف والرقم القومي والحالة.",
                     "• كل تقرير مكتوب فيه: خدمات مصر - فرع الإسكندرية.",
                     "• 🖨️ الطباعة: تفتح صفحة منسقة بلوجو خدمات مصر وخانات التوقيع وتبدأ نافذة الطباعة (ويمكن حفظها PDF).",
                     "• اضغط «معاينة الأعداد» لترى عدد السجلات في كل قسم قبل الاستخراج."]:
            tk.Label(self.preview_box, text=line, bg=WHITE, fg=TEXT, font=(FONT, 10), anchor="e",
                     justify="right").pack(fill="x", padx=22, pady=3)
        self.counts_lbl = tk.Label(self.preview_box, text="", bg=WHITE, fg=INK, font=(FONT, 11, "bold"),
                                   anchor="e", justify="right")
        self.counts_lbl.pack(fill="x", padx=22, pady=(10, 6))

    # ---------- الإذونات والغياب ----------
    def show_attendance(self):
        starts = "، ".join(f"{d} {fmt_time(t)}" for d, t in DEPT_START.items())
        self.page_title("🕒 الحضور والانصراف والإذونات والغياب",
                        f"إذن الخروج: {PERMIT_LIMIT} فقط لكل عامل خلال الشهر   |   بداية الدوام: {starts}")
        rp = self.panel(); rp.pack(fill="x", padx=24, pady=(4, 2))
        self.att_from = tk.StringVar(value=date.today().replace(day=1).strftime("%d/%m/%Y"))
        self.att_to = tk.StringVar(value=today_text())
        self.att_emp = tk.StringVar(value="(كل العاملين)")
        self.att_kind = tk.StringVar(value="attendance")
        self.add_date(rp, "من تاريخ", self.att_from, 0, 3)
        self.add_date(rp, "إلى تاريخ", self.att_to, 0, 2)
        self.add_combo(rp, "العامل", self.att_emp, ["(كل العاملين)"] + self.worker_names(), 0, 1)
        kb = tk.Frame(rp, bg=WHITE); kb.grid(row=1, column=0, columnspan=4, sticky="e", padx=14, pady=(0, 10))
        tk.Label(kb, text="نوع التقرير:", bg=WHITE, fg=TEXT, font=(FONT, 10, "bold")).pack(side="right", padx=6)
        for val, label in [("attendance", "الحضور والانصراف"), ("permits", "إذونات الخروج"),
                           ("absences", "الغياب والتأخير"), ("attall", "الكل")]:
            ttk.Radiobutton(kb, text=label, value=val, variable=self.att_kind).pack(side="right", padx=7)
        self.btn(kb, "🖨️ طباعة تقرير", self.print_attendance_report, AMBER).pack(side="right", padx=(18, 4))
        self.btn(kb, "📊 استخراج Excel", self.export_attendance_report, GREEN).pack(side="right", padx=4)
        kb2 = tk.Frame(rp, bg=WHITE); kb2.grid(row=2, column=0, columnspan=4, sticky="e", padx=14, pady=(0, 10))
        tk.Label(kb2, text="ملف بيانات الحضور والانصراف (للنسخ أو التعديل ثم الإدراج):", bg=WHITE, fg=TEXT,
                 font=(FONT, 10, "bold")).pack(side="right", padx=6)
        self.btn(kb2, "📤 استخراج بيانات", self.export_att_data, PRIMARY).pack(side="right", padx=4)
        self.btn(kb2, "📥 إدراج بيانات", self.import_att_data, PURPLE).pack(side="right", padx=4)
        nb = ttk.Notebook(self.content); nb.pack(fill="both", expand=True, padx=24, pady=(2, 8))
        tw = tk.Frame(nb, bg=BG); tp = tk.Frame(nb, bg=BG); ta = tk.Frame(nb, bg=BG)
        nb.add(tw, text="  ⏱️ الحضور والانصراف  ")
        nb.add(tp, text="  🚪 إذونات الخروج  "); nb.add(ta, text="  📅 الغياب والتأخير  ")
        self.build_clock_tab(tw); self.build_permits_tab(tp); self.build_absence_tab(ta)

    # ---------- قاعدة بيانات موظفي الجهات ----------
    def show_govstaff(self):
        self.page_title("🗂️ قاعدة بيانات موظفي الجهات",
                        "موظفو الجهات الحكومية الثماني: الاسم والوظيفة والهاتف والرقم القومي والحالة (أضفهم هنا أو مباشرة من صفحة الحضور)")
        form = self.panel(); form.pack(fill="x", padx=24, pady=(4, 2))
        self.gs_ent = tk.StringVar(value=GOV_ENTITIES[0]); self.gs_name = tk.StringVar(); self.gs_job = tk.StringVar()
        self.gs_phone = tk.StringVar(); self.gs_nid = tk.StringVar(); self.gs_stat = tk.StringVar(value=STATUS_ON)
        self.gs_notes = tk.StringVar()
        self.add_combo(form, "الجهة", self.gs_ent, GOV_ENTITIES, 0, 2)
        self.add_field(form, "اسم الموظف", self.gs_name, 0, 1, 260)
        self.add_field(form, "الوظيفة", self.gs_job, 0, 0, 220)
        self.add_field(form, "الهاتف", self.gs_phone, 1, 2, 200)
        self.add_field(form, "الرقم القومي (14 رقم)", self.gs_nid, 1, 1, 220)
        self.add_combo(form, "الحالة", self.gs_stat, [STATUS_ON, STATUS_OFF], 1, 0)
        self.add_field(form, "ملاحظات", self.gs_notes, 2, 2, 260)
        bx = tk.Frame(form, bg=WHITE); bx.grid(row=3, column=0, columnspan=3, sticky="e", padx=14, pady=(0, 10))
        self.btn(bx, "💾 حفظ الموظف", self.save_govstaff, GREEN).pack(side="right", padx=4)
        self.btn(bx, "✏️ تعديل المحدد", lambda: self.edit_record("gov_staff", self.gs_tree), BLUE).pack(side="right", padx=4)
        self.btn(bx, "حذف المحدد", lambda: self.delete_selected("gov_staff", self.gs_tree), RED).pack(side="right", padx=4)
        self.btn(bx, "🖨️ طباعة", self.print_govstaff, AMBER).pack(side="right", padx=(18, 4))
        self.btn(bx, "📊 استخراج Excel", self.export_govstaff, GREEN).pack(side="right", padx=4)
        fp = self.panel(); fp.pack(fill="x", padx=24, pady=(2, 2))
        self.gs_f_ent = tk.StringVar(value="(كل الجهات)"); self.gs_f_stat = tk.StringVar(value="الكل"); self.gs_f_key = tk.StringVar()
        self.add_combo(fp, "عرض الجهة", self.gs_f_ent, ["(كل الجهات)"] + GOV_ENTITIES, 0, 2)
        self.add_combo(fp, "الحالة", self.gs_f_stat, ["الكل", STATUS_ON, STATUS_OFF], 0, 1)
        self.add_field(fp, "بحث (اسم / هاتف / رقم قومي)", self.gs_f_key, 0, 0, 260)
        for v in (self.gs_f_ent, self.gs_f_stat, self.gs_f_key):
            v.trace_add("write", lambda *_: self.govstaff_reload())
        self.gs_tree = self.make_tree(self.content, ("ent", "name", "job", "phone", "nid", "stat", "notes"),
                                      ("الجهة", "اسم الموظف", "الوظيفة", "الهاتف", "الرقم القومي", "الحالة", "ملاحظات"),
                                      (130, 190, 120, 100, 130, 90, 160), 8)
        self.gs_tree.tag_configure("off", foreground=MUTED)
        self.gs_tree.bind("<Double-1>", lambda e: self.edit_record("gov_staff", self.gs_tree))
        self.govstaff_reload()

    def _govstaff_query(self):
        ent, st, key = self.gs_f_ent.get(), self.gs_f_stat.get(), self.gs_f_key.get().strip()
        sql = "SELECT id, entity, name, job_title, phone, national_id, status, notes FROM gov_staff WHERE 1=1"
        args = []
        if not ent.startswith("("):
            sql += " AND entity=?"; args.append(ent)
        if st != "الكل":
            sql += " AND status=?"; args.append(st)
        if key:
            sql += " AND (name LIKE ? OR phone LIKE ? OR national_id LIKE ?)"; args += [f"%{key}%"] * 3
        with db() as con:
            return con.execute(sql + " ORDER BY entity, (status=?), name", args + [STATUS_OFF]).fetchall()

    def govstaff_reload(self):
        if not hasattr(self, "gs_tree") or not self.gs_tree.winfo_exists():
            return
        self.gs_tree.delete(*self.gs_tree.get_children())
        for n, (rid, ent, nm, jb, ph, nid, st, nt) in enumerate(self._govstaff_query()):
            tags = ["odd"] if n % 2 else []
            if st == STATUS_OFF: tags.append("off")
            self.gs_tree.insert("", "end", iid=str(rid), tags=tags,
                                values=(ent, nm, jb or "-", ph or "-", nid or "-", st or STATUS_ON, nt or "-"))

    def save_govstaff(self):
        ent, name = self.gs_ent.get(), " ".join(self.gs_name.get().split())
        nid = self.gs_nid.get().strip().translate(_AR_DIGITS)
        if not name:
            messagebox.showwarning("بيانات ناقصة", "اكتب اسم الموظف."); return
        if nid and not (nid.isdigit() and len(nid) == 14):
            messagebox.showwarning("الرقم القومي", "الرقم القومي لازم يكون 14 رقمًا."); return
        with db() as con:
            if con.execute("SELECT 1 FROM gov_staff WHERE entity=? AND name=?", (ent, name)).fetchone():
                messagebox.showwarning("مكرر", f"«{name}» مسجل بالفعل في {ent}."); return
            con.execute("INSERT INTO gov_staff(entity,name,job_title,phone,national_id,status,notes,created_at) VALUES(?,?,?,?,?,?,?,?)",
                        (ent, name, self.gs_job.get().strip(), self.gs_phone.get().strip(), nid, self.gs_stat.get() or STATUS_ON,
                         self.gs_notes.get().strip(), now_stamp()))
        for v in (self.gs_name, self.gs_job, self.gs_phone, self.gs_nid, self.gs_notes): v.set("")
        self.gs_stat.set(STATUS_ON)
        self.govstaff_reload()

    def _govstaff_sections(self):
        rows = self._govstaff_query()
        if not rows:
            messagebox.showinfo("لا توجد بيانات", "لا يوجد موظفون بهذه الفلاتر."); return None
        flt = [x for x in (self.gs_f_ent.get() if not self.gs_f_ent.get().startswith("(") else "",
                           self.gs_f_stat.get() if self.gs_f_stat.get() != "الكل" else "",
                           self.gs_f_key.get().strip()) if x]
        period = f"تاريخ الاستخراج: {today_text()}" + ("   |   " + "   |   ".join(flt) if flt else "")
        secs = []
        for ent in GOV_ENTITIES:
            rs = [(nm, jb, ph, nid, st, nt) for _, e, nm, jb, ph, nid, st, nt in rows if e == ent]
            if rs:
                secs.append({"name": ent, "title": f"موظفو {ent}", "period": period,
                             "headers": ["اسم الموظف", "الوظيفة", "الهاتف", "الرقم القومي", "الحالة", "ملاحظات"],
                             "widths": [30, 20, 16, 20, 14, 34], "rows": rs})
        return secs

    def print_govstaff(self):
        secs = self._govstaff_sections()
        if not secs: return
        try:
            open_html(sections_html("قاعدة بيانات موظفي الجهات", secs), "govstaff")
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تجهيز الطباعة:\n{ex}")

    def export_govstaff(self):
        secs = self._govstaff_sections()
        if not secs: return
        self.export_sections_excel(secs, f"موظفو_الجهات_{date.today():%Y-%m-%d}.xlsx", "قاعدة بيانات موظفي الجهات")

    # ---------- حضور وانصراف الجهات ----------
    def show_gov(self):
        self.page_title("🏛️ حضور وانصراف الجهات",
                        "اضغط على اسم الجهة فتفتح قائمة موظفيها، اختر الموظف وسجّل دخول أو خروج أو ملاحظة (ولو موظف جديد اكتب اسمه)")
        self.gov_entity = tk.StringVar(value=getattr(self, "_last_gov", GOV_ENTITIES[0]))
        chips = tk.Frame(self.content, bg=BG); chips.pack(fill="x", padx=24, pady=(0, 4))
        self.gov_chips = {}
        for i, ent in enumerate(GOV_ENTITIES):
            r, c = divmod(i, 4)
            b = tk.Button(chips, text=ent, command=lambda e=ent: self.gov_select(e, True), relief="flat", bd=0,
                          font=(FONT, 11, "bold"), padx=10, pady=7, cursor="hand2")
            b.grid(row=r, column=3 - c, sticky="nsew", padx=4, pady=3)
            chips.grid_columnconfigure(c, weight=1, uniform="gc")
            self.gov_chips[ent] = b
        bar = self.panel(); bar.pack(fill="x", padx=24, pady=(2, 2))
        top = tk.Frame(bar, bg=WHITE); top.pack(fill="x", padx=12, pady=(8, 2))
        self.g_title = tk.Label(top, text="", bg=WHITE, fg=BLUE, font=(FONT, 12, "bold")); self.g_title.pack(side="right")
        self.g_view = tk.StringVar(value="all")
        for val, label in (("today", "اليوم فقط"), ("all", "كل السجلات")):
            ttk.Radiobutton(top, text=label, value=val, variable=self.g_view, command=self.gov_reload).pack(side="left", padx=6)
        bx = tk.Frame(bar, bg=WHITE); bx.pack(fill="x", padx=12, pady=(2, 10))
        self.btn(bx, "🖨️ طباعة تقارير", self.gov_report_dialog, AMBER).pack(side="right", padx=(0, 4))
        self.btn(bx, "👥 اختيار موظف وتسجيل", self.gov_picker, GREEN).pack(side="right", padx=4)
        self.btn(bx, "✏️ تعديل المحدد", lambda: self.edit_record("gov_attendance", self.g_tree), BLUE).pack(side="right", padx=4)
        self.btn(bx, "حذف المحدد", lambda: self.delete_selected("gov_attendance", self.g_tree), RED).pack(side="right", padx=4)
        self.btn(bx, "📤 استخراج بيانات", self.export_gov_att, PRIMARY).pack(side="right", padx=(18, 4))
        self.btn(bx, "📥 إدراج بيانات", self.import_gov_att, PURPLE).pack(side="right", padx=4)
        self.g_tree = self.make_tree(self.content, ("person", "date", "in", "out", "hrs", "notes"),
                                     ("اسم الموظف", "التاريخ", "الدخول", "الخروج", "ساعات العمل", "ملاحظات"),
                                     (200, 100, 95, 95, 100, 260), 10)
        self.g_tree.bind("<Double-1>", lambda e: self.edit_record("gov_attendance", self.g_tree))
        self.gov_select(self.gov_entity.get())

    def gov_select(self, ent, open_picker=False):
        self.gov_entity.set(ent); self._last_gov = ent
        for e, b in self.gov_chips.items():
            on = e == ent
            b.config(bg=PRIMARY if on else SOFT, fg="white" if on else TEXT, activebackground=BLUE if on else LINE)
        self.g_title.config(text=f"سجل حضور وانصراف: {ent}")
        self.gov_reload()
        if open_picker and self.can("gov", "add"):
            self.gov_picker()

    def gov_reload(self):
        if not hasattr(self, "g_tree") or not self.g_tree.winfo_exists():
            return
        ent = self.gov_entity.get()
        sql = "SELECT id, person, att_date, in_time, out_time, notes FROM gov_attendance WHERE entity=?"
        args = [ent]
        if self.g_view.get() == "today":
            sql += " AND att_date=?"; args.append(date.today().isoformat())
        with db() as con:
            att = con.execute(sql + " ORDER BY att_date DESC, id DESC LIMIT 400", args).fetchall()
        self.g_tree.delete(*self.g_tree.get_children())
        for n, (rid, pn, d, i, o, nt) in enumerate(att):
            w = _mins(o) - _mins(i) if i and o else None
            self.g_tree.insert("", "end", iid=str(rid), tags=("odd",) if n % 2 else (),
                               values=(pn, fmt_iso_date(d), fmt_time(i) or "-", fmt_time(o) or "-", fmt_hours(w), nt or "-"))

    def _gov_report_names(self, entity=""):
        """أسماء موظفي الجهات (من قاعدة الموظفين ومن سجلات الحضور) لاختيار موظف في تقرير الحضور."""
        sql1 = "SELECT DISTINCT person FROM gov_attendance WHERE person <> ''"
        sql2 = "SELECT DISTINCT name FROM gov_staff WHERE name <> ''"
        args = []
        if entity and not entity.startswith("("):
            sql1 += " AND entity=?"
            sql2 += " AND entity=?"
            args = [entity]
        with db() as con:
            names = {r[0] for r in con.execute(sql1, args) if r[0]}
            names |= {r[0] for r in con.execute(sql2, args) if r[0]}
        return sorted(names)

    def gov_report_dialog(self):
        """تقرير حضور وانصراف موظفي الجهات: بالمدة / بالجهة / بالاسم — طباعة أو Excel."""
        win = self._dialog("طباعة تقارير حضور الجهات", "900x700")
        self._minsize(win, 780, 600)
        tk.Label(win, text="🖨️ تقارير حضور وانصراف موظفي الجهات", bg=BG, fg=INK,
                 font=(FONT, 16, "bold")).pack(anchor="e", padx=22, pady=(14, 2))
        tk.Label(win, text="حدد المدة، واختر جهة أو موظفًا (أو اترك الكل)، ثم اطبع أو استخرج Excel.",
                 bg=BG, fg=MUTED, font=(FONT, 9)).pack(anchor="e", padx=24)
        form = self.panel(win)
        form.pack(fill="x", padx=16, pady=10)
        f_var = tk.StringVar(value=earliest_record_date("gov_attendance").strftime("%d/%m/%Y"))
        t_var = tk.StringVar(value=today_text())
        self.add_date(form, "من تاريخ", f_var, 0, 2)
        self.add_date(form, "إلى تاريخ", t_var, 0, 1)
        quick = tk.Frame(form, bg=WHITE)
        quick.grid(row=0, column=0, padx=10, sticky="e")
        for label, a, b in [("هذا الشهر", date.today().replace(day=1).strftime("%d/%m/%Y"), today_text()),
                            ("هذه السنة", date.today().replace(month=1, day=1).strftime("%d/%m/%Y"), today_text())]:
            self.btn(quick, label, lambda a=a, b=b: (f_var.set(a), t_var.set(b)), BLUE).pack(side="right", padx=3)
        ent_var = tk.StringVar(value="(كل الجهات)")
        per_var = tk.StringVar(value="(كل الموظفين)")
        ent_cb = self.add_combo(form, "الجهة", ent_var, ["(كل الجهات)"] + GOV_ENTITIES, 1, 2)
        per_cb = self.add_combo(form, "الموظف", per_var, ["(كل الموظفين)"] + self._gov_report_names(), 1, 1)
        tk.Label(form, text="اترك «كل الجهات» و«كل الموظفين» لعمل تقرير شامل.", bg=WHITE, fg=MUTED,
                 font=(FONT, 9)).grid(row=1, column=0, padx=12, sticky="e")

        def refresh_names(*_):
            vals = ["(كل الموظفين)"] + self._gov_report_names(ent_var.get())
            per_cb.configure(values=vals)
            if per_var.get() not in vals:
                per_var.set(vals[0])
        ent_cb.bind("<<ComboboxSelected>>", refresh_names)

        kind = tk.StringVar(value="all")
        kb = tk.Frame(form, bg=WHITE)
        kb.grid(row=2, column=0, columnspan=3, sticky="e", padx=12, pady=(0, 10))
        tk.Label(kb, text="محتوى التقرير:", bg=WHITE, fg=TEXT, font=(FONT, 10, "bold")).pack(side="right", padx=6)
        for val, label in [("details", "تفاصيل الحضور"), ("summary", "ملخص بالموظف والجهة"), ("all", "التفاصيل + الملخص")]:
            ttk.Radiobutton(kb, text=label, value=val, variable=kind).pack(side="right", padx=7)

        bb = tk.Frame(win, bg=BG)
        bb.pack(side="bottom", fill="x", pady=12)
        tk.Label(win, text="التقرير يطبع على ورق A4 منسق بلوجو خدمات مصر وخانات التوقيع، ويمكن حفظه PDF.",
                 bg=BG, fg=MUTED, font=(FONT, 9)).pack(side="bottom", pady=(0, 4))

        def collect():
            try:
                s = parse_date(f_var.get()).isoformat()
                e = parse_date(t_var.get()).isoformat()
                if e < s:
                    raise ValueError
            except ValueError:
                messagebox.showerror("التاريخ", "راجع تاريخ البداية والنهاية (يوم/شهر/سنة).", parent=win)
                return None
            ent = ent_var.get()
            ent = "" if ent.startswith("(") else ent
            per = per_var.get()
            per = "" if per.startswith("(") else per
            secs = gov_att_sections(s, e, ent, per)
            if kind.get() == "details":
                secs = secs[:1]
            elif kind.get() == "summary":
                secs = secs[1:]
            if not any(x["rows"] for x in secs):
                messagebox.showinfo("لا توجد بيانات", "لا توجد سجلات حضور في المدة/الاختيار المحدد.", parent=win)
                return None
            return secs, ent, per

        def do_print():
            got = collect()
            if not got:
                return
            secs, ent, per = got
            win.destroy()
            try:
                open_html(sections_html("تقرير حضور وانصراف موظفي الجهات", secs), "gov_report")
            except Exception as ex:
                messagebox.showerror("خطأ", f"تعذر تجهيز الطباعة:\n{ex}")

        def do_excel():
            got = collect()
            if not got:
                return
            secs, ent, per = got
            tag = ent or "كل_الجهات"
            if per:
                tag = f"{per}_{tag}"
            name = f"حضور_الجهات_{tag}_{date.today():%Y-%m-%d}.xlsx".replace(" ", "_")
            win.destroy()
            self.export_sections_excel(secs, name, "تقرير حضور موظفي الجهات")

        self.btn(bb, "🖨️ طباعة / PDF", do_print, PURPLE).pack(side="right", padx=6)
        self.btn(bb, "📊 استخراج Excel", do_excel, GREEN).pack(side="right", padx=6)
        self.btn(bb, "إلغاء", win.destroy, MUTED).pack(side="right", padx=6)

    def gov_picker(self):
        """قائمة موظفي الجهة المختارة: اختر الموظف ثم سجّل دخول / خروج / ملاحظة، أو أضف موظفًا جديدًا."""
        ent = self.gov_entity.get()
        win = self._dialog(f"موظفو {ent}", "720x640")
        self._minsize(win, 620, 500)
        tk.Label(win, text=f"🏛️ {ent}", bg=BG, fg=INK, font=(FONT, 17, "bold")).pack(anchor="e", padx=22, pady=(12, 0))
        tk.Label(win, text="اضغط مرتين على الموظف (أو اختره ثم «اختيار وتسجيل»). لو موظف جديد اكتب اسمه في الأسفل.",
                 bg=BG, fg=MUTED, font=(FONT, 9)).pack(anchor="e", padx=24, pady=(0, 6))
        sv = tk.StringVar()
        sr = tk.Frame(win, bg=BG); sr.pack(fill="x", padx=22, pady=(0, 4))
        se = ttk.Entry(sr, textvariable=sv, justify="right"); se.pack(side="right", fill="x", expand=True)
        tk.Label(sr, text="🔍 بحث:", bg=BG, fg=TEXT, font=(FONT, 10, "bold")).pack(side="right", padx=6)
        bx = tk.Frame(win, bg=BG); bx.pack(side="bottom", fill="x", pady=8)
        newp = self.panel(win); newp.pack(side="bottom", fill="x", padx=16, pady=(0, 4))
        tree = self.make_tree(win, ("name", "job", "today"), ("اسم الموظف", "الوظيفة", "حالة اليوم"), (270, 170, 190), 12)
        tree.tag_configure("in", foreground=GREEN); tree.tag_configure("done", foreground=MUTED)

        def load(*_):
            key = sv.get().strip(); iso = date.today().isoformat()
            with db() as con:
                staff = con.execute("SELECT name, job_title FROM gov_staff WHERE entity=? AND status<>? AND (?='' OR name LIKE ?) ORDER BY name",
                                    (ent, STATUS_OFF, key, f"%{key}%")).fetchall()
                today = {r[0]: (r[1], r[2]) for r in con.execute(
                    "SELECT person, in_time, out_time FROM gov_attendance WHERE entity=? AND att_date=?", (ent, iso))}
            tree.delete(*tree.get_children())
            for n, (nm, jb) in enumerate(staff):
                i, o = today.get(nm, ("", ""))
                if i and o: st, tag = f"دخل {fmt_time(i)} - خرج {fmt_time(o)}", "done"
                elif i: st, tag = f"دخل {fmt_time(i)}", "in"
                elif o: st, tag = f"خرج {fmt_time(o)}", "in"
                else: st, tag = "لم يُسجَّل اليوم", ""
                tree.insert("", "end", iid=str(n), tags=(tag,) if tag else (), values=(nm, jb or "-", st))
        sv.trace_add("write", load); load()

        def pick(_=None):
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("اختر موظفًا", "اختر الموظف من القائمة أولًا.", parent=win); return
            self.gov_action_dialog(ent, tree.item(sel[0], "values")[0], lambda: (load(), self.gov_reload()))

        nv, jv = tk.StringVar(), tk.StringVar()
        tk.Label(newp, text="➕ موظف جديد:", bg=WHITE, fg=TEXT, font=(FONT, 10, "bold")).pack(side="right", padx=8, pady=10)
        ttk.Entry(newp, textvariable=nv, width=26, justify="right").pack(side="right", padx=4)
        tk.Label(newp, text="الوظيفة:", bg=WHITE, fg=MUTED, font=(FONT, 9)).pack(side="right", padx=(8, 2))
        ttk.Entry(newp, textvariable=jv, width=14, justify="right").pack(side="right", padx=4)

        def add_new():
            name = " ".join(nv.get().split())
            if not name:
                messagebox.showwarning("بيانات ناقصة", "اكتب اسم الموظف الجديد.", parent=win); return
            with db() as con:
                ex = con.execute("SELECT status FROM gov_staff WHERE entity=? AND name=?", (ent, name)).fetchone()
                if ex is None:
                    con.execute("INSERT INTO gov_staff(entity,name,job_title,status,created_at) VALUES(?,?,?,?,?)",
                                (ent, name, jv.get().strip(), STATUS_ON, now_stamp()))
                elif ex[0] == STATUS_OFF:
                    if not messagebox.askyesno("موظف سابق", f"«{name}» مسجل بحالة «ترك العمل». هل تريد إعادته على رأس العمل؟", parent=win):
                        return
                    con.execute("UPDATE gov_staff SET status=? WHERE entity=? AND name=?", (STATUS_ON, ent, name))
            nv.set(""); jv.set(""); load()
            self.gov_action_dialog(ent, name, lambda: (load(), self.gov_reload()))
        self.btn(newp, "إضافة وتسجيل", add_new, BLUE).pack(side="right", padx=8)
        tree.bind("<Double-1>", pick); tree.bind("<Return>", pick)
        inner = tk.Frame(bx, bg=BG); inner.pack()
        self.btn(inner, "✅ اختيار وتسجيل", pick, GREEN).pack(side="right", padx=6)
        self.btn(inner, "إغلاق", win.destroy, MUTED).pack(side="right", padx=6)
        se.focus_set()

    def gov_action_dialog(self, ent, name, refresh=None):
        """تسجيل دخول / خروج / ملاحظة لموظف جهة (الوقت الحالي افتراضيًا ويمكن تعديله)."""
        win = self._dialog(f"تسجيل - {name}", "540x470")
        tk.Label(win, text=f"👤 {name}", bg=BG, fg=INK, font=(FONT, 17, "bold")).pack(anchor="e", padx=22, pady=(14, 0))
        tk.Label(win, text=ent, bg=BG, fg=MUTED, font=(FONT, 10)).pack(anchor="e", padx=24)
        info = tk.Label(win, text="", bg=BG, fg=BLUE, font=(FONT, 10, "bold"), wraplength=480, justify="right")
        info.pack(anchor="e", padx=24, pady=(6, 0))
        dv, tv, nv = tk.StringVar(value=today_text()), tk.StringVar(value=datetime.now().strftime("%H:%M")), tk.StringVar()
        box = self.panel(win); box.pack(fill="x", padx=16, pady=8)
        self.add_date(box, "التاريخ", dv, 0, 1)
        self.add_field(box, "الوقت (مثل 8:30 أو 3:00 م)", tv, 0, 0, 200)
        self.add_field(box, "ملاحظات (اختياري)", nv, 1, 1, 330)

        def rec():
            try:
                iso = parse_date(dv.get()).isoformat()
            except ValueError:
                return None, None
            with db() as con:
                r = con.execute("SELECT id,in_time,out_time,notes FROM gov_attendance WHERE entity=? AND person=? AND att_date=?",
                                (ent, name, iso)).fetchone()
            return iso, r

        def upd(*_):
            iso, r = rec()
            if iso is None: info.config(text="", fg=MUTED); return
            if not r: info.config(text="لا يوجد تسجيل لهذا اليوم بعد.", fg=MUTED); return
            bits = []
            if r[1]: bits.append(f"الدخول {fmt_time(r[1])}")
            if r[2]: bits.append(f"الخروج {fmt_time(r[2])}")
            if r[3]: bits.append(f"ملاحظات: {r[3]}")
            info.config(text="مسجل لهذا اليوم: " + "  |  ".join(bits or ["بدون أوقات"]), fg=BLUE)
        dv.trace_add("write", upd); upd()

        def do(kind):
            iso, r = rec()
            if iso is None:
                messagebox.showerror("التاريخ", "راجع التاريخ.", parent=win); return
            note, t = nv.get().strip(), ""
            if kind != "note":
                try:
                    t = parse_time_flex(tv.get(), kind == "out", (r[1] if r else "") or None)
                except ValueError as ex:
                    messagebox.showerror("الوقت", str(ex), parent=win); return
                if not t:
                    messagebox.showwarning("الوقت", "اكتب الوقت.", parent=win); return
            elif not note:
                messagebox.showwarning("ملاحظة", "اكتب الملاحظة أولًا.", parent=win); return
            merged = " | ".join(x for x in ((r[3] if r else ""), note) if x)
            if kind == "in":
                if r and r[1]:
                    messagebox.showwarning("مسجل بالفعل", f"تم تسجيل دخول {name} الساعة {fmt_time(r[1])} في هذا اليوم.\n"
                                           "لو عايز تغيّره عدّل السجل من الجدول.", parent=win); return
                if r and r[2] and t >= r[2]:
                    messagebox.showwarning("الوقت", "وقت الدخول لازم يكون قبل وقت الخروج المسجل.", parent=win); return
            elif kind == "out":
                if r and r[2]:
                    messagebox.showwarning("مسجل بالفعل", f"تم تسجيل خروج {name} الساعة {fmt_time(r[2])} في هذا اليوم.\n"
                                           "لو عايز تغيّره عدّل السجل من الجدول.", parent=win); return
                if r and r[1] and t <= r[1]:
                    messagebox.showwarning("الوقت", "وقت الخروج لازم يكون بعد وقت الدخول.", parent=win); return
            with db() as con:
                if r:
                    col = {"in": "in_time", "out": "out_time"}.get(kind)
                    if col:
                        con.execute(f"UPDATE gov_attendance SET {col}=?, notes=? WHERE id=?", (t, merged, r[0]))
                    else:
                        con.execute("UPDATE gov_attendance SET notes=? WHERE id=?", (merged, r[0]))
                else:
                    con.execute("INSERT INTO gov_attendance(entity,person,att_date,in_time,out_time,notes,created_at) VALUES(?,?,?,?,?,?,?)",
                                (ent, name, iso, t if kind == "in" else "", t if kind == "out" else "", note, now_stamp()))
                if not con.execute("SELECT 1 FROM gov_staff WHERE entity=? AND name=?", (ent, name)).fetchone():
                    con.execute("INSERT INTO gov_staff(entity,name,status,created_at) VALUES(?,?,?,?)", (ent, name, STATUS_ON, now_stamp()))
            win.destroy()
            if refresh: refresh()
            what = {"in": f"دخول الساعة {fmt_time(t)}", "out": f"خروج الساعة {fmt_time(t)}", "note": "الملاحظة"}[kind]
            messagebox.showinfo("تم التسجيل", f"تم تسجيل {what} لـ «{name}».")
        row = tk.Frame(win, bg=BG); row.pack(pady=10)
        self.btn(row, "✅ تسجيل دخول", lambda: do("in"), GREEN).pack(side="right", padx=5)
        self.btn(row, "🚪 تسجيل خروج", lambda: do("out"), ORANGE).pack(side="right", padx=5)
        self.btn(row, "📝 حفظ ملاحظة", lambda: do("note"), BLUE).pack(side="right", padx=5)
        self.btn(win, "إلغاء", win.destroy, MUTED).pack(pady=(0, 10))

    def export_gov_att(self):
        ent = self.gov_entity.get()
        scope = messagebox.askyesnocancel("استخراج بيانات", f"نعم = حضور وانصراف «{ent}» فقط\nلا = كل الجهات الثمانية\nإلغاء = رجوع")
        if scope is None:
            return
        p = self._ask_period("استخراج بيانات الحضور والانصراف", "gov_attendance")
        if not p:
            return
        sql, args = ("SELECT entity, person, att_date, in_time, out_time, notes FROM gov_attendance "
                     "WHERE att_date BETWEEN ? AND ?"), list(p)
        if scope:
            sql += " AND entity=?"; args.append(ent)
        with db() as con:
            data = con.execute(sql + " ORDER BY entity, att_date, person", args).fetchall()
        rows = [(e, n, date.fromisoformat(d), fmt_time(i), fmt_time(o), nt or "") for e, n, d, i, o, nt in data]
        self._write_data_file(f"حضور_الجهات_{(ent if scope else 'كل_الجهات')}_{date.today():%Y-%m-%d}.xlsx".replace(" ", "_"),
                              "الحضور والانصراف", ["الجهة", "اسم الموظف", "التاريخ", "وقت الحضور", "وقت الانصراف", "ملاحظات"],
                              rows, [24, 28, 15, 15, 15, 32], 2, "بيانات حضور الجهات", GOV_HELP)

    def import_gov_att(self):
        self._import_flow("gov", default_entity=self.gov_entity.get(),
                          after=lambda: self.gov_reload() if self.active_key == "gov" else None)

    def _att_sections(self):
        try:
            s = parse_date(self.att_from.get()).isoformat(); e = parse_date(self.att_to.get()).isoformat()
            if e < s: raise ValueError
        except ValueError:
            messagebox.showerror("التاريخ", "راجع تاريخ البداية والنهاية (يوم/شهر/سنة)."); return None
        kind = self.att_kind.get()
        emp = self.att_emp.get().strip()
        emp = "" if emp.startswith("(") else emp
        secs = filter_sections_by_employee(report_sections(kind, s, e), emp)
        if emp:
            for x in secs:
                x["period"] += f"   |   العامل: {emp}"
        return kind, secs

    def _write_data_file(self, default_name, sheet, headers, rows, widths, date_col, what, help_lines=None):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name,
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return False
        try:
            write_simple_xlsx(path, sheet, headers, rows, widths, date_col, help_lines or ATT_HELP)
        except PermissionError:
            messagebox.showerror("الملف مفتوح", "الملف مفتوح في Excel. أغلقه أو اختر اسمًا آخر ثم أعد المحاولة."); return False
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر إنشاء الملف:\n{ex}"); return False
        if messagebox.askyesno("تم الاستخراج", f"تم استخراج {what} ({len(rows)} سجل):\n{path}\n\nهل تريد فتح الملف الآن؟"):
            try: os.startfile(path)
            except Exception: pass
        return True

    def _import_flow(self, kind, default_entity=None, after=None):
        """اختيار ملف -> قراءة -> معاينة -> سؤال عن المكرر -> إدراج -> ملخص."""
        page = "attend" if kind == "main" else "gov"
        path = filedialog.askopenfilename(title="اختر ملف بيانات الحضور والانصراف",
                                          filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("كل الملفات", "*.*")])
        if not path:
            return
        try:
            rows = read_sheet_rows(path)
            recs, errs = parse_attendance_rows(rows)
        except PermissionError:
            messagebox.showerror("الملف مفتوح", "الملف مفتوح في Excel. أغلقه ثم أعد المحاولة."); return
        except Exception as ex:
            messagebox.showerror("تعذر قراءة الملف", str(ex)); return
        prev = apply_attendance_import(kind, recs, False, default_entity, dry=True)
        errs = sorted(errs + prev["errors"])
        if not prev["added"] and not prev["dups"]:
            messagebox.showwarning("لا توجد بيانات صالحة", "لم أجد أي سجل صالح للإدراج.\n\n" + self._errs_text(errs)); return
        msg = f"الملف: {os.path.basename(path)}\n\nسجلات جديدة سيتم إضافتها: {prev['added']}\n"
        if prev["dups"]:
            msg += f"سجلات موجودة بالفعل (نفس الاسم والتاريخ): {prev['dups']}\n"
        if errs:
            msg += f"صفوف بها أخطاء سيتم تجاهلها: {len(errs)}\n"
        update = False
        if prev["dups"]:
            ans = messagebox.askyesnocancel("إدراج بيانات", msg + "\nنعم = تحديث السجلات الموجودة بالبيانات الجديدة\n"
                                            "لا = تخطي الموجودة وإضافة الجديد فقط\nإلغاء = إيقاف العملية")
            if ans is None:
                return
            update = ans
            if update and not self.can(page, "edit"):
                update = False
                messagebox.showinfo("صلاحية", "تحديث السجلات الموجودة يحتاج صلاحية التعديل، لذلك سيتم تخطيها وإضافة الجديد فقط.")
        elif not messagebox.askyesno("إدراج بيانات", msg + "\nهل تريد المتابعة؟"):
            return
        try:
            res = apply_attendance_import(kind, recs, update, default_entity)
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر إدراج البيانات:\n{ex}"); return
        if after: after()
        out = f"تمت إضافة {res['added']} سجل"
        if res["updated"]: out += f"، وتحديث {res['updated']} سجل"
        if res["dups"]: out += f"، وتخطي {res['dups']} سجل موجود"
        messagebox.showinfo("تم الإدراج", out + ".\n\n" + (self._errs_text(errs) if errs else ""))

    @staticmethod
    def _errs_text(errs):
        if not errs:
            return ""
        lines = [f"صف {n}: {m}" for n, m in errs[:15]]
        if len(errs) > 15:
            lines.append(f"... و{len(errs) - 15} خطأ آخر")
        return "الصفوف التي لم تُدرج:\n" + "\n".join(lines)

    def export_att_data(self):
        try:
            s_ = parse_date(self.att_from.get()).isoformat(); e_ = parse_date(self.att_to.get()).isoformat()
            if e_ < s_: raise ValueError
        except ValueError:
            messagebox.showerror("التاريخ", "راجع تاريخ البداية والنهاية (يوم/شهر/سنة)."); return
        emp = self.att_emp.get().strip(); emp = "" if emp.startswith("(") else emp
        sql, args = "SELECT employee, att_date, in_time, out_time, notes FROM attendance WHERE att_date BETWEEN ? AND ?", [s_, e_]
        if emp:
            sql += " AND employee=?"; args.append(emp)
        with db() as con:
            data = con.execute(sql + " ORDER BY att_date, employee", args).fetchall()
        rows = [(n, date.fromisoformat(d), fmt_time(i), fmt_time(o), nt or "") for n, d, i, o, nt in data]
        self._write_data_file(f"بيانات_الحضور_والانصراف_{date.today():%Y-%m-%d}.xlsx", "الحضور والانصراف",
                              ["اسم العامل", "التاريخ", "وقت الحضور", "وقت الانصراف", "ملاحظات"], rows,
                              [28, 15, 15, 15, 32], 1, "بيانات الحضور والانصراف")

    def import_att_data(self):
        self._import_flow("main", after=lambda: self.load_clock() if self.active_key == "attend" else None)

    def print_attendance_report(self):
        r = self._att_sections()
        if not r: return
        kind, secs = r
        try:
            open_html(sections_html(ATT_TITLES[kind], secs), "attendance")
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تجهيز الطباعة:\n{ex}")

    def export_attendance_report(self):
        r = self._att_sections()
        if not r: return
        kind, secs = r
        self.export_sections_excel(secs, f"{ATT_TITLES[kind]}_{date.today():%Y-%m-%d}.xlsx", ATT_TITLES[kind])

    # ---------- الحضور والانصراف (كتابة يدوية) ----------
    def build_clock_tab(self, tab):
        names = self.worker_names()
        form = self.panel(tab); form.pack(fill="x", pady=4)
        self.c_emp = tk.StringVar(); self.c_date = tk.StringVar(value=today_text())
        self.c_in = tk.StringVar(); self.c_out = tk.StringVar(); self.c_notes = tk.StringVar()
        self.add_combo(form, "اسم العامل", self.c_emp, names, 0, 2)
        self.add_date(form, "التاريخ", self.c_date, 0, 1)
        info = tk.Label(form, text="", bg=WHITE, fg=BLUE, font=(FONT, 10, "bold"), justify="right", wraplength=320)
        info.grid(row=0, column=0, padx=14, sticky="e")
        self.add_field(form, "وقت الحضور (اكتبه يدويًا مثل 7:30)", self.c_in, 1, 2, 220)
        self.add_field(form, "وقت الانصراف (اكتبه يدويًا مثل 4:00 م)", self.c_out, 1, 1, 220)
        self.add_field(form, "ملاحظات", self.c_notes, 1, 0, 260)

        # ---- تسجيل سريع بضغطة واحدة: الوقت يُؤخذ من ساعة الجهاز ----
        quick = tk.Frame(form, bg=WHITE)
        quick.grid(row=2, column=0, columnspan=3, sticky="e", padx=14, pady=(4, 0))
        tk.Label(quick, text="تسجيل سريع بساعة الجهاز:", bg=WHITE, fg=MUTED,
                 font=(FONT, 9, "bold")).pack(side="right", padx=(0, 8))
        tk.Label(quick, text="⏱️ الوقت بيتاخد من ساعة الجهاز لحظة الضغط (والتاريخ المختار فوق)",
                 bg=WHITE, fg=MUTED, font=(FONT, 9)).pack(side="right", padx=6)
        self.btn(quick, "✅ حضور الآن", lambda: self.clock_now("in"), GREEN).pack(side="right", padx=4)
        self.btn(quick, "🚪 انصراف الآن", lambda: self.clock_now("out"), ORANGE).pack(side="right", padx=4)

        bx = tk.Frame(form, bg=WHITE); bx.grid(row=3, column=0, columnspan=3, sticky="e", padx=14, pady=(4, 10))
        self.btn(bx, "💾 حفظ الحضور والانصراف", self.save_clock, GREEN).pack(side="right", padx=4)
        self.btn(bx, "↺ المواعيد الرسمية للقسم", lambda: fill_default(True), BLUE).pack(side="right", padx=4)
        self.btn(bx, "حذف المحدد", self.delete_clock, RED).pack(side="right", padx=4)

        def fill_default(force=False):
            name = self.c_emp.get()
            if not name:
                return
            s0, s1 = shift_of(worker_dept(name))
            if force or not self.c_in.get(): self.c_in.set(fmt_time(s0))
            if force or not self.c_out.get(): self.c_out.set(fmt_time(s1))

        def today_status(name):
            if not name:
                return ""
            try:
                with db() as con:
                    r = con.execute("SELECT in_time, out_time FROM attendance WHERE employee=? AND att_date=?",
                                    (name, date.today().isoformat())).fetchone()
            except Exception:
                return ""
            if not r:
                return "   |   النهاردة: مفيش تسجيل"
            return f"   |   النهاردة: حضور {fmt_time(r[0]) or '—'} · انصراف {fmt_time(r[1]) or '—'}"

        def upd(*_):
            name = self.c_emp.get()
            dept = worker_dept(name) if name else ""
            if not dept:
                info.config(text="الأمن: 7ص - 7م   |   الفنيين: 7:30ص - 4م   |   الخدمات: 7ص - 4م" + today_status(name)); return
            s0, s1 = shift_of(dept)
            info.config(text=f"قسم {DEPT_LABEL.get(dept, dept)}: الدوام من {fmt_time(s0)} إلى {fmt_time(s1)}"
                             + today_status(name))

        def on_emp(*_):
            self.c_in.set(""); self.c_out.set(""); fill_default(); upd()

        self._clock_refresh = upd
        self.c_emp.trace_add("write", on_emp); upd()
        self.c_tree = self.make_tree(tab, ("name", "dept", "date", "in", "out", "hrs", "late", "early", "notes"),
                                     ("اسم العامل", "القسم", "التاريخ", "الحضور", "الانصراف", "ساعات العمل",
                                      "تأخير (دقيقة)", "انصراف مبكر (دقيقة)", "ملاحظات"),
                                     (150, 70, 95, 85, 85, 90, 90, 120, 160), 8)
        self.c_tree.tag_configure("warn", foreground=ORANGE)
        self.load_clock()

    def save_clock(self):
        name = self.c_emp.get().strip()
        if not name:
            messagebox.showwarning("بيانات ناقصة", "اختر العامل من القائمة أولًا."); return
        try:
            iso = parse_date(self.c_date.get()).isoformat()
        except ValueError:
            messagebox.showerror("التاريخ", "راجع التاريخ."); return
        try:
            in_t = parse_time_flex(self.c_in.get())
            out_t = parse_time_flex(self.c_out.get(), True, in_t or None)
        except ValueError as ex:
            messagebox.showerror("الوقت", str(ex)); return
        ok, msg = add_attendance(name, iso, in_t, out_t, self.c_notes.get().strip())
        if not ok:
            messagebox.showwarning("تنبيه", msg); return
        self.c_notes.set(""); self.c_in.set(""); self.c_out.set("")
        self.load_clock()
        if hasattr(self, "_clock_refresh"):
            self._clock_refresh()
        messagebox.showinfo("تم الحفظ", f"تم تسجيل حضور وانصراف «{name}».")

    def clock_now(self, kind):
        """تسجيل حضور/انصراف بضغطة واحدة: الوقت يُؤخذ من ساعة الجهاز لحظة الضغط."""
        name = self.c_emp.get().strip()
        if not name:
            messagebox.showwarning("اختر العامل", "اختر اسم العامل من القائمة أولًا، ثم اضغط الحضور أو الانصراف.")
            return
        try:
            iso = parse_date(self.c_date.get()).isoformat()
        except ValueError:
            messagebox.showerror("التاريخ", "راجع التاريخ (يوم/شهر/سنة).")
            return
        now_t = datetime.now().strftime("%H:%M")
        if iso != date.today().isoformat():
            if not messagebox.askyesno("تاريخ مختلف",
                                       f"التاريخ المختار ({fmt_iso_date(iso)}) مش تاريخ النهاردة.\n"
                                       f"هل تسجّل الوقت الحالي ({fmt_time(now_t)}) على هذا التاريخ؟"):
                return
        notes = self.c_notes.get().strip()
        with db() as con:
            r = con.execute("SELECT id, in_time, out_time FROM attendance WHERE employee=? AND att_date=?",
                            (name, iso)).fetchone()

        if kind == "in":
            if r and r[1]:
                messagebox.showwarning("مسجل بالفعل",
                                       f"«{name}» مسجل حضور الساعة {fmt_time(r[1])} بتاريخ {fmt_iso_date(iso)}.\n"
                                       "لو عايز تصحّح الوقت، عدّل السجل من الجدول أو استخدم الحفظ اليدوي.")
                return
            if r:
                with db() as con:
                    if notes:
                        con.execute("UPDATE attendance SET in_time=?, notes=? WHERE id=?", (now_t, notes, r[0]))
                    else:
                        con.execute("UPDATE attendance SET in_time=? WHERE id=?", (now_t, r[0]))
            else:
                ok, msg = add_attendance(name, iso, now_t, "", notes)
                if not ok:
                    messagebox.showwarning("تنبيه", msg)
                    return
            self.c_in.set(fmt_time(now_t))
            title = "تم تسجيل الحضور"
            body = f"✅ حضور «{name}» الساعة {fmt_time(now_t)}\n{fmt_iso_date(iso)}"
        else:
            if not r or not r[1]:
                messagebox.showwarning("مفيش حضور مسجل",
                                       f"«{name}» مالوش حضور مسجل بتاريخ {fmt_iso_date(iso)}.\n"
                                       "سجّل الحضور الأول (زر ✅ حضور الآن) وبعدها اضغط انصراف.")
                return
            if r[2]:
                messagebox.showwarning("مسجل بالفعل",
                                       f"«{name}» مسجل انصراف الساعة {fmt_time(r[2])} بتاريخ {fmt_iso_date(iso)}.")
                return
            if _mins(now_t) <= _mins(r[1]):
                messagebox.showwarning("الوقت",
                                       f"وقت الانصراف الحالي ({fmt_time(now_t)}) لازم يكون بعد وقت الحضور ({fmt_time(r[1])}).\n"
                                       "لو الدوام ليلاً، استخدم الحفظ اليدوي للانصراف.")
                return
            with db() as con:
                if notes:
                    con.execute("UPDATE attendance SET out_time=?, notes=? WHERE id=?", (now_t, notes, r[0]))
                else:
                    con.execute("UPDATE attendance SET out_time=? WHERE id=?", (now_t, r[0]))
            self.c_out.set(fmt_time(now_t))
            title = "تم تسجيل الانصراف"
            body = f"🚪 انصراف «{name}» الساعة {fmt_time(now_t)}\n{fmt_iso_date(iso)}"

        self.c_notes.set("")
        self.load_clock()
        if hasattr(self, "_clock_refresh"):
            self._clock_refresh()
        messagebox.showinfo(title, body)

    def load_clock(self):
        with db() as con:
            rows = con.execute("""SELECT t.id, t.employee, COALESCE(e.department,'-'), t.att_date, t.in_time, t.out_time, t.notes
                FROM attendance t LEFT JOIN employees e ON e.name=t.employee
                ORDER BY t.att_date DESC, t.id DESC LIMIT 300""").fetchall()
        self.c_tree.delete(*self.c_tree.get_children())
        for n, (rid, name, dept, d, i, o, nt) in enumerate(rows):
            w, l, e2 = attendance_calc(dept, i, o)
            tags = ["odd"] if n % 2 else []
            if l or e2: tags.append("warn")
            self.c_tree.insert("", "end", iid=str(rid), tags=tags,
                               values=(name, DEPT_LABEL.get(dept, dept), fmt_iso_date(d), fmt_time(i) or "-",
                                       fmt_time(o) or "-", fmt_hours(w), l or "-", e2 or "-", nt or "-"))

    def delete_clock(self):
        sel = self.c_tree.selection()
        if not sel:
            messagebox.showwarning("اختر سجلًا", "حدد السجل المطلوب حذفه من الجدول."); return
        if not messagebox.askyesno("تأكيد", "حذف سجل الحضور والانصراف المحدد؟"):
            return
        with db() as con:
            con.execute("DELETE FROM attendance WHERE id=?", (int(sel[0]),))
        self.load_clock()

    def build_permits_tab(self, tab):
        names = self.worker_names()
        form = self.panel(tab); form.pack(fill="x", pady=4)
        self.p_emp = tk.StringVar(); self.p_date = tk.StringVar(value=today_text())
        self.p_out = tk.StringVar(); self.p_back = tk.StringVar(); self.p_reason = tk.StringVar()
        self.add_combo(form, "اسم العامل", self.p_emp, names, 0, 2)
        self.add_date(form, "تاريخ الإذن", self.p_date, 0, 1)
        self.add_combo(form, "وقت الخروج", self.p_out, time_options(7, 20, 15), 0, 0)
        self.add_field(form, "السبب", self.p_reason, 1, 2, 260)
        quota = tk.Label(form, text="", bg=WHITE, fg=BLUE, font=(FONT, 10, "bold"), justify="right", wraplength=300)
        quota.grid(row=1, column=1, padx=14, sticky="e")
        self.add_combo(form, "وقت العودة (اختياري)", self.p_back, [""] + time_options(7, 20, 15), 1, 0)
        bx = tk.Frame(form, bg=WHITE); bx.grid(row=2, column=0, columnspan=3, sticky="e", padx=14, pady=(0, 10))
        self.btn(bx, "💾 حفظ الإذن", self.save_permit, GREEN).pack(side="right", padx=4)
        self.btn(bx, "حذف المحدد", self.delete_permit, RED).pack(side="right", padx=4)

        def upd(*_):
            name = self.p_emp.get()
            if not names:
                quota.config(text="سجّل العاملين أولًا من «قاعدة بيانات العاملين»", fg=RED); return
            try:
                iso = parse_date(self.p_date.get()).isoformat()
            except ValueError:
                return
            if not name:
                quota.config(text=f"اختر العامل لعرض المتبقي من إذونات الشهر (الحد {PERMIT_LIMIT})", fg=MUTED); return
            with db() as con:
                n = permits_in_month(con, name, iso)
            left = max(0, PERMIT_LIMIT - n)
            quota.config(text=f"إذونات {name} في {MONTHS_AR[int(iso[5:7]) - 1]} {iso[:4]}: استخدم {n} من {PERMIT_LIMIT} - المتبقي {left}"
                         + ("\n⛔ لا يسمح بإذن ثالث" if left == 0 else ""), fg=RED if left == 0 else GREEN)
        self.p_emp.trace_add("write", upd); self.p_date.trace_add("write", upd); upd()
        self._p_upd = upd
        self.p_tree = self.make_tree(tab, ("name", "dept", "date", "out", "back", "reason", "no"),
                                     ("اسم العامل", "القسم", "التاريخ", "الخروج", "العودة", "السبب", "رقم الإذن بالشهر"),
                                     (160, 70, 100, 90, 90, 220, 120), 8)
        self.load_permits()

    def save_permit(self):
        name = self.p_emp.get().strip()
        if not name:
            messagebox.showwarning("بيانات ناقصة", "اختر العامل من القائمة أولًا."); return
        try:
            iso = parse_date(self.p_date.get()).isoformat()
        except ValueError:
            messagebox.showerror("التاريخ", "راجع تاريخ الإذن."); return
        out_t, back_t = parse_time_label(self.p_out.get()), parse_time_label(self.p_back.get())
        if not out_t:
            messagebox.showwarning("بيانات ناقصة", "اختر وقت الخروج."); return
        if back_t and back_t <= out_t:
            messagebox.showerror("الوقت", "وقت العودة لازم يكون بعد وقت الخروج."); return
        ok, res = add_permit(name, iso, out_t, back_t, self.p_reason.get().strip())
        if not ok:
            messagebox.showerror("غير مسموح - لا يسمح بإذن ثالث", res); return
        self.p_reason.set(""); self.p_back.set(""); self.p_out.set("")
        self.load_permits(); self._p_upd()
        messagebox.showinfo("تم الحفظ", f"تم تسجيل الإذن رقم {res} من {PERMIT_LIMIT} لهذا الشهر.")

    def load_permits(self):
        with db() as con:
            rows = con.execute("""SELECT p.id, p.employee, COALESCE(e.department,'-'), p.permit_date, p.out_time, p.back_time,
                    p.reason, (SELECT COUNT(*) FROM permits q WHERE q.employee=p.employee
                               AND substr(q.permit_date,1,7)=substr(p.permit_date,1,7) AND q.id<=p.id)
                FROM permits p LEFT JOIN employees e ON e.name=p.employee
                ORDER BY p.permit_date DESC, p.id DESC LIMIT 200""").fetchall()
        self.p_tree.delete(*self.p_tree.get_children())
        for n, (rid, name, dept, d, o, b, why, k) in enumerate(rows):
            self.p_tree.insert("", "end", iid=str(rid), tags=("odd",) if n % 2 else (),
                               values=(name, dept, fmt_iso_date(d), fmt_time(o), fmt_time(b) or "-", why or "-", f"{k} من {PERMIT_LIMIT}"))

    def delete_permit(self):
        sel = self.p_tree.selection()
        if not sel:
            messagebox.showwarning("اختر سجلًا", "حدد الإذن المطلوب حذفه من الجدول."); return
        if not messagebox.askyesno("تأكيد", "حذف الإذن؟ (سيعود للعامل رصيد إذن في هذا الشهر)"):
            return
        with db() as con:
            con.execute("DELETE FROM permits WHERE id=?", (int(sel[0]),))
        self.load_permits(); self._p_upd()

    def build_absence_tab(self, tab):
        form = self.panel(tab); form.pack(fill="x", pady=4)
        self.a_emp = tk.StringVar(); self.a_date = tk.StringVar(value=today_text())
        self.a_kind = tk.StringVar(value=ABSENCE_KINDS[0]); self.a_arr = tk.StringVar(); self.a_notes = tk.StringVar()
        self.add_combo(form, "اسم العامل", self.a_emp, self.worker_names(), 0, 2)
        self.add_date(form, "التاريخ", self.a_date, 0, 1)
        self.add_combo(form, "النوع", self.a_kind, ABSENCE_KINDS, 0, 0)
        arr = self.add_combo(form, "وقت الحضور (للتأخير فقط)", self.a_arr, time_options(6, 12, 5), 1, 2)
        info = tk.Label(form, text="", bg=WHITE, fg=BLUE, font=(FONT, 10, "bold"), justify="right", wraplength=300)
        info.grid(row=1, column=1, padx=14, sticky="e")
        self.add_field(form, "ملاحظات", self.a_notes, 1, 0, 260)
        bx = tk.Frame(form, bg=WHITE); bx.grid(row=2, column=0, columnspan=3, sticky="e", padx=14, pady=(0, 10))
        self.btn(bx, "💾 حفظ", self.save_absence, GREEN).pack(side="right", padx=4)
        self.btn(bx, "حذف المحدد", self.delete_absence, RED).pack(side="right", padx=4)

        def upd(*_):
            is_late = self.a_kind.get() == "تأخير"
            arr.configure(state="readonly" if is_late else "disabled")
            if not is_late and self.a_arr.get():
                self.a_arr.set("")
            name = self.a_emp.get()
            dept = worker_dept(name) if name else ""
            txt = f"قسم {dept}: بداية الدوام {fmt_time(DEPT_START.get(dept, '07:00'))}" if dept else ""
            if is_late and self.a_arr.get():
                lm = late_minutes(dept, parse_time_label(self.a_arr.get()))
                txt += f"\nمدة التأخير: {lm} دقيقة" if lm else "\nحضر في الميعاد (لا يوجد تأخير)"
            info.config(text=txt)
        for v in (self.a_emp, self.a_kind, self.a_arr):
            v.trace_add("write", upd)
        upd()
        self.a_tree = self.make_tree(tab, ("name", "dept", "date", "kind", "arr", "late", "notes"),
                                     ("اسم العامل", "القسم", "التاريخ", "النوع", "وقت الحضور", "دقائق التأخير", "ملاحظات"),
                                     (160, 70, 100, 120, 100, 100, 200), 8)
        self.a_tree.tag_configure("warn", foreground=ORANGE)
        self.load_absences()

    def save_absence(self):
        name = self.a_emp.get().strip()
        if not name:
            messagebox.showwarning("بيانات ناقصة", "اختر العامل من القائمة أولًا."); return
        try:
            iso = parse_date(self.a_date.get()).isoformat()
        except ValueError:
            messagebox.showerror("التاريخ", "راجع التاريخ."); return
        ok, res = add_absence(name, iso, self.a_kind.get(), parse_time_label(self.a_arr.get()), self.a_notes.get().strip())
        if not ok:
            messagebox.showwarning("تنبيه", res); return
        self.a_notes.set(""); self.a_arr.set("")
        self.load_absences()
        messagebox.showinfo("تم الحفظ", "تم التسجيل." + (f" مدة التأخير {res} دقيقة." if self.a_kind.get() == "تأخير" else ""))

    def load_absences(self):
        with db() as con:
            rows = con.execute("""SELECT a.id, a.employee, COALESCE(e.department,'-'), a.absence_date, a.kind,
                    a.arrival_time, a.late_minutes, a.notes
                FROM absences a LEFT JOIN employees e ON e.name=a.employee
                ORDER BY a.absence_date DESC, a.id DESC LIMIT 200""").fetchall()
        self.a_tree.delete(*self.a_tree.get_children())
        for n, (rid, name, dept, d, kind, arr, lm, notes) in enumerate(rows):
            tags = ["odd"] if n % 2 else []
            if kind == "غياب بدون إذن": tags.append("bad")
            elif kind == "تأخير": tags.append("warn")
            self.a_tree.insert("", "end", iid=str(rid), tags=tags,
                               values=(name, dept, fmt_iso_date(d), kind, fmt_time(arr) or "-", lm or "-", notes or "-"))

    def delete_absence(self):
        sel = self.a_tree.selection()
        if not sel:
            messagebox.showwarning("اختر سجلًا", "حدد السجل المطلوب حذفه من الجدول."); return
        if not messagebox.askyesno("تأكيد", "حذف السجل المحدد؟"):
            return
        with db() as con:
            con.execute("DELETE FROM absences WHERE id=?", (int(sel[0]),))
        self.load_absences()

    # ---------- مقارنة الفترات ----------
    def show_compare(self):
        self.page_title("📈 مقارنة الفترات", "قارن أي فترتين: الزيارات والصيانة والأخطاء والإشادات وأداء كل عامل - مع ملخص ذكي ورسم بياني")
        panel = self.panel(); panel.pack(fill="x", padx=24, pady=4)
        self.cmp_vars = [tk.StringVar(value=fmt_iso_date(x)) for x in preset_periods("month")]
        tk.Label(panel, text="الفترة الحالية (أ)", bg=WHITE, fg=BLUE, font=(FONT, 11, "bold")).grid(row=0, column=2, padx=14, sticky="e")
        self.add_date(panel, "من", self.cmp_vars[0], 0, 1)
        self.add_date(panel, "إلى", self.cmp_vars[1], 0, 0)
        tk.Label(panel, text="الفترة السابقة (ب)", bg=WHITE, fg=ORANGE, font=(FONT, 11, "bold")).grid(row=1, column=2, padx=14, sticky="e")
        self.add_date(panel, "من", self.cmp_vars[2], 1, 1)
        self.add_date(panel, "إلى", self.cmp_vars[3], 1, 0)
        bar = tk.Frame(panel, bg=WHITE); bar.grid(row=2, column=0, columnspan=3, sticky="e", padx=14, pady=(0, 10))
        self.btn(bar, "⚖️ قارن الآن", self.run_compare, GREEN).pack(side="right", padx=3)
        for label, k in (("الشهر مقابل الماضي", "month"), ("الأسبوع مقابل الماضي", "week"), ("السنة مقابل الماضية", "year")):
            self.btn(bar, label, lambda k=k: self.set_preset(k), BLUE).pack(side="right", padx=3)
        self.btn(bar, "🖨️ طباعة", self.print_compare, PURPLE).pack(side="right", padx=3)
        self.btn(bar, "📥 Excel", self.export_compare, ORANGE).pack(side="right", padx=3)
        self.cmp_area = tk.Frame(self.content, bg=BG); self.cmp_area.pack(fill="both", expand=True)
        self.cmp_data = None
        self.run_compare()

    def set_preset(self, kind):
        for var, val in zip(self.cmp_vars, preset_periods(kind)):
            var.set(fmt_iso_date(val))
        self.run_compare()

    def read_compare(self):
        try:
            d = [parse_date(v.get()).isoformat() for v in self.cmp_vars]
            if d[1] < d[0] or d[3] < d[2]: raise ValueError
            return d
        except ValueError:
            messagebox.showerror("التاريخ", "راجع تواريخ الفترتين (يوم/شهر/سنة) وتأكد أن تاريخ النهاية بعد البداية.")
            return None

    def run_compare(self):
        p = self.read_compare()
        if not p: return
        self.cmp_data = data = compare_periods(*p)
        for w in self.cmp_area.winfo_children():
            w.destroy()
        self.insight_box(self.cmp_area, "🧠 الملخص الذكي", data["insights"]).pack(fill="x", padx=31, pady=(4, 2))
        nb = ttk.Notebook(self.cmp_area); nb.pack(fill="both", expand=True, padx=24, pady=(6, 8))
        t1 = tk.Frame(nb, bg=BG); t2 = tk.Frame(nb, bg=BG)
        nb.add(t1, text="  الملخص والرسم البياني  "); nb.add(t2, text="  مقارنة العاملين  ")

        holder = tk.Frame(t1, bg=BG); holder.pack(fill="both", expand=True)
        tf = tk.Frame(holder, bg=BG); tf.pack(side="right", fill="y")
        tree = self.make_tree(tf, ("item", "a", "b", "d", "pct"), ("البند", "الحالية", "السابقة", "الفرق", "النسبة"),
                              (140, 70, 70, 80, 70), 8)
        for n, (name, a, b, d, pct) in enumerate(data["rows"]):
            pol = POLARITY.get(name, 0)
            tags = ["odd"] if n % 2 else []
            if d and pol: tags.append("good" if pol * d > 0 else "bad")
            t1_, t2_ = _sign_txt(d, pct)
            tree.insert("", "end", values=(name, a, b, t1_, t2_), tags=tags)
        cv = tk.Canvas(holder, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        cv.pack(side="left", fill="both", expand=True, padx=(24, 0), pady=(6, 10))
        labels = [r[0] for r in data["rows"]]
        cv.bind("<Configure>", lambda e: draw_bars(cv, labels, [r[1] for r in data["rows"]], [r[2] for r in data["rows"]]))

        et = self.make_tree(t2, ("name", "dept", "ea", "eb", "d", "pa", "na"),
                            ("اسم العامل", "القسم", "أخطاء حالية", "أخطاء سابقة", "الفرق", "إشادات حالية", "ملاحظات حالية"),
                            (190, 80, 100, 100, 80, 100, 110), 9)
        for n, e in enumerate(data["emps"]):
            d = e["err_a"] - e["err_b"]
            tags = ["odd"] if n % 2 else []
            if d > 0: tags.append("bad")
            elif d < 0: tags.append("good")
            et.insert("", "end", values=(e["name"], e["dept"], e["err_a"], e["err_b"], _sign_txt(d, None)[0],
                                         e["praise_a"], e["note_a"]), tags=tags)

    def print_compare(self):
        if not self.cmp_data: self.run_compare()
        if not self.cmp_data: return
        try:
            open_html(build_compare_html(self.cmp_data), "compare")
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تجهيز الطباعة:\n{ex}")

    def export_compare(self):
        if not self.cmp_data: self.run_compare()
        if not self.cmp_data: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"مقارنة_{APP_NAME}.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path: return
        try:
            build_compare_excel(path, self.cmp_data)
        except PermissionError:
            messagebox.showerror("الملف مفتوح", "الملف مفتوح في Excel. أغلقه أو اختر اسمًا آخر."); return
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر إنشاء الملف:\n{ex}"); return
        if messagebox.askyesno("تم", f"تم حفظ المقارنة:\n{path}\n\nهل تريد فتح الملف الآن؟"):
            try: os.startfile(path)
            except Exception: pass

    # ---------- التقارير: طباعة ----------
    def print_report(self):
        p = self.read_period()
        if not p: return
        try:
            open_html(build_report_html(self.report_type.get(), *p), "report")
        except Exception as ex:
            messagebox.showerror("خطأ", f"تعذر تجهيز الطباعة:\n{ex}")

    def q_today(self):
        self.rep_from.set(today_text()); self.rep_to.set(today_text())

    def q_month(self):
        self.rep_from.set(date.today().replace(day=1).strftime("%d/%m/%Y")); self.rep_to.set(today_text())

    def q_year(self):
        self.rep_from.set(date.today().replace(month=1, day=1).strftime("%d/%m/%Y")); self.rep_to.set(today_text())

    def read_period(self):
        try:
            s = parse_date(self.rep_from.get()).isoformat(); e = parse_date(self.rep_to.get()).isoformat()
            if e < s: raise ValueError
            return s, e
        except ValueError:
            messagebox.showerror("التاريخ", "راجع تاريخ البداية والنهاية (يوم/شهر/سنة).")
            return None

    def preview_counts(self):
        p = self.read_period()
        if not p: return
        c = period_counts(*p)
        self.counts_lbl.config(text="عدد السجلات في الفترة:\n" + "\n".join(f"{k}: {v}" for k, v in c.items()))

    def export_report(self):
        p = self.read_period()
        if not p: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"تقرير_{APP_NAME}.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path: return
        self.root.config(cursor="watch"); self.root.update_idletasks()
        try:
            res = build_report(path, self.report_type.get(), *p)
        except PermissionError:
            self.root.config(cursor="")
            messagebox.showerror("الملف مفتوح", "الملف مفتوح في Excel. أغلقه أو اختر اسمًا آخر ثم أعد المحاولة."); return
        except Exception as ex:
            self.root.config(cursor="")
            messagebox.showerror("خطأ", f"تعذر إنشاء التقرير:\n{ex}"); return
        self.root.config(cursor="")
        total = sum(v for k, v in res.items()
                    if k not in SUMMARY_ONLY + ("قاعدة العاملين", "قاعدة بيانات موظفي الجهات"))
        if self.report_type.get() == "workers":
            total = res.get("قاعدة العاملين", 0)
        if self.report_type.get() == "govstaff":
            total = res.get("قاعدة بيانات موظفي الجهات", 0)
        msg = f"تم إنشاء التقرير بنجاح ({total} سجل):\n{path}"
        if total == 0:
            msg += "\n\nملاحظة: لا توجد سجلات في هذه الفترة."
        if messagebox.askyesno("تم استخراج التقرير", msg + "\n\nهل تريد فتح الملف الآن؟"):
            try: os.startfile(path)
            except Exception: pass


# ---------- حراسة العمليات التي تكتب في البيانات (إضافة = add ، تعديل/حذف = edit) ----------
_TABLE_PAGE = {"visits": "visits", "maintenance": "maintenance", "employee_notes": "employees",
               "gov_staff": "govstaff", "gov_attendance": "gov"}


def _guard(level, page):
    def deco(fn):
        @functools.wraps(fn)
        def wrapper(self, *a, **k):
            pg = page(self, *a, **k) if callable(page) else page
            if not self.need(pg, level):
                return None
            return fn(self, *a, **k)
        return wrapper
    return deco


for _name, _lv, _pg in (
        ("save_visit", "add", "visits"), ("save_maintenance", "add", "maintenance"),
        ("save_employee_note", "add", "employees"),
        ("edit_record", "edit", lambda self, table, tree: _TABLE_PAGE[table]),
        ("delete_selected", "edit", lambda self, table, tree: _TABLE_PAGE[table]),
        ("delete_worker", "edit", "workers"),
        ("save_inventory", "add", lambda self, category: category),
        ("stock_in_dialog", "add", lambda self, category: category),
        ("edit_cell", "add", lambda self, *a: self.active_key),
        ("purchase_dialog", "add", "purchases"), ("attach_invoice", "edit", "purchases"),
        ("delete_purchase", "edit", "purchases"), ("edit_paid_by", "edit", "purchases"),
        ("save_clock", "add", "attend"), ("clock_now", "add", "attend"), ("delete_clock", "edit", "attend"),
        ("save_permit", "add", "attend"), ("delete_permit", "edit", "attend"),
        ("save_absence", "add", "attend"), ("delete_absence", "edit", "attend"),
        ("import_att_data", "add", "attend"), ("import_gov_att", "add", "gov"),
        ("gov_picker", "add", "gov"), ("gov_action_dialog", "add", "gov"), ("save_govstaff", "add", "govstaff"),
        ("save_note", "add", "notes"), ("delete_note", "edit", "notes")):
    setattr(CenterApp, _name, _guard(_lv, _pg)(getattr(CenterApp, _name)))


def start_net_server():
    """يشغّل خادم الشبكة في وضع «جهاز رئيسي (سيرفر)». يرجع كائن الخادم أو None."""
    global NET_SERVER
    if not netmod or NET_MODE != netmod.MODE_SERVER:
        return None
    if NET_SERVER and NET_SERVER.httpd:
        return NET_SERVER

    def _log(msg):
        try:
            with open(os.path.join(BASE, "net.log"), "a", encoding="utf-8") as f:
                f.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "  " + str(msg) + "\n")
        except Exception:
            pass

    NET_SERVER = netmod.Server(
        DB_PATH,
        token=NET_CFG.get("token", ""),
        verify=lambda u, p: authenticate_local(u, p)[0],
        photos_dir=PHOTOS_DIR,
        backup_fn=backup_now,
        backup_full_fn=backup_full_now,
        log=_log,
    )
    try:
        NET_SERVER.start(NET_CFG.get("bind_host", "0.0.0.0"), int(NET_CFG.get("port", 8765)))
    except Exception as ex:
        _log("تعذر تشغيل السيرفر: %s" % ex)
        try:
            messagebox.showerror(
                "الشبكة", f"تعذر تشغيل خادم الشبكة على المنفذ {NET_CFG.get('port')}:\n{ex}\n\n"
                          "افتح نافذة «الشبكة (LAN)» من قائمة المستخدم وغيّر المنفذ.")
        except Exception:
            pass
        return None
    return NET_SERVER


if __name__ == "__main__":
    apply_cli()
    try:
        auto_backup()
    except Exception:
        pass
    root = tk.Tk()
    root.withdraw()

    def _on_error(exc_type, exc, tb):
        """تسجيل أي خطأ في ملف + رسالة للمستخدم (مهم في نسخة exe بدون كونسول)."""
        try:
            with open(CRASH_LOG, "a", encoding="utf-8") as f:
                f.write("\n" + "=" * 70 + "\n" +
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
                traceback.print_exception(exc_type, exc, tb, file=f)
        except Exception:
            pass
        try:
            messagebox.showerror("خطأ غير متوقع",
                                 f"حصل خطأ غير متوقع:\n{exc}\n\n"
                                 f"تم تسجيل التفاصيل في الملف:\n{CRASH_LOG}")
        except Exception:
            pass

    root.report_callback_exception = _on_error
    sys.excepthook = _on_error
    try:
        if not ensure_admin_password(root):
            root.destroy()
            sys.exit(0)
        start_net_server()
        _user = None
        while _user is None:
            _user = login_dialog(root)
            if _user == "RETRY":
                _user = None
        if not _user:
            root.destroy()
            sys.exit(0)
        root.deiconify()
        CenterApp(root, _user)
        root.mainloop()
    except SystemExit:
        raise
    except Exception as _ex:      # فشل مبكر (قبل بناء الواجهة) - يُسجَّل بدل الاختفاء الصامت
        try:
            with open(CRASH_LOG, "a", encoding="utf-8") as f:
                f.write("\n" + "=" * 70 + "\n" +
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
                traceback.print_exc(file=f)
        except Exception:
            pass
        try:
            messagebox.showerror("تعذر تشغيل البرنامج",
                                 f"{_ex}\n\nالتفاصيل في الملف:\n{CRASH_LOG}")
        except Exception:
            pass
        raise


# ============================================================================
#  ★ نقطة نهاية آخر تحديث (END OF LAST UPDATE) — الإصدار 2.11.0 ★
# ----------------------------------------------------------------------------
#  كل الكود فوق السطر ده خلص ومتّفق عليه. أي إضافة أو تعديل جديد يبدأ من هنا.
#
#  آخر حاجة اتعملت (2.11.0): ثيم «ذهبي فاخر» (Luxury Gold) من مواصفات تصميم المستخدم —
#    • THEMES["gold"]  : ثيم ليلي فاخر بالألوان المطلوبة بالظبط:
#        BG #1B2028 (خلفية اللوحة)   NAVY/NAVY_2 #14181F (الشريط العلوي + القائمة)
#        WHITE #323844 (تعبئة البطاقات) SOFT #252A34 (نهاية التدرّج) ENTRY #14181F (حقول الإدخال)
#        GOLD #E0B963 (الذهبي الأساسي) PRIMARY/BLUE #C89D46 (الذهبي الثانوي)
#        TEXT #FFFFFF  MUTED #A0A8B5  LINE #3A4150
#        H1 #0E3B33 → H2 #151B2E (تدرّج الشريط = تدرّج شاشة الدخول)
#        + مفاتيح خاصة: BTN="gold" / FONT_KEY="cairo" / LOGIN_G1,LOGIN_G2,LOGIN_BOX,LOGIN_BTN.
#    • BTN_STYLES["gold"] : شكل أزرار جديد «ذهبي فاخر» — تعبئة داكنة + إطار ذهبي 1px +
#      نص ذهبي، وعند المرور يظهر لون الزر الأصلي (أخضر للحفظ، أحمر للحذف… إلخ).
#    • apply_appearance() : أي ثيم بيعرّف BTN/FONT_KEY بيطّبق هويته كاملة (أزرار + خط)
#      لو المستخدم مش حدّدهم بنفسه → «الثيم بيغيّر لون الواجهة والأزرار».
#    • gradient_button()  : دالة جديدة ترسم زرار بتدرّج لوني حقيقي على Canvas بإطار ذهبي
#      ونص ذهبي (تدرّج أخضر ← ذهبي ← كحلي زي مواصفات زر الدخول).
#    • login_dialog()     : في الثيمات اللي بتدعم الهوية الفاخرة — خلفية متدرّجة
#      (Canvas بيتعاد رسمه مع أي تغيير حجم) + صندوق داكن «زجاجي» بإطار ذهبي + زرار دخول
#      متدرّج. الثيمات التانية زي ما هي بالظبط.
#    • الخطوط المرفقة     : مجلد fonts/ فيه Cairo (Regular+Bold) و Tajawal (Regular+Bold)
#      برخصة OFL، + register_bundled_fonts() يسجّلها تلقائيًا عند التشغيل
#      (ويندوز: AddFontResourceExW بصلاحية خاصة للعملية / لينكس: ~/.local/share/fonts + fc-cache)،
#      فمفيش حاجة المستخدم يعملها. وUI_FONTS بقى فيه cairo و tajawal.
#    • RES_DIR            : مسار الموارد (img / fonts) بيدعم الـ exe (onefile → _MEIPASS).
#    • KhadamatMisr.spec  : datas بقى فيه ('fonts', 'fonts') عشان الخطوط تدخل الـ exe.
#    • شاشة المظهر        : عنوان الثيمات بيحسب العدد تلقائيًا (بقى 9) + سطر توضيحي
#      إن ثيم «ذهبي فاخر» بيغيّر الأزرار والخط معاه.
#
#  قبل كده (2.10.0): شاشة المظهر بمسطرتين + ملء الشاشة المتكيّف + توحيد الاسم —
#    • appearance_dialog() : محتوى شاشة «المظهر والثيمات» بقى جوه Canvas بمسطرتين
#      (رأسية + أفقية) عشان توصل لآخر الثيمات على أي مقاس شاشة، + عجلة الفأرة جوه
#      الشاشة، + سطر الاختصارات بيتلوّن تلقائيًا (wraplength متكيّف)، + حد أدنى
#      للنافذة 420×320 بدل 840×660.
#    • toggle_fullscreen() : ملء الشاشة بيعيد ضبط الواجهة فورًا (وعلى ويندوز
#      attributes("-fullscreen") مع fallback لمقاس الشاشة) — فمفيش حاجة مختفية
#      على أي رزوليوشن.
#    • _relayout()         : دالة جديدة تعيد ترتيب القائمة الجانبية + المحتوى +
#      تمرير القائمة بعد أي تغيير حجم/ملء شاشة.
#    • توحيد الاسم «قاعدة بيانات موظفي الجهات» (نفس صيغة القائمة الجانبية) في:
#      لستة أنواع التقارير + ملخص الأعداد + ورقة Excel + نص المساعدة +
#      سطر الاختصارات في شاشة المظهر + ترجمة i18n.
#
#  قبل كده (2.9.0): مسطرة تمرير + تنقّل بالأسهم في القائمة الجانبية —
#    • nav_wrap/nav_canvas/nav_vsb/nav_inner/nav_inner_id : أزرار القائمة الجانبية
#      جوه Canvas بمسطرة تمرير رأسية، فبتوصل لآخر القائمة على أي مقاس شاشة.
#    • _fit_nav()      : يضبط scrollregion + يخفي المسطرة تلقائيًا لو القائمة كلها بانية.
#    • nav_move()      : ينقل التحديد في القائمة الجانبية (لفوق/لأسفل) بحدود آمنة.
#    • _nav_reveal()   : ينزّل/يطلّع القائمة عشان العنصر المختار يبان بالكامل.
#    • _menu_arrow()   : ↑/↓ تتحرّك في القائمة — وبتسيب المفاتيح للجداول/الخانات
#      لما يكون التركيز عليها (Entry/Treeview/Listbox/…) عشان ما تعطّلهاش.
#    • bind_keys()     : <Up>/<Down> على النافذة، وكل زرار قائمة عنده ↑↓→←.
#    • _on_wheel()     : عجلة الفأرة بتعمل تمرير للقائمة الجانبية كمان.
#    • shortcuts_dialog() : سطر جديد «↑ / ↓ — التحرّك في القائمة الجانبية».
#
#  قبل كده (2.8.0): الواجهة تبان كاملة على أي شاشة + الكيبورد —
#    • fit_to_screen()      : بيقيس الشاشة الحقيقية ويفتح بمقاس مضمون إنه يدخل فيها.
#    • content_canvas/content_vsb/content_hsb : منطقة المحتوى جوه Canvas بشريطي تمرير.
#    • _fit_content()/_show_scrollbars() : يخلي المحتوى يملأ المتاح + يخفي الشريط لو مش محتاجه.
#    • layout() (جوه build_shell) : هيدر متكيّف — يصغّر العنوان ويزحزحه، ويختصر نصوص الأزرار
#      على الشاشات الضيقة جدًا، فمفيش تداخل بين العنوان والأزرار.
#    • _side_width()/_on_root_config() : عرض القائمة الجانبية بيتغيّر حسب عرض النافذة.
#    • _on_wheel()/_bind_wheel() : تمرير المحتوى بعجلة الفأرة (ويندوز ولينكس).
#    • toggle_fullscreen()  : ملء الشاشة (قائمة المستخدم + Ctrl+F).
#    • bind_keys()          : **إصلاح** F1…F11 كانت bind("F1") بدون أقواس → bind("<F1>")
#      (كانت مش شغّالة أصلاً)، + PageUp/Down و Ctrl+Home/End للتمرير.
#    • shortcuts_dialog()   : نافذة «⌨️ اختصارات لوحة المفاتيح» (Shift+F1).
#    • _clamp_size()/_minsize() : كل النوافذ المنبثقة بتتصغّر لو أكبر من الشاشة.
#
#  قبل كده (2.7.0): محرّك المظهر والثيمات —
#    • THEMES/THEME_LABELS/THEME_ORDER/DARK_THEMES : 8 ثيمات (ألوان مستخرجة من صور المستخدم).
#    • UI_FONTS/FONT_LABELS + apply_font()         : 6 خطوط (يغيّر FONT العام).
#    • FONT_SIZES/FONT_SIZE_LABELS + apply_font_size() : صغير/عادي/كبير عبر tk scaling (_apply_scaling).
#    • BTN_STYLES + apply_btn_style()              : مملوء/محدّد/ناعم — btn() بيحترم BTN_STYLE.
#    • apply_theme()   : يحدّث الألوان العامة + HDR_TEXT/HDR_SUB/SIDE_TEXT/SIDE_MUTED/SIDE_ACCENT تلقائيًا.
#    • appearance_dialog()/apply_appearance() : شاشة «🎨 المظهر والثيمات» (تطبيق فوري + حفظ في settings.json).
#    • toggle_theme()  : 🌙/☀️ بيفضل فاكر آخر ثيم نهاري (settings["last_light"]).
#    • bind_keys()/cycle_page()/HOTKEYS : F1–F11 أقسام، F12 مظهر، Alt+→/← تنقّل، Esc إغلاق.
#
#  قبل كده (2.6.2): شيل «🧠 ملخص ذكي» من الصفحة الرئيسية —
#    • show_home() : اتشال منها صندوق الملخص الذكي اللي كان بياخد مساحة فوق الكروت.
#      (الملخص الذكي نفسه لسه موجود جوه صفحة «📈 مقارنة الفترات» — insight_box.)
#
#  قبل كده (2.6.1): تنبيه بصوت عند وصول ملاحظة جديدة —
#    • جدول notes      : body/target/priority/created_by/created_at/status/done_by/done_at.
#    • show_notes()    : شاشة «📢 ملاحظات» (إرسال + فلتر الكل/جديد/تم + تقديم + حذف + مفتاح الصوت).
#    • build_notes_banner() : شريط أصفر أعلى أي صفحة على الوحدة المستهدفة + زر «✅ قدمها».
#    • mark_note_done()     : تحويل الملاحظة لـ «تم» (محليًا أو عبر /note_done على الشبكة).
#    • _play_note_sound()   : تنبيه صوتي عند وصول ملاحظة جديدة (العاجلة بجرسين) — يُحفظ في settings.json.
#    • net.py               : notes في خريطة الصلاحيات + نقطة /note_done + دالة العميل mark_note_done().
#
#  تعديل لاحق (صلاحيات): user_edit_dialog — لوحة الصلاحيات بقت جوه Canvas بمسطرة جانبية
#    + تمرير بالعجلة عشان كل الأقسام تظهر حتى لو النافذة صغيرة.
#
#  قبل ما تبدأ أي تحديث، افتح ملف «التطوير.md» — فيه:
#    • آخر نقطة توقف (الإصدار + آخر commit + آخر بناء ناجح)
#    • خريطة كل دالة فين (بالسطر)
#    • قائمة المتبقي (Pending) والدوال المطلوب لمسها
#    • القواعد الثابتة (ممنوع تشكيل بيانات قاعدة البيانات، ممنوع رفع بيانات العملاء)
# ============================================================================
