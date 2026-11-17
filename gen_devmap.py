# -*- coding: utf-8 -*-
"""يولّد ملف «التطوير.md»: خريطة البرنامج + آخر نقطة توقف + خطة التحديث القادم.

شغّله بعد أي تعديل:  python3 tools/gen_devmap.py
"""
import ast
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "center_management.py")
OUT = os.path.join(ROOT, "التطوير.md")

APP_VERSION = "?"
for line in io.open(SRC, encoding="utf-8-sig"):
    if line.startswith("APP_VERSION"):
        APP_VERSION = line.split("=", 1)[1].strip().strip('"')
        break

# آخر نقطة توقف (حدّثها بعد كل بناء/رفع)
LAST_COMMIT = "8d2fa9f — الواجهة تبان كاملة على أي شاشة + إصلاح اختصارات الكيبورد (2.8.0)"
LAST_RUN = "44104691904"
EXE_BYTES = "16,372,059"
EXE_MD5 = "205bea10828b8f6cb663036f9564ce1f"
PKG_NAME = "KhadamatMisr_v2.8.0_windows.zip"
PKG_FILES = "559"
SHOTS = "18"


def structure(path):
    src = io.open(path, encoding="utf-8-sig").read()
    tree = ast.parse(src)
    out = []
    for node in tree.body:
        if isinstance(node, ast.ClassDef):
            methods = [
                (m.name, m.lineno)
                for m in node.body
                if isinstance(m, (ast.FunctionDef, ast.AsyncFunctionDef))
            ]
            out.append(("CLASS", node.name, node.lineno, methods))
        elif isinstance(node, ast.FunctionDef):
            out.append(("FUNC", node.name, node.lineno, []))
    return out, src.count("\n") + 1


def main():
    cm, cm_lines = structure(SRC)
    i18n_lines = sum(1 for _ in io.open(os.path.join(ROOT, "i18n.py"), encoding="utf-8-sig"))
    net_path = os.path.join(ROOT, "net.py")
    net_lines = sum(1 for _ in io.open(net_path, encoding="utf-8-sig")) if os.path.exists(net_path) else 0

    buf = []
    w = buf.append
    w("# ملف متابعة تطوير «خدمات مصر»")
    w("")
    w("> اقرأ الملف ده **الأول** قبل أي تعديل. فيه آخر نقطة وقفنا عندها، وخريطة كل حاجة فين،")
    w("> وخطة التحديث الجاي بالدوال المطلوب لمسها. مش محتاج تقرأ الـ 5000 سطر كلهم.")
    w("")
    w("## 1) آخر نقطة توقف (Checkpoint)")
    w("")
    w("| البند | القيمة |")
    w("|---|---|")
    w(f"| الإصدار الحالي | `{APP_VERSION}` |")
    w("| الفرع | `khadamat-misr` |")
    w("| المستودع | `hazemfeteha-lgtm/my-first-pro` |")
    w(f"| آخر commit | `{LAST_COMMIT}` |")
    w(f"| آخر بناء ناجح | GitHub Actions run `{LAST_RUN}` |")
    w(f"| ملف exe | `KhadamatMisr.exe` · {EXE_BYTES} بايت · MD5 `{EXE_MD5}` |")
    w(f"| الحزمة | `{PKG_NAME}` · {PKG_FILES} ملف |")
    w(f"| عدد اللقطات | {SHOTS} لقطة + `all_screens.png` + `gallery.html` |")
    w("")
    w("### خلص (Done)")
    w("")
    w("- 2.1.0: أمان (كلمة مرور، `_ident()`، crash log، تشكيل العربي، أيقونات onefile).")
    w("- 2.2.0: تقارير حضور الجهات، نوعا تقرير `govatt`/`govstaff`، 7 لوحات ألوان + وضع ليلي + 3 أشكال عرض.")
    w("- 2.3.0: حضور/انصراف بضغطة واحدة بساعة الجهاز (`clock_now`).")
    w("- 2.3.1: تشكيل العربي في خانات الجداول + قيم القوائم على لينكس، مع قراءة خام نظيفة.")
    w("- 2.4.0: **شبكة LAN** — جهاز رئيسي (سيرفر) + محطات، باسورد محدود لكل جهاز،")
    w("  فحص صلاحيات على السيرفر نفسه، سجل عمليات (مين أضاف إيه ومن أنهي جهاز)،")
    w("  تحديث تلقائي للشاشات، رفع/تنزيل الصور، وطابور أوفلاين لو الشبكة قطعت.")
    w("- 2.5.0: **نسخة احتياطية كاملة + شاشة إدارة النسخ** — نسخة كاملة (قاعدة + صور) في zip،")
    w("  استرجاع موحّد (نسخة أمان قبل الاستبدال + حفظ المستخدمين + دمج الصور)، وشاشة إدارة")
    w("  النسخ (عرض/إنشاء سريعة وكاملة/استرجاع/حذف/فتح المجلد) + دعم كامل على الشبكة.")
    w("- 2.6.0: **ملاحظات وتعليمات الوحدات** — المدير يكتب ملاحظة ويحدد الوحدة المستهدفة والأولوية،")
    w("  وتظهر عند الوحدة في شريط واضح أعلى الشاشة مع زر «قدمها»؛ الضغط عليه يحوّل الحالة لـ «تم»")
    w("  على السيرفر ويسجّل مين قدّمها وامتى. جدول `notes` + نقطة `/note_done` على الشبكة.")
    w("- 2.6.1: **تنبيه بصوت** عند وصول ملاحظة جديدة (`_play_note_sound`/`_beep`) — العاجلة بجرسين،")
    w("  ومفتاح تشغيل/إيقاف في شاشة الملاحظات (`toggle_note_sound`) محفوظ في `settings.json` لكل جهاز.")
    w("- 2.6.2: **تنظيف الصفحة الرئيسية** — اتشال صندوق «ملخص ذكي» اللي كان بياخد مساحة فوق الكروت")
    w("  (`show_home`)، والملخص الذكي لسه موجود جوه «مقارنة الفترات» (`insight_box`).")
    w("- 2.7.0: **محرّك المظهر والثيمات** — شاشة `appearance_dialog` فيها 8 ثيمات (ألوان مستخرجة من")
    w("  صور المستخدم) + 6 خطوط (`UI_FONTS`/`apply_font`) + 3 أحجام (`FONT_SIZES`/`apply_font_size`")
    w("  عبر tk scaling) + 3 أشكال أزرار (`BTN_STYLES`/`apply_btn_style`، و`btn()` بيحترمها) + 3 كثافات")
    w("  (`UI_STYLES`). تغيير الثيم = `apply_theme` + `rebuild_ui`. تلوين نص الشريط تلقائيًا")
    w("  (`HDR_TEXT`/`HDR_SUB`) والقائمة (`SIDE_TEXT`/`SIDE_MUTED`/`SIDE_ACCENT`). اختصارات:")
    w("  `bind_keys` (F1–F11 أقسام، F12 مظهر، Alt+→/← `cycle_page`، Esc إغلاق).")
    w("- 2.8.0: **الواجهة تبان كاملة على أي شاشة** — `fit_to_screen()` بيقيس الشاشة ويفتح بمقاس")
    w("  مضمون إنه يدخل فيها. منطقة المحتوى جوه `Canvas` بشريطي تمرير (`content_canvas`/")
    w("  `content_vsb`/`content_hsb`) + `_fit_content` + `_show_scrollbars` (بيخفي الشريط لو مش محتاجه).")
    w("  الهيدر متكيّف (`layout`) والقائمة بتضيق (`_side_width`/`_on_root_config`). `toggle_fullscreen`")
    w("  (قائمة المستخدم + `Ctrl+F`). **إصلاح**: اختصارات `F1…F11` كانت `bind(\"F1\")` بدون أقواس")
    w("  فماكانتش شغّالة → `bind(\"<F1>\")`. نافذة `shortcuts_dialog` (`Shift+F1`). النوافذ المنبثقة")
    w("  بتتصغّر لو أكبر من الشاشة (`_clamp_size`/`_minsize`).")
    w("")
    w("### لسه (Pending) — ابدأ من هنا")
    w("")
    w("| # | المطلوب | الملف/الدوال |")
    w("|---|---|---|")
    w("| 1 | ترجمة تقارير Excel والطباعة للإنجليزي | `i18n.py` + `build_report` + `report_sections` + `gov_att_sections` + `_style_sheet` |")
    w("| 2 | تشكيل حقول الإدخال (Entry) على لينكس | `i18n.install()` في `i18n.py` |")
    w("| 3 | إعادة هيكلة الملف الواحد لموديولات | تقسيم `center_management.py` |")
    w("| 4 | تحسين الشبكة (اختياري) | إشعار لحظي بدل polling · قناة آمنة TLS · صلاحية أوفلاين أدق (lastrowid) — `net.py` |")
    w("")
    w("## 2) خريطة الملفات")
    w("")
    w("| الملف | الحجم/الأسطر | الدور |")
    w("|---|---|---|")
    w(f"| `center_management.py` | {cm_lines} سطر | كل منطق البرنامج والواجهة |")
    w(f"| `i18n.py` | {i18n_lines} سطر | القاموس عربي/إنجليزي + تشكيل العربي + patch واجهة Tk |")
    w(f"| `net.py` | {net_lines} سطر | طبقة الشبكة: سيرفر LAN + عميل + صلاحيات + سجل عمليات |")
    w("| `libs/` | — | `arabic_reshaper`, `bidi`, `openpyxl` (مضمّنة) |")
    w("| `img/`, `logo.ico` | — | الأيقونات |")
    w("| `photos/` | — | صور العاملين (بيانات عملاء — لا ترفع) |")
    w("| `center_management.db` | — | قاعدة البيانات (بيانات عملاء — لا ترفع) |")
    w("| `net_settings.json` | — | إعدادات الشبكة (وضع/عنوان/منفذ/توكن) — يُنشأ تلقائيًا |")
    w("| `KhadamatMisr.spec`, `build_exe.bat` | — | إعداد بناء الـ exe |")
    w("| `.github/workflows/` | — | بناء ويندوز على GitHub Actions |")
    w("| `screenshots/` | — | اللقطات + المعرض |")
    w("")
    w("## 3) خريطة الدوال في `center_management.py`")
    w("")
    w("> استخدم الأرقام دي للانتقال المباشر للدالة بدل قراءة الملف كله.")
    w("")
    w("### الدوال العامة (خارج الكلاس)")
    w("")
    w("| الدالة | السطر |")
    w("|---|---|")
    for kind, name, ln, _ in cm:
        if kind == "FUNC":
            w(f"| `{name}` | {ln} |")
    w("")
    for kind, name, ln, methods in cm:
        if kind == "CLASS":
            w(f"### الكلاس `{name}` (سطر {ln})")
            w("")
            w("| الميثود | السطر |")
            w("|---|---|")
            for mname, mln in methods:
                w(f"| `{mname}` | {mln} |")
            w("")
    w("## 4) أقسام الصفحات (فين كل صفحة بتتبني)")
    w("")
    w("| الصفحة | الدوال |")
    w("|---|---|")
    w("| الرئيسية | `show_home` · `stat_card` · `quick_info` |")
    w("| العاملون | `show_workers` · `load_workers` · `worker_dialog` |")
    w("| حضور الجهات | `show_govstaff` · `govstaff_reload` |")
    w("| الحضور والانصراف | `show_attendance` · `build_clock_tab` · `clock_now` · `build_permits_tab` · `build_absence_tab` |")
    w("| حضور الجهات (التسجيل) | `show_gov` · `gov_reload` · `gov_report_dialog` |")
    w("| الزيارات | `show_visits` |")
    w("| الصيانة | `show_maintenance` |")
    w("| الجرد | `show_inventory` · `save_inventory` · `edit_cell` |")
    w("| المشتريات | `show_purchases` · `purchase_dialog` |")
    w("| الملاحظات | `show_employees` |")
    w("| ملاحظات الوحدات | `show_notes` · `load_notes` · `save_note` · `delete_note` · `mark_note_done` · `build_notes_banner` |")
    w("| المقارنة | `show_compare` · `run_compare` |")
    w("| التقارير | `show_reports` · `build_report` |")
    w("| المظهر والثيمات | `appearance_dialog` · `apply_appearance` · `apply_theme` · `apply_font` · `apply_font_size` · `apply_btn_style` · `toggle_theme` · `bind_keys` · `cycle_page` |")
    w("| ملاءمة الشاشات + الكيبورد | `fit_to_screen` · `_clamp_size` · `_minsize` · `_fit_content` · `_show_scrollbars` · `_side_width` · `_on_root_config` · `_on_wheel` · `toggle_fullscreen` · `shortcuts_dialog` · `_scroll_page` |")
    w("| المستخدمون | `users_dialog` · `user_edit_dialog` |")
    w("| الشبكة (LAN) | `net_dialog` · `refresh_net_status` · `auto_refresh` · `start_net_server` (خارج الكلاس) |")
    w("| سجل العمليات | `audit_dialog` · `net.audit_rows` · `net.fetch_audit` |")
    w("| النسخ الاحتياطية | `backups_dialog` · `manual_backup` · `restore_backup` · `backup_now` · `backup_full_now` · `list_backups` · `do_restore` · `open_backup_archive` (خارج الكلاس) |")
    w("")
    w("## 5) قواعد ثابتة (لا تتغيّر)")
    w("")
    w("- التشكيل **للعرض بس**: `_shape()` في `i18n.py`، و`_NEEDS_SHAPING = not sys.platform.startswith(\"win\")`.")
    w("- **ممنوع** تشكيل أي حاجة بتتكتب في قاعدة البيانات — القراءة بترجع النص الخام (خزن `_khm_raw_cells`).")
    w("- **بيانات العملاء** (`center_management.db`, `photos/`) مستثناة من git — لا ترفع.")
    w("- بناء الـ exe على GitHub Actions فقط (`windows-latest`) — مفيش بناء ويندوز محلي.")
    w("- **الشبكة**: كل استعلامات قاعدة البيانات تمر من `db()` وحدها — في وضع المحطة تتحول لعميل شبكة.")
    w("  الصلاحيات تُفحص **على السيرفر** حسب جدول `users`، و`net_settings.json` لا يُرفع (فيه التوكن).")
    w("- **النسخ الاحتياطي**: النسخة اليومية التلقائية = قاعدة البيانات فقط (حجم صغير، آخر 14 يوم).")
    w("  النسخة الكاملة (`full_*.zip`) = قاعدة + صور. **قبل أي استرجاع** تُؤخذ نسخة أمان")
    w("  (`before_restore_*`) والمستخدمون/كلمات المرور **لا تتغير**، والصور تُدمج (لا تُحذف صور أحدث).")
    w("- **ملاحظات الوحدات** (جدول `notes`): القراءة و«قدمها» متاحين لأي وحدة؛ الإضافة/الحذف للمدير فقط.")
    w("  «قدمها» على المحطة بتمر من نقطة `/note_done` (مش كتابة مباشرة)، والشريط بيتحدّث مع `auto_refresh`.")
    w("- **ملاءمة الشاشة**: النافذة بتتظبط في `fit_to_screen()` حسب مقاس الشاشة، والمحتوى جوه `Canvas`")
    w("  بشرائط تمرير — أي صفحة أطول/أعرض من الشاشة بتتمرّر، مفيش حاجة بتتقص. لو غيّرت حجم النافذة")
    w("  راجع `_on_root_config` + `_fit_content`.")
    w("- أي تعديل: عدّل في `KhadamatMisr/` → انسخ للـ `repo/` → commit → push `khadamat-misr` → نزّل الـ artifact → أعد بناء الحزمة واللقطات.")
    w("- اختبارات قبل التسليم: `verify_clock.py` (23) · `verify_new.py` (56) · `verify_responsive.py` ·")
    w("  `audit_pages.py` · `verify_net.py` (34) · `verify_lan_app.py` (19) · `verify_lan_2proc.py` (6) ·")
    w("  `verify_backup.py` (29) · `verify_backup_gui.py` (9) · `verify_notes_gui.py` (25).")
    w("")

    io.open(OUT, "w", encoding="utf-8").write("\n".join(buf) + "\n")
    print("written:", OUT, "lines:", len(buf))


if __name__ == "__main__":
    main()
