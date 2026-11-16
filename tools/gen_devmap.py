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
    w("| آخر commit | `31014ae` — علامة «نقطة نهاية آخر تحديث» + ملف التطوير (2.3.1) |")
    w("| آخر بناء ناجح | GitHub Actions run `43841093803` |")
    w("| ملف exe | `KhadamatMisr.exe` · 16,249,406 بايت · MD5 `a3c6f59f8501db7de881d38f955ebb0e` |")
    w("| الحزمة | `KhadamatMisr_v2.3.1_windows.zip` · 636 ملف |")
    w("| عدد اللقطات | 39 لقطة مصحّحة + `all_screens.png` + `gallery.html` |")
    w("")
    w("### خلص (Done)")
    w("")
    w("- 2.1.0: أمان (كلمة مرور، `_ident()`، crash log، تشكيل العربي، أيقونات onefile).")
    w("- 2.2.0: تقارير حضور الجهات، نوعا تقرير `govatt`/`govstaff`، 7 لوحات ألوان + وضع ليلي + 3 أشكال عرض.")
    w("- 2.3.0: حضور/انصراف بضغطة واحدة بساعة الجهاز (`clock_now`).")
    w("- 2.3.1: تشكيل العربي في خانات الجداول + قيم القوائم على لينكس، مع قراءة خام نظيفة.")
    w("")
    w("### لسه (Pending) — ابدأ من هنا")
    w("")
    w("| # | المطلوب | الملف/الدوال |")
    w("|---|---|---|")
    w("| 1 | ترجمة تقارير Excel والطباعة للإنجليزي | `i18n.py` + `build_report` (1288) + `report_sections` (1103) + `gov_att_sections` (1224) + `_style_sheet` (1051) |")
    w("| 2 | تشكيل حقول الإدخال (Entry) على لينكس | `i18n.install()` في `i18n.py` |")
    w("| 3 | إعادة هيكلة الملف الواحد لموديولات | تقسيم `center_management.py` |")
    w("")
    w("## 2) خريطة الملفات")
    w("")
    w("| الملف | الحجم/الأسطر | الدور |")
    w("|---|---|---|")
    w(f"| `center_management.py` | {cm_lines} سطر | كل منطق البرنامج والواجهة |")
    w(f"| `i18n.py` | {i18n_lines} سطر | القاموس عربي/إنجليزي + تشكيل العربي + patch واجهة Tk |")
    w("| `libs/` | — | `arabic_reshaper`, `bidi`, `openpyxl` (مضمّنة) |")
    w("| `img/`, `logo.ico` | — | الأيقونات |")
    w("| `photos/` | — | صور العاملين (بيانات عملاء — لا ترفع) |")
    w("| `center_management.db` | — | قاعدة البيانات (بيانات عملاء — لا ترفع) |")
    w("| `KhadamatMisr.spec`, `build_exe.bat` | — | إعداد بناء الـ exe |")
    w("| `.github/workflows/` | — | بناء ويندوز على GitHub Actions |")
    w("| `screenshots/` | 39 | اللقطات + المعرض |")
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
    w("| الرئيسية | `show_home` 2689 · `stat_card` 2572 · `quick_info` 2601 |")
    w("| العاملون | `show_workers` 2862 · `load_workers` 2894 · `worker_dialog` 2956 |")
    w("| حضور الجهات | `show_govstaff` 4042 · `govstaff_reload` 4089 |")
    w("| الحضور والانصراف | `show_attendance` 4011 · `build_clock_tab` 4609 · `clock_now` 4701 · `build_permits_tab` 4797 · `build_absence_tab` 4877 |")
    w("| حضور الجهات (التسجيل) | `show_gov` 4147 · `gov_reload` 4189 · `gov_report_dialog` 4219 |")
    w("| الزيارات | `show_visits` 2749 |")
    w("| الصيانة | `show_maintenance` 2779 |")
    w("| الجرد | `show_inventory` 3311 · `save_inventory` 3442 · `edit_cell` 3413 |")
    w("| المشتريات | `show_purchases` 3646 · `purchase_dialog` 3836 |")
    w("| الملاحظات | `show_employees` 2809 |")
    w("| المقارنة | `show_compare` 4953 · `run_compare` 4987 |")
    w("| التقارير | `show_reports` 3959 · `build_report` 1288 |")
    w("| المظهر والألوان | `appearance_dialog` 2078 · `apply_appearance` 2126 |")
    w("| المستخدمون | `users_dialog` 2246 · `user_edit_dialog` 2307 |")
    w("")
    w("## 5) قواعد ثابتة (لا تتغيّر)")
    w("")
    w("- التشكيل **للعرض بس**: `_shape()` في `i18n.py`، و`_NEEDS_SHAPING = not sys.platform.startswith(\"win\")`.")
    w("- **ممنوع** تشكيل أي حاجة بتتكتب في قاعدة البيانات — القراءة بترجع النص الخام (خزن `_khm_raw_cells`).")
    w("- **بيانات العملاء** (`center_management.db`, `photos/`) مستثناة من git — لا ترفع.")
    w("- بناء الـ exe على GitHub Actions فقط (`windows-latest`) — مفيش بناء ويندوز محلي.")
    w("- أي تعديل: عدّل في `KhadamatMisr/` → انسخ للـ `repo/` → commit → push `khadamat-misr` → نزّل الـ artifact → أعد بناء الحزمة واللقطات.")
    w("- اختبارات قبل التسليم: `verify_clock.py` (23) · `verify_new.py` (38) · `audit_pages.py`.")
    w("")

    io.open(OUT, "w", encoding="utf-8").write("\n".join(buf) + "\n")
    print("written:", OUT, "lines:", len(buf))


if __name__ == "__main__":
    main()
