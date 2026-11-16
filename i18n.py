# -*- coding: utf-8 -*-
"""Bilingual (Arabic / English) UI layer for Khadamat Misr.

The app is Arabic-first: every label, button, menu, message box and window
title passes through :func:`T`.  When the language is English the Arabic
string is replaced by its translation; strings with no translation (data
coming from the database, names, etc.) are kept as they are.

Public API
----------
``LANG``          current language ("ar" or "en")
``is_en()``       True when the English interface is active
``T(text)``       translate one UI string (safe to call on any value)
``set_lang(code)`` switch language ("ar" / "en") without touching widgets
``install()``     install the tkinter translation + layout mirroring hooks
``MONTHS``        month names in the active language (for date pickers)
``month_index(name)`` 1..12 from a month name in either language
"""
import re as _re
import sys as _sys
import tkinter as _tk
from tkinter import ttk as _ttk, messagebox as _mb, filedialog as _fd

LANG = "ar"

# ------------------------------------------------------------------ dictionary
_DICT = {'خ': 'K',
 'تم': 'Done',
 'من': 'from',
 'إلى': 'to',
 'حفظ': 'Save',
 'خطأ': 'Error',
 ' في ': ' in ',
 ' من ': ' from ',
 'فرع ': 'Branch ',
 'قسم ': 'Department ',
 'مكرر': 'Duplicate',
 ' إلى ': ' to ',
 ' بند.': ' item(s).',
 ' جنيه': ' EGP',
 ' سجل.': ' record(s).',
 ' لـ «': ' for «',
 'إغلاق': 'Close',
 'إلغاء': 'Cancel',
 'البند': 'Item',
 'الجهة': 'Entity',
 'الدور': 'Role',
 'السبب': 'Reason',
 'السعر': 'Price',
 'الصنف': 'Item',
 'الفرق': 'Difference',
 'القسم': 'Department',
 'النوع': 'Type',
 'الوقت': 'Time',
 'تأخير': 'Lateness',
 'تأكيد': 'Confirm',
 'تنبيه': 'Warning',
 'يوجد ': 'There is ',
 '💾 حفظ': '💾 Save',
 '🔍 عرض': '🔍 View',
 ' من 5)': ' of 5)',
 'الحالة': 'Status',
 'الحضور': 'Check-in',
 'الخروج': 'Check-out',
 'الدخول': 'Check-in',
 'الدرجة': 'Severity',
 'الصورة': 'Photo',
 'العامل': 'Employee',
 'العودة': 'Return',
 'القسم:': 'Department:',
 'الكل: ': 'All: ',
 'الكمية': 'Quantity',
 'المورد': 'Supplier',
 'النسبة': 'Percentage',
 'الهاتف': 'Phone',
 'صلاحية': 'Permission',
 'متابعة': 'Continue',
 'ملاحظة': 'Note',
 'موظفو ': 'Staff of ',
 '💾 حفظ ': '💾 Save ',
 '🔍 بحث:': '🔍 Search:',
 ' دقيقة.': ' min.',
 ' سجل):\n': ' record(s)):\n',
 'إذونات ': 'Permits ',
 'التاريخ': 'Date',
 'الحالة:': 'Status:',
 'الحالية': 'current',
 'السابقة': 'Previous',
 'العنوان': 'Title',
 'الوظيفة': 'Job title',
 'تم حفظ ': 'Saved ',
 'ملاحظات': 'Notes',
 '🔓  دخول': '🔓  Check-in',
 ' الساعة ': ' at ',
 'اسم مكرر': 'Duplicate name',
 'الإجمالي': 'Total',
 'الانصراف': 'Check-out',
 'التفاصيل': 'Details',
 'الفاتورة': 'Invoice',
 'الوظيفة:': 'Job title:',
 'تسجيل - ': 'Entry - ',
 'تم الحفظ': 'Saved',
 'غير ممكن': 'Not possible',
 'من تاريخ': 'From date',
 '🖨️ طباعة': '🖨️ Print',
 ': استخدم ': ': use ',
 'إلى تاريخ': 'To date',
 'اسم الشخص': 'Person name',
 'المشتريات': 'Purchases',
 'بدء العمل': 'Start date',
 'تم تسجيل ': 'Recorded ',
 'جرد موجود': 'Stock count exists',
 'عرض الجهة': 'Show entity',
 'غير مسموح': 'Not allowed',
 'موظف سابق': 'Former staff',
 'نفس الملف': 'Same file',
 '» نهائيًا؟': '» permanently?',
 'اختر سجلًا': 'Select a record',
 'اسم العامل': 'Employee name',
 'اسم الموظف': 'Staff name',
 'الإجمالي: ': 'Total: ',
 'املأ حقل «': 'Fill in «',
 'بدون أوقات': 'no times',
 'تم الإدراج': 'Imported',
 'تم التسجيل': 'Recorded',
 'تم التعديل': 'Updated',
 'حالة اليوم': "Today's status",
 'حذف المحدد': 'Delete selected',
 'رقم الهاتف': 'Phone number',
 'سعر الوحدة': 'Unit price',
 'لا تغييرات': 'No changes',
 'وقت الحضور': 'Check-in time',
 'وقت الخروج': 'Check-out time',
 ' - المتبقي ': ' - remaining ',
 'أخطاء حالية': 'Current errors',
 'أخطاء سابقة': 'Previous errors',
 'اختر الصورة': 'Select photo',
 'اختر عاملًا': 'Select an employee',
 'اختر موظفًا': 'Select a staff member',
 'اكتب الوقت.': 'Enter the time.',
 'الملف مفتوح': 'File is open',
 'تأكيد الحذف': 'Confirm deletion',
 'تاريخ الإذن': 'Permit date',
 'تم استخراج ': 'Exported ',
 'تم التسجيل.': 'Recorded.',
 'حفظ الزيارة': 'Save visit',
 'ساعات العمل': 'Working hours',
 'سبب الزيارة': 'Visit reason',
 'سجل الصيانة': 'Maintenance log',
 'سُدد بواسطة': 'Paid by',
 'كلمة المرور': 'Password',
 'مسجل بالفعل': 'Already recorded',
 'نوع التقرير': 'Report type',
 '💾 حفظ الإذن': '💾 Save permit',
 '💾 حفظ الجرد': '💾 Save count',
 ' لهذا الشهر.': ' for this month.',
 ': الدوام من ': ': Shift from ',
 'إدراج بيانات': 'Import data',
 'إذن الخروج: ': 'Exit permit: ',
 'إشادات حالية': 'Current commendations',
 'إضافة مستخدم': 'Add user',
 'إضافة وتسجيل': 'Add and record',
 'اسم المستخدم': 'Username',
 'الرقم القومي': 'National ID',
 'انتهاء العمل': 'End date',
 'بيانات ناقصة': 'Incomplete data',
 'تسجيل الخروج': 'Log out',
 'تسجيل الدخول': 'Log in',
 'تعبئة سريعة:': 'Quick fill:',
 'تعديل مستخدم': 'Edit user',
 'تم الاستخراج': 'Exported',
 'رقم الفاتورة': 'Invoice no.',
 'عدد البنود: ': 'Items: ',
 'فتح القسم  ←': 'Open section  →',
 'لا توجد صورة': 'No photo',
 'لا يوجد رصيد': 'No balance',
 'ملف غير صالح': 'Invalid file',
 'نوع التقرير:': 'Report type:',
 '⚖️ قارن الآن': '⚖️ Compare now',
 '✅ تسجيل دخول': '✅ Check in',
 '➕ إضافة عامل': '➕ Add employee',
 '➕ موظف جديد:': '➕ New staff member:',
 '💾 حفظ الموظف': '💾 Save staff',
 '💾 حفظ الوارد': '💾 Save stock-in',
 '📋 سجل الوارد': '📋 Stock-in log',
 '📝 حفظ ملاحظة': '📝 Save note',
 '🚪 تسجيل خروج': '🚪 Check out',
 ' مدة التأخير ': ' late ',
 'اختر مستخدمًا': 'Select a user',
 'بيانات العامل': 'Employee data',
 'تأخير (دقيقة)': 'Late (min)',
 'تاريخ الميلاد': 'Date of birth',
 'تسجيل وارد - ': 'Stock-in - ',
 'تم حفظ تعديل ': 'Updated ',
 'دقائق التأخير': 'Late minutes',
 'راجع التاريخ.': 'Check the date.',
 'زيارات المركز': 'Center visits',
 'سجل الوارد - ': 'Stock-in log - ',
 'لا توجد أصناف': 'No items',
 'لا توجد سجلات': 'No records',
 'مرحبًا بك في ': 'Welcome to ',
 'مسار الصورة:\n': 'Photo path:\n',
 'ملاحظات حالية': 'Current notes',
 'نسخة احتياطية': 'Backup',
 'نوع المشتريات': 'Purchase type',
 '💾 حفظ التعديل': '💾 Save changes',
 '🗑️ حذف المحدد': '🗑️ Delete selected',
 '\nتاريخ الملف: ': '\nFile date: ',
 '   |   الصور: ': '   |   Photos: ',
 'استخراج بيانات': 'Export data',
 'الإجراء المتخذ': 'Action taken',
 'التاريخ والوقت': 'Date and time',
 'الجهة / الشركة': 'Entity / Company',
 'الفترة الحالية': 'Current period',
 'الفترة السابقة': 'Previous period',
 'المورد / المحل': 'Supplier / Shop',
 'تاريخ الاستلام': 'Received date',
 'تاريخ الفاتورة': 'Invoice date',
 'تصميم وتطوير: ': 'Designed & developed by: ',
 'تم تسجيل خروج ': 'Checked out ',
 'تم تسجيل دخول ': 'Checked in ',
 'تم تسجيل وارد ': 'Stock-in recorded ',
 'حذف المستخدم «': 'Delete user «',
 'قيمة غير صحيحة': 'Invalid value',
 'كمية غير صحيحة': 'Invalid quantity',
 'لا توجد بيانات': 'No data',
 'لا توجد فاتورة': 'No invoice',
 'معاينة الأعداد': 'Preview counts',
 '➕ إضافة مستخدم': '➕ Add user',
 '📎 إدراج فاتورة': '📎 Attach invoice',
 '📥 إدراج بيانات': '📥 Import data',
 '🖨️ طباعة تقرير': '🖨️ Print report',
 '🧠 الملخص الذكي': '🧠 Smart summary',
 '🧾 عرض الفاتورة': '🧾 View invoice',
 'إضافة عامل جديد': 'Add new employee',
 'اضغط للتفاصيل ›': 'Tap for details ›',
 'الأقسام المتاحة': 'Available sections',
 'الصفحة الرئيسية': 'Home',
 'الفني / المسؤول': 'Technician / Officer',
 'الكمية في صنف «': 'Quantity in item «',
 'تاريخ بدء العمل': 'Start date',
 'محتويات التقرير': 'Report contents',
 'ملاحظات الصيانة': 'Maintenance notes',
 '✅ اختيار وتسجيل': '✅ Select and record',
 '✏️ تعديل المحدد': '✏️ Edit selected',
 '✏️ تعديل مستخدم': '✏️ Edit user',
 '➕ إضافة للقائمة': '➕ Add to list',
 '🌙  الوضع الليلي': '🌙  Dark mode',
 '💾 حفظ التعديلات': '💾 Save changes',
 '💾 حفظ المشتريات': '💾 Save purchase',
 '📊 استخراج Excel': '📊 Export Excel',
 '📋 سجل الوارد - ': '📋 Stock-in log - ',
 '🚪  تسجيل الخروج': '🚪  Log out',
 ' صنفًا في الجرد.': ' item(s) in the count.',
 'اختر وقت الخروج.': 'Select check-out time.',
 'استبدال الفاتورة': 'Replace invoice',
 'اكتب اسم العامل.': 'Enter the employee name.',
 'اكتب اسم الموظف.': 'Enter the staff name.',
 'السجل غير موجود.': 'Record not found.',
 'القائمة الرئيسية': 'Main menu',
 'الوظيفة / المسمى': 'Job title / Position',
 'بحث باسم العامل:': 'Search by employee name:',
 'تعذر قراءة الملف': 'Could not read the file',
 'تم السداد بواسطة': 'Paid by',
 'حذف الصنف المحدد': 'Delete selected item',
 'رقم الإذن بالشهر': 'Permit no. in month',
 'فتح القسم كاملًا': 'Open full section',
 'قاعدة البيانات: ': 'Database: ',
 'ليس لديك صلاحية ': "You don't have permission ",
 'معلومات سريعة - ': 'Quick info - ',
 '📈 مقارنة الفترات': '📈 Period comparison',
 '📤 استخراج بيانات': '📤 Export data',
 '🖨️ طباعة التقرير': '🖨️ Print report',
 '» مسجل بالفعل في ': '» already recorded on ',
 'إجمالي الفاتورة: ': 'Invoice total: ',
 'بنود نفس الفاتورة': 'items of the same invoice',
 'تعذر حفظ الصورة:\n': 'Could not save the photo:\n',
 'تعذر نسخ الصورة:\n': 'Could not copy the photo:\n',
 'تغيير كلمة المرور': 'Change password',
 'تم حفظ المقارنة:\n': 'Comparison saved:\n',
 'حذف السجل المحدد؟': 'Delete the selected record?',
 'حفظ عملية الصيانة': 'Save maintenance',
 'راجع تاريخ الإذن.': 'Check the permit date.',
 'سعر الوحدة (جنيه)': 'Unit price (EGP)',
 'مسجل لهذا اليوم: ': 'Recorded for today: ',
 'ملاحظات (اختياري)': 'Notes (optional)',
 'نسخة احتياطية   💾': 'Backup   💾',
 '☀️  الوضع النهاري': '☀️  Light mode',
 '📥 تسجيل وارد جديد': '📥 Record new stock-in',
 '\nهل تريد المتابعة؟': '\nDo you want to continue?',
 'اختر صورة الفاتورة': 'Select invoice photo',
 'اسم العامل (ثلاثي)': 'Employee name (full)',
 'اكتب اسم المستخدم.': 'Enter the username.',
 'الفترة الحالية (أ)': 'Current period (A)',
 'الفترة السابقة (ب)': 'Previous period (B)',
 'تاريخ ووقت التسجيل': 'Recorded date and time',
 'تاريخ ووقت الزيارة': 'Visit date and time',
 'تعذر إنشاء الملف:\n': 'Could not create the file:\n',
 'تعذر تجهيز الجرد:\n': 'Could not prepare the count:\n',
 'تعذر تجهيز الملف:\n': 'Could not prepare the file:\n',
 'تم استخراج التقرير': 'Report exported',
 'سجل حضور وانصراف: ': 'Attendance record: ',
 '  مقارنة العاملين  ': '  Employee comparison  ',
 '  🚪 إذونات الخروج  ': '  🚪 Exit permits  ',
 'اسم الفني / المسؤول': 'Technician / Officer name',
 'الصلاحيات حسب القسم': 'Permissions by section',
 'انصراف مبكر (دقيقة)': 'Early leave (min)',
 'تسجيل مشتريات جديدة': 'Record new purchase',
 'تعذر إنشاء النسخة:\n': 'Could not create the backup:\n',
 'تم تسجيل الإذن رقم ': 'Permit no. recorded ',
 '➕ إضافة مستخدم جديد': '➕ Add new user',
 '🔑 تغيير كلمة المرور': '🔑 Change password',
 '🧾 عرض صورة الفاتورة': '🧾 View invoice image',
 '\n⛔ لا يسمح بإذن ثالث': '\n⛔ A third permit is not allowed',
 '   |   تركوا العمل: ': '   |   Left work: ',
 '   |   عدد السجلات: ': '   |   Records: ',
 ' (4 خانات على الأقل)': ' (at least 4 characters)',
 'اكتب الملاحظة أولًا.': 'Enter the note first.',
 'تعذر إنشاء التقرير:\n': 'Could not create the report:\n',
 'تعذر تجهيز الطباعة:\n': 'Could not prepare printing:\n',
 'راجع تاريخ الاستلام.': 'Check the received date.',
 'راجع تاريخ الفاتورة.': 'Check the invoice date.',
 'لا توجد بيانات صالحة': 'No valid data',
 'لم تُرفق صورة فاتورة': 'No invoice image attached',
 'لم يتم تعديل أي سجل.': 'No records were updated.',
 'نعم = حضور وانصراف «': 'Yes = attendance «',
 'وقت العودة (اختياري)': 'Return time (optional)',
 '👥 اختيار موظف وتسجيل': '👥 Select staff and record',
 '📥 تسجيل وارد جديد - ': '📥 Record new stock-in - ',
 '🔑  تغيير كلمة المرور': '🔑  Change password',
 '  📅 الغياب والتأخير  ': '  📅 Absence & lateness  ',
 ' - اختر القسم المطلوب': ' - choose the section',
 'اكتب اسم الشخص أولًا.': "Enter the person's name first.",
 'الرقم القومي (14 رقم)': 'National ID (14 digits)',
 'المستخدمون والصلاحيات': 'Users & permissions',
 'تعذر إدراج البيانات:\n': 'Could not import the data:\n',
 'هل تريد تسجيل الخروج؟': 'Do you want to log out?',
 '✏️ عرض / تعديل المحدد': '✏️ View / edit selected',
 '➕ تسجيل مشتريات جديدة': '➕ Record new purchase',
 '📎 إرفاق صورة الفاتورة': '📎 Attach invoice image',
 '🖨️ طباعة الجرد الحالي': '🖨️ Print current count',
 '🛒 تسجيل مشتريات جديدة': '🛒 Record new purchase',
 '   |   على رأس العمل: ': '   |   Active: ',
 'تعذر تجهيز المشتريات:\n': 'Could not prepare purchases:\n',
 '🏛️ حضور وانصراف الجهات': '🏛️ Entity attendance',
 '💾 حفظ الحضور والانصراف': '💾 Save attendance',
 '📊 استخراج Excel بالمدة': '📊 Export Excel by period',
 '📎 إدراج فاتورة من صورة': '📎 Attach invoice from image',
 '  ⏱️ الحضور والانصراف  ': '  ⏱️ Attendance  ',
 'إدراج نسخة احتياطية   📂': 'Restore backup   📂',
 'استخراج وطباعة التقارير': 'Reports & printing',
 'اكتب اسم الموظف الجديد.': 'Enter the new staff name.',
 'تفاصيل الملاحظة / الخطأ': 'Note / error details',
 'تم تسجيل حضور وانصراف «': 'Attendance recorded «',
 'عدد السجلات في الفترة:\n': 'Records in the period:\n',
 'هذه العملية للمدير فقط.': 'This operation is for the admin only.',
 'يوجد مستخدم بنفس الاسم.': 'A user with the same name exists.',
 '👥 المستخدمون والصلاحيات': '👥 Users & permissions',
 '👷 قاعدة بيانات العاملين': '👷 Employees database',
 'اختر الصنف أو اكتب اسمه.': 'Select an item or type its name.',
 'الإجراء المتخذ (اختياري)': 'Action taken (optional)',
 'تعذر حفظ صورة الفاتورة:\n': 'Could not save the invoice image:\n',
 'نظام إدارة المركز - فرع ': 'Center Management System - Branch ',
 'وقت الحضور (للتأخير فقط)': 'Check-in time (for lateness only)',
 '↺ المواعيد الرسمية للقسم': '↺ Official department hours',
 '👥  المستخدمون والصلاحيات': '👥  Users & permissions',
 '\n\nهل تريد فتح الملف الآن؟': '\n\nDo you want to open the file now?',
 '   |   إجمالي المشتريات: ': '   |   Purchases total: ',
 '  الملخص والرسم البياني  ': '  Summary & chart  ',
 'استخراج التقرير إلى Excel': 'Export report to Excel',
 'اسم العامل (اختر أو اكتب)': 'Employee name (select or type)',
 '📝 ملاحظات وأخطاء العاملين': '📝 Employee notes & errors',
 ' صنفًا وإضافته على الرصيد.': ' item(s) added to stock.',
 'الوقت (مثل 8:30 أو 3:00 م)': 'Time (e.g. 8:30 or 3:00 PM)',
 'راجع الفترة (يوم/شهر/سنة).': 'Check the period (day/month/year).',
 'لا يمكنك حذف حسابك الحالي.': 'You cannot delete your current account.',
 '👤 تعديل «تم السداد بواسطة»': '👤 Edit “Paid by”',
 'اكتب ملاحظات الصيانة أولًا.': 'Enter the maintenance notes first.',
 'الصنف (اختر أو اكتب يدويًا)': 'Item (select or type manually)',
 'بحث (اسم / هاتف / رقم قومي)': 'Search (name / phone / national ID)',
 'تم تغيير كلمة المرور بنجاح.': 'Password changed successfully.',
 'حدد العامل من الجدول أولًا.': 'Select the employee from the table first.',
 ' (اتركها فارغة لعدم التغيير)': ' (leave blank to keep unchanged)',
 '» لازم تكون رقم أكبر من صفر.': '» must be a number greater than zero.',
 'أدخل كمية رقمية صفر أو أكبر.': 'Enter a numeric quantity of zero or more.',
 'اكتب كمية رقمية أكبر من صفر.': 'Enter a numeric quantity greater than zero.',
 'اكتب وقت الحضور أو الانصراف.': 'Enter check-in or check-out time.',
 'بحث (اسم / هاتف / رقم قومي):': 'Search (name / phone / national ID):',
 'تأكيد كلمة المرور غير مطابق.': 'Password confirmation does not match.',
 'تم إدراج صورة الفاتورة لعدد ': 'Invoice image attached to ',
 'تم السداد بواسطة (اسم الشخص)': 'Paid by (person name)',
 'حدد سجل الوارد المطلوب حذفه.': 'Select the stock-in record to delete.',
 'حدد سجل الوارد لعرض فاتورته.': 'Select the stock-in record to view its invoice.',
 'لا يوجد موظفون بهذه الفلاتر.': 'No staff match these filters.',
 'يوجد عامل بنفس الاسم بالفعل.': 'An employee with the same name already exists.',
 '📊 استخراج الجرد الحالي Excel': '📊 Export current count to Excel',
 '🗂️ قاعدة بيانات موظفي الجهات': '🗂️ Entity staff database',
 'أدخل كمية صنف واحد على الأقل.': 'Enter a quantity for at least one item.',
 'اختر العامل من القائمة أولًا.': 'Select the employee from the list first.',
 'اختر الموظف من القائمة أولًا.': 'Select the staff member from the list first.',
 'تأكيد إدراج النسخة الاحتياطية': 'Confirm restoring the backup',
 'حدد المستخدم من الجدول أولًا.': 'Select the user from the table first.',
 'غير مسموح - لا يسمح بإذن ثالث': 'Not allowed - a third permit is not permitted',
 'لا يوجد تسجيل لهذا اليوم بعد.': 'No record for today yet.',
 'لم أجد أي سجل صالح للإدراج.\n\n': 'No valid records found to import.\n\n',
 'حدد سجل الوارد لإدراج فاتورته.': 'Select the stock-in record to attach its invoice.',
 'كلمة المرور الحالية غير صحيحة.': 'The current password is incorrect.',
 'كلمة المرور لا تقل عن 4 خانات.': 'Password must be at least 4 characters.',
 'هذا الحساب موقوف. راجع المدير.': 'This account is disabled. Contact the admin.',
 '⬇️ تعبئة الجرد بالرصيد المتوقع': '⬇️ Fill count with expected balance',
 '      (على رأس العمل بالقسم -> ': '      (active in department -> ',
 'استخراج بيانات الحضور والانصراف': 'Export attendance data',
 'اكتب اسم المستخدم وكلمة المرور.': 'Enter username and password.',
 'تم حفظ «تم السداد بواسطة» لعدد ': '“Paid by” saved for ',
 'تم حفظ نسخة من قاعدة البيانات:\n': 'Database backup saved:\n',
 'حدد بند المشتريات المطلوب حذفه.': 'Select the purchase item to delete.',
 'حدد بند المشتريات لعرض فاتورته.': 'Select the purchase item to view its invoice.',
 ' في هذا القسم.\nصلاحيتك الحالية: ': ' in this section.\nYour current permission: ',
 ' ميجابايت   |   إجمالي السجلات: ': ' MB   |   Total records: ',
 'اختر ملف النسخة الاحتياطية (.db)': 'Select the backup file (.db)',
 'اختر ملف بيانات الحضور والانصراف': 'Select the attendance data file',
 'الرقم القومي لازم يكون 14 رقمًا.': 'National ID must be 14 digits.',
 'تعذر قراءة الملف كقاعدة بيانات:\n': 'Could not read the file as a database:\n',
 'تم تسجيل البيان بالتاريخ والوقت.': 'Entry recorded with date and time.',
 'حذف سجل الحضور والانصراف المحدد؟': 'Delete the selected attendance record?',
 'تاريخ انتهاء العمل (إن ترك العمل)': 'End date (if left work)',
 'تم تسجيل الزيارة بالتاريخ والوقت.': 'Visit recorded with date and time.',
 'تم حفظ الصورة\n(اضغط عرض للمشاهدة)': 'Photo saved\n(press View to open)',
 'حدد الإذن المطلوب حذفه من الجدول.': 'Select the permit to delete from the table.',
 'حدد السجل المطلوب حذفه من الجدول.': 'Select the record to delete from the table.',
 'هل تريد حذف السجل المحدد نهائيًا؟': 'Delete the selected record permanently?',
 'حدد بند المشتريات من الجدول أولًا.': 'Select the purchase item from the table first.',
 'هذا البند لم تُرفق له صورة فاتورة.': 'No invoice image is attached to this item.',
 'هذا السجل لم تُرفق له صورة فاتورة.': 'No invoice image is attached to this record.',
 'وقت الانصراف لازم يكون بعد الحضور.': 'Check-out time must be after check-in.',
 'وقت الحضور (اكتبه يدويًا مثل 7:30)': 'Check-in time (type manually, e.g. 7:30)',
 '.\nراجع المدير لو محتاج صلاحية أعلى.': '.\nContact the admin if you need higher permission.',
 'المورد / رقم إذن الاستلام (اختياري)': 'Supplier / receipt no. (optional)',
 'تاريخ انتهاء العمل قبل تاريخ البدء.': 'End date is before the start date.',
 ' غير صحيح، اكتبه بالشكل يوم/شهر/سنة.': ' is invalid, use day/month/year format.',
 'اكتب سعرًا رقميًا (أو اتركه فارغًا).': 'Enter a numeric price (or leave it blank).',
 'تم إدراج صورة الفاتورة للسجل المحدد.': 'Invoice image attached to the selected record.',
 'لازم يفضل مدير واحد مفعّل على الأقل.': 'At least one active admin must remain.',
 'وقت الخروج لازم يكون بعد وقت الدخول.': 'Check-out time must be after check-in time.',
 'وقت العودة لازم يكون بعد وقت الخروج.': 'Return time must be after check-out time.',
 '🕒 الحضور والانصراف والإذونات والغياب': '🕒 Attendance, permits & absence',
 'كلمة المرور الجديدة لا تقل عن 4 خانات.': 'The new password must be at least 4 characters.',
 'ملف الفاتورة غير موجود في مجلد photos.': 'The invoice file was not found in the photos folder.',
 'وقت الانصراف (اكتبه يدويًا مثل 4:00 م)': 'Check-out time (type manually, e.g. 4:00 PM)',
 'تم تسجيل عملية الصيانة بالتاريخ والوقت.': 'Maintenance recorded with date and time.',
 'حدد السجل المطلوب حذفه من الجدول أولًا.': 'Select the record to delete from the table first.',
 'اكتب اسم الشخص الذي سدّد قيمة المشتريات:': 'Enter the name of the person who paid for the purchase:',
 'الحساب مفعّل (إلغاء التفعيل يمنع الدخول)': 'Account active (disabling blocks login)',
 'حدد العامل المطلوب حذفه من الجدول أولًا.': 'Select the employee to delete from the table first.',
 'حدد السجل المطلوب تعديله من الجدول أولًا.': 'Select the record to edit from the table first.',
 '» فقط\nلا = كل الجهات الثمانية\nإلغاء = رجوع': '» only\nNo = all eight entities\nCancel = back',
 'اكتب اسم العامل وتفاصيل الملاحظة أو الخطأ.': 'Enter the employee name and note/error details.',
 'راجع تاريخ البداية والنهاية (يوم/شهر/سنة).': 'Check the start and end dates (day/month/year).',
 'اكتب كمية صنف واحد على الأقل (أكبر من صفر).': 'Enter a quantity greater than zero for at least one item.',
 'وقت الدخول لازم يكون قبل وقت الخروج المسجل.': 'Check-in time must be before the recorded check-out.',
 'الملف مفتوح في Excel. أغلقه ثم أعد المحاولة.': 'The file is open in Excel. Close it and try again.',
 ' فقط لكل عامل خلال الشهر   |   بداية الدوام: ': ' only per employee per month   |   Shift start: ',
 'الملف مفتوح في Excel. أغلقه أو اختر اسمًا آخر.': 'The file is open in Excel. Close it or choose another name.',
 'تم إدخال بيانات خاطئة 5 مرات. سيُغلق البرنامج.': 'Wrong data entered 5 times. The program will close.',
 'سجّل العاملين أولًا من «قاعدة بيانات العاملين»': 'Add employees first from “Employees database”',
 'اختر العامل لعرض المتبقي من إذونات الشهر (الحد ': 'Select the employee to see remaining monthly permits (limit ',
 'حذف الإذن؟ (سيعود للعامل رصيد إذن في هذا الشهر)': 'Delete the permit? (the employee regains a permit this month)',
 'اسم المستخدم أو كلمة المرور غير صحيحة. (المحاولة ': 'Incorrect username or password. (Attempt ',
 'الفترة الافتراضية تشمل كل السجلات، وتقدر تغيّرها.': 'The default period covers all records; you can change it.',
 'حذف سجل الوارد المحدد؟ (سيُخصم من الرصيد المتوقع)': 'Delete the selected stock-in record? (will be deducted from '
                                                      'expected balance)',
 'تم حفظ التعديل (تاريخ ووقت التسجيل الأصلي كما هو).': 'Changes saved (original date/time preserved).',
 'هذا الملف ليس نسخة سليمة من قاعدة بيانات البرنامج.': 'This file is not a valid copy of the program database.',
 ' في هذا اليوم.\nلو عايز تغيّره عدّل السجل من الجدول.': ' for this day.\nTo change it, edit the record in the table.',
 'اكتب التاريخ بالشكل يوم/شهر/سنة وتأكد من صحة الفترة.': 'Enter the date as day/month/year and check the period.',
 'الرقم القومي ليس 14 رقمًا. هل تريد الحفظ على أي حال؟': 'National ID is not 14 digits. Save anyway?',
 'تعذر إدراج النسخة (تم الإبقاء على بياناتك الحالية):\n': 'Could not restore the backup (your current data was '
                                                          'kept):\n',
 'سيتم استبدال كل البيانات الحالية ببيانات هذه النسخة:\n': 'All current data will be replaced with this backup:\n',
 '🧠 ملخص ذكي: هذا الشهر مقابل نفس الفترة من الشهر الماضي': '🧠 Smart summary: this month vs. the same period last month',
 '» مسجل بحالة «ترك العمل». هل تريد إعادته على رأس العمل؟': '» is marked as “Left work”. Restore to active?',
 'أدخل بيانات الزائر، وسيتم تسجيل التاريخ والوقت تلقائيًا': 'Enter the visitor data; date and time are recorded '
                                                            'automatically',
 'حدد بند المشتريات المراد إدراج فاتورته من الجدول أولًا.': 'Select the purchase item to attach its invoice from the '
                                                            'table first.',
 'مدير (كل الصلاحيات + إدارة المستخدمين والنسخ الاحتياطي)': 'Admin (full permissions + users & backups)',
 'أضف صنفًا واحدًا على الأقل للقائمة (زر «إضافة للقائمة»).': 'Add at least one item to the list (“Add to list”).',
 'ملف بيانات الحضور والانصراف (للنسخ أو التعديل ثم الإدراج):': 'Attendance data file (copy/edit then import):',
 ' ميجابايت   |   نسخة احتياطية يومية تلقائية في مجلد backups': ' MB   |   automatic daily backup in the backups '
                                                                'folder',
 'الملف مفتوح في Excel. أغلقه أو اختر اسمًا آخر ثم أعد المحاولة.': 'The file is open in Excel. Close it or choose '
                                                                   'another name, then retry.',
 'الأمن: 7ص - 7م   |   الفنيين: 7:30ص - 4م   |   الخدمات: 7ص - 4م': 'Security: 7AM - 7PM   |   Technicians: 7:30AM - '
                                                                    '4PM   |   Services: 7AM - 4PM',
 'هذا البند له فاتورة مرفقة بالفعل. هل تريد استبدالها بصورة جديدة؟': 'This item already has an invoice. Replace it '
                                                                     'with a new image?',
 'تم إدراج النسخة الاحتياطية بنجاح.\nنسخة بياناتك السابقة محفوظة في:\n': 'Backup restored successfully.\n'
                                                                         'Your previous data was saved at:\n',
 'حدد الفترة ونوع التقرير، راجع الأعداد ثم استخرج ملف Excel أو اطبعه': 'Choose the period and report type, review '
                                                                       'counts, then export Excel or print',
 'هذا هو ملف قاعدة البيانات الحالي نفسه. اختر ملف نسخة احتياطية آخر.': 'This is the current database file. Choose '
                                                                       'another backup file.',
 'إضافة الكميات إلى المخزون كوارد (للأصناف الموجودة في قائمة الجرد فقط)': 'Add quantities to stock as received (only '
                                                                          'items in the count list)',
 'سجّل مشتريات الورقيات والأحبار والأدوات يدويًا مع إرفاق صورة الفاتورة': 'Record paper, toner and tools purchases '
                                                                          'manually with an invoice image',
 'اضغط مرتين على «الكمية» أو «ملاحظات» للتعديل، ثم اضغط «حفظ التعديلات».': 'Double-click “Quantity” or “Notes” to '
                                                                           'edit, then press “Save changes”.',
 'راجع تواريخ الفترتين (يوم/شهر/سنة) وتأكد أن تاريخ النهاية بعد البداية.': 'Check both periods (day/month/year) and '
                                                                           'make sure the end date is after the start.',
 '\n\nملاحظة: صور العاملين والفواتير في فولدر photos لا تتغير بهذه العملية.': '\n'
                                                                              '\n'
                                                                              'Note: employee and invoice photos in '
                                                                              'the photos folder are not affected.',
 'سجل أعمال الصيانة مع الجهة أو الفني والملاحظات، والتاريخ والوقت تلقائيًا': 'Log maintenance work with the '
                                                                             'entity/technician and notes; date and '
                                                                             'time are automatic',
 ' بنود بنفس التاريخ والمورد ورقم الفاتورة.\nهل تريد تطبيق الاسم عليها كلها؟': ' items share the same date, supplier '
                                                                               'and invoice no.\n'
                                                                               'Apply the name to all of them?',
 'أضف مستخدمًا وحدد لكل قسم: ممنوع / عرض فقط / عرض + إضافة / تعديل وحذف كامل.': 'Add a user and set each section: None '
                                                                                '/ View only / View + add / Full edit '
                                                                                '& delete.',
 ' بنود بنفس التاريخ والمورد ورقم الفاتورة.\nهل تريد إدراج هذه الصورة لها كلها؟': ' items share the same date, '
                                                                                  'supplier and invoice no.\n'
                                                                                  'Attach this image to all of them?',
 'لا توجد أرصدة سابقة أو وارد مسجل بعد. أدخل الجرد يدويًا أو سجّل الوارد أولًا.': 'No previous balances or stock-in '
                                                                                  'yet. Enter the count manually or '
                                                                                  'record stock-in first.',
 'يوجد جرد محفوظ لنفس الفترة. هل تريد استبداله؟\n(لتجنب تكرار الأصناف في التقرير)': 'A count is already saved for this '
                                                                                    'period. Replace it?\n'
                                                                                    '(To avoid duplicate items in the '
                                                                                    'report)',
 'حذف بند المشتريات المحدد؟\n(لو كان مضافًا للمخزون كوارد فسيُخصم من الرصيد أيضًا)': 'Delete the selected purchase '
                                                                                     'item?\n'
                                                                                     '(If it was added to stock, it '
                                                                                     'will be deducted too)',
 'تحديث السجلات الموجودة يحتاج صلاحية التعديل، لذلك سيتم تخطيها وإضافة الجديد فقط.': 'Updating existing records needs '
                                                                                     'edit permission, so they will be '
                                                                                     'skipped and only new ones added.',
 '\n\nسيحفظ البرنامج نسخة من بياناتك الحالية تلقائيًا قبل الاستبدال.\nهل تريد المتابعة؟': '\n'
                                                                                          '\n'
                                                                                          'The program will '
                                                                                          'automatically back up your '
                                                                                          'current data before '
                                                                                          'replacing.\n'
                                                                                          'Continue?',
 'اضغط مرتين على الموظف (أو اختره ثم «اختيار وتسجيل»). لو موظف جديد اكتب اسمه في الأسفل.': 'Double-click the staff '
                                                                                           'member (or select then '
                                                                                           '“Select and record”). For '
                                                                                           'a new one, type the name '
                                                                                           'below.',
 'يُضاف الوارد على الرصيد الموجود. اكتب الكمية المستلمة لكل صنف فقط واترك الباقي فارغًا.': 'Stock-in is added to the '
                                                                                           'current balance. Enter '
                                                                                           'received quantity for each '
                                                                                           'item only, leave the rest '
                                                                                           'blank.',
 'قارن أي فترتين: الزيارات والصيانة والأخطاء والإشادات وأداء كل عامل - مع ملخص ذكي ورسم بياني': 'Compare any two '
                                                                                                'periods: visits, '
                                                                                                'maintenance, errors, '
                                                                                                'commendations and '
                                                                                                "each employee's "
                                                                                                'performance - with a '
                                                                                                'smart summary and '
                                                                                                'chart',
 'أضف الأصناف واحدًا واحدًا (اختر من القائمة أو اكتب الاسم يدويًا) ثم أرفق صورة الفاتورة واحفظ.': 'Add items one by '
                                                                                                  'one (pick from the '
                                                                                                  'list or type the '
                                                                                                  'name), attach the '
                                                                                                  'invoice image and '
                                                                                                  'save.',
 'سجّل الوارد الشهري فيُضاف على الرصيد تلقائيًا، ثم أدخل الجرد الفعلي ليظهر الاستهلاك في التقرير': 'Record monthly '
                                                                                                   'stock-in to add it '
                                                                                                   'to the balance '
                                                                                                   'automatically, '
                                                                                                   'then enter the '
                                                                                                   'actual count to '
                                                                                                   'show consumption '
                                                                                                   'in the report',
 'سجّل الملاحظة أو الخطأ أو التنبيه أو الإشادة لكل عامل، ثم استخرج تقريرًا بالفترة من صفحة التقارير': 'Record a note, '
                                                                                                      'error, warning '
                                                                                                      'or commendation '
                                                                                                      'for each '
                                                                                                      'employee, then '
                                                                                                      'export a period '
                                                                                                      'report from '
                                                                                                      'Reports',
 '\nنعم = تحديث السجلات الموجودة بالبيانات الجديدة\nلا = تخطي الموجودة وإضافة الجديد فقط\nإلغاء = إيقاف العملية': '\n'
                                                                                                                  'Yes '
                                                                                                                  '= '
                                                                                                                  'update '
                                                                                                                  'existing '
                                                                                                                  'records '
                                                                                                                  'with '
                                                                                                                  'new '
                                                                                                                  'data\n'
                                                                                                                  'No '
                                                                                                                  '= '
                                                                                                                  'skip '
                                                                                                                  'existing '
                                                                                                                  'and '
                                                                                                                  'add '
                                                                                                                  'only '
                                                                                                                  'new\n'
                                                                                                                  'Cancel '
                                                                                                                  '= '
                                                                                                                  'stop',
 'اضغط على اسم الجهة فتفتح قائمة موظفيها، اختر الموظف وسجّل دخول أو خروج أو ملاحظة (ولو موظف جديد اكتب اسمه)': 'Click '
                                                                                                               'an '
                                                                                                               'entity '
                                                                                                               'name '
                                                                                                               'to '
                                                                                                               'open '
                                                                                                               'its '
                                                                                                               'staff '
                                                                                                               'list, '
                                                                                                               'select '
                                                                                                               'the '
                                                                                                               'staff '
                                                                                                               'member '
                                                                                                               'and '
                                                                                                               'record '
                                                                                                               'check-in, '
                                                                                                               'check-out '
                                                                                                               'or a '
                                                                                                               'note '
                                                                                                               '(type '
                                                                                                               'a name '
                                                                                                               'for a '
                                                                                                               'new '
                                                                                                               'one)',
 'موظفو الجهات الحكومية الثماني: الاسم والوظيفة والهاتف والرقم القومي والحالة (أضفهم هنا أو مباشرة من صفحة الحضور)': 'Staff '
                                                                                                                     'of '
                                                                                                                     'the '
                                                                                                                     'eight '
                                                                                                                     'government '
                                                                                                                     'entities: '
                                                                                                                     'name, '
                                                                                                                     'job, '
                                                                                                                     'phone, '
                                                                                                                     'national '
                                                                                                                     'ID '
                                                                                                                     'and '
                                                                                                                     'status '
                                                                                                                     '(add '
                                                                                                                     'here '
                                                                                                                     'or '
                                                                                                                     'from '
                                                                                                                     'the '
                                                                                                                     'attendance '
                                                                                                                     'page)',
 'بيانات كل عامل كاملة: القسم والصور وبداية العمل وحالته (على رأس العمل / ترك العمل) - اضغط مرتين على العامل لعرض بياناته وتعديلها': 'Full '
                                                                                                                                     'employee '
                                                                                                                                     'data: '
                                                                                                                                     'department, '
                                                                                                                                     'photos, '
                                                                                                                                     'start '
                                                                                                                                     'date '
                                                                                                                                     'and '
                                                                                                                                     'status '
                                                                                                                                     '(Active '
                                                                                                                                     '/ '
                                                                                                                                     'Left) '
                                                                                                                                     '- '
                                                                                                                                     'double-click '
                                                                                                                                     'an '
                                                                                                                                     'employee '
                                                                                                                                     'to '
                                                                                                                                     'view '
                                                                                                                                     'and '
                                                                                                                                     'edit',
 'حذف العامل نهائيًا من القاعدة؟\n(الأفضل لمن ترك العمل تغيير حالته إلى «ترك العمل» بدل الحذف)\nملاحظاته السابقة تظل محفوظة في السجل.': 'Delete '
                                                                                                                                        'the '
                                                                                                                                        'employee '
                                                                                                                                        'permanently?\n'
                                                                                                                                        '(Better '
                                                                                                                                        'to '
                                                                                                                                        'set '
                                                                                                                                        'status '
                                                                                                                                        'to '
                                                                                                                                        '“Left '
                                                                                                                                        'work” '
                                                                                                                                        'instead)\n'
                                                                                                                                        'Previous '
                                                                                                                                        'notes '
                                                                                                                                        'stay '
                                                                                                                                        'in '
                                                                                                                                        'the '
                                                                                                                                        'log.',
 'الرصيد المتوقع = كمية آخر جرد + الوارد المسجل بعد تاريخه.   اضغط مرتين على «الجرد الفعلي» أو «ملاحظات» للتعديل (Enter حفظ - Esc إلغاء).': 'Expected '
                                                                                                                                            'balance '
                                                                                                                                            '= '
                                                                                                                                            'last '
                                                                                                                                            'count '
                                                                                                                                            '+ '
                                                                                                                                            'stock-in '
                                                                                                                                            'after '
                                                                                                                                            'its '
                                                                                                                                            'date.   '
                                                                                                                                            'Double-click '
                                                                                                                                            '“Actual '
                                                                                                                                            'count” '
                                                                                                                                            'or '
                                                                                                                                            '“Notes” '
                                                                                                                                            'to '
                                                                                                                                            'edit '
                                                                                                                                            '(Enter '
                                                                                                                                            '= '
                                                                                                                                            'save '
                                                                                                                                            '- '
                                                                                                                                            'Esc '
                                                                                                                                            '= '
                                                                                                                                            'cancel).',
 'خدمات مصر': 'Khadamat Misr',
 'الإسكندرية': 'Alexandria',
 'حازم فتيحه': 'Hazem Fteha',
 'نظام إدارة المركز': 'Center Management System',
 'الرئيسية': 'Home',
 'قاعدة بيانات العاملين': 'Employees database',
 'قاعدة بيانات موظفي الجهات': 'Entity staff database',
 'حضور وانصراف العاملين': 'Employee attendance',
 'حضور وانصراف الجهات': 'Entity attendance',
 'الصيانة': 'Maintenance',
 'جرد خامات النظافة': 'Cleaning supplies count',
 'جرد الأدوات الورقية': 'Paper supplies count',
 'جرد الأحبار': 'Toner count',
 'ملاحظات وأخطاء العاملين': 'Employee notes & errors',
 'مقارنة الفترات': 'Period comparison',
 'الحضور والانصراف': 'Attendance',
 'حضور وانصراف': 'Attendance',
 'إذونات الخروج': 'Exit permits',
 'الغياب والتأخير': 'Absence & lateness',
 'مقارنة العاملين': 'Employee comparison',
 'ممنوع': 'None',
 'عرض فقط': 'View only',
 'عرض + إضافة': 'View + add',
 'تعديل وحذف كامل': 'Full edit & delete',
 'الأمن': 'Security',
 'الخدمات': 'Services',
 'الفنيين': 'Technicians',
 'أمن': 'Security',
 'خدمات': 'Services',
 'صيانة': 'Maintenance',
 'على رأس العمل': 'Active',
 'ترك العمل': 'Left work',
 'مدير': 'Admin',
 'مستخدم': 'User',
 'مفعّل': 'Active',
 'موقوف': 'Disabled',
 'لا شيء': 'None',
 'كل الأقسام + إدارة المستخدمين والنسخ الاحتياطي': 'All sections + users & backups',
 'تنبيه شفهي': 'Verbal warning',
 'إشادة': 'Commendation',
 'بسيطة': 'Low',
 'متوسطة': 'Medium',
 'جسيمة': 'Severe',
 'غياب بدون إذن': 'Unexcused absence',
 'غياب بعذر': 'Excused absence',
 'إجازة مرضية': 'Sick leave',
 'إجازة اعتيادية': 'Annual leave',
 'جرد النظافة': 'Cleaning count',
 'الجرد الورقي': 'Paper count',
 'ورقيات': 'Paper',
 'أحبار': 'Toner',
 'أدوات': 'Tools',
 'متوفر': 'Available',
 'نفد': 'Out of stock',
 'لم يُجرد': 'Not counted',
 'لم تُجرد بعد': 'Not counted yet',
 'الجرد الفعلي': 'Actual count',
 'يناير': 'January',
 'فبراير': 'February',
 'مارس': 'March',
 'أبريل': 'April',
 'مايو': 'May',
 'يونيو': 'June',
 'يوليو': 'July',
 'أغسطس': 'August',
 'سبتمبر': 'September',
 'أكتوبر': 'October',
 'نوفمبر': 'November',
 'ديسمبر': 'December',
 'الإثنين': 'Monday',
 'الثلاثاء': 'Tuesday',
 'الأربعاء': 'Wednesday',
 'الخميس': 'Thursday',
 'الجمعة': 'Friday',
 'السبت': 'Saturday',
 'الأحد': 'Sunday',
 'السجل التجاري': 'Commercial Registry',
 'نيابة الأسرة': 'Family Prosecution',
 'المرور': 'Traffic',
 'التأمينات الاجتماعية': 'Social Insurance',
 'الجوازات': 'Passports',
 'الشهر العقاري': 'Real Estate Registry',
 'الأحوال المدنية': 'Civil Status',
 'الأدلة الجنائية': 'Forensic Evidence',
 'عدد': 'No. of',
 'العدد': 'Count',
 'بداية الدوام': 'Shift start',
 'نهاية الدوام': 'Shift end',
 'إجمالي': 'Total',
 'اسم': 'Name',
 'اسم الجهة': 'Entity name',
 'الشخص': 'Person',
 'وارد': 'Stock-in',
 'جرد': 'Count',
 'رصيد': 'Balance',
 'استهلاك': 'Consumption',
 'مسجل': 'Recorded',
 'محفوظ': 'Saved',
 'ملغي': 'Cancelled',
 'السجلات': 'records',
 'عدد السجلات': 'Records',
 'إجمالي السجلات': 'Total records',
 'قاعدة البيانات': 'Database',
 'ميجابايت': 'MB',
 'نسخة احتياطية يومية تلقائية في مجلد backups': 'automatic daily backup in the backups folder',
 'كل السجلات': 'All records',
 'بداية الشهر': 'start of the month',
 'هذا الشهر': 'this month',
 'الشهر الماضي': 'last month',
 'بيانات كاملة بالصور وبداية العمل والحالة (أمن - خدمات - صيانة)': 'Full data with photos, start date and status '
                                                                   '(Security - Services - Maintenance)',
 'تسجيل الزائر والجهة وسبب الزيارة بالتاريخ والوقت': 'Record visitor, entity and visit reason with date and time',
 'أعمال الصيانة وملاحظاتها وتاريخ تنفيذها': 'Maintenance work, notes and completion date',
 'ملاحظات وأخطاء وتنبيهات وإشادات كل عامل': 'Notes, errors, warnings and commendations for each employee',
 'رصيد الخامات وجرد كل أسبوع': 'Supplies balance and weekly count',
 'أرصدة الورق والأدوات المكتبية': 'Paper and office supplies balances',
 'الأحبار والتونرات حسب الطابعة': 'Toner and cartridges by printer',
 'ورقيات وأحبار وأدوات - إضافة يدوية مع صورة الفاتورة': 'Paper, toner & tools - manual entry with invoice image',
 'حضور وانصراف يدوي + إذنان شهريًا + الغياب والتأخير': 'Manual attendance + 2 permits/month + absence & lateness',
 'موظفو الجهات الثماني: البيانات والوظيفة والهاتف والحالة': 'Staff of the eight entities: data, job, phone & status',
 'اختر الجهة ثم الموظف وسجّل دخول أو خروج أو ملاحظة': 'Choose the entity then staff, and record check-in, check-out or '
                                                      'a note',
 'قارن شهرين أو أسبوعين مع رسم بياني وملخص ذكي': 'Compare two months or weeks with a chart and smart summary',
 'Excel أو طباعة جاهزة لأي فترة': 'Ready Excel or print for any period',
 'الجرد الفعلي الجديد': 'New actual count',
 'كمية آخر جرد': 'Last count qty',
 'وارد بعد آخر جرد': 'Stock-in after last count',
 'الرصيد المتوقع': 'Expected balance',
 'تاريخ آخر جرد': 'Last count date',
 'الموجود حاليًا بالمخزن': 'Currently in stock',
 'الجرد الفعلي (للكتابة)': 'Actual count (handwrite)',
 'الرصيد السابق': 'Previous balance',
 'الاستهلاك': 'Consumption',
 'الوارد': 'Stock-in',
 'م': 'No.',
 '• الزيارات: اسم الشخص، الجهة، سبب الزيارة، التاريخ والوقت.': '• Visits: person, entity, visit reason, date and time.',
 '• الصيانة: الفني أو المسؤول، الشركة، ملاحظات الصيانة، التاريخ والوقت.': '• Maintenance: technician/officer, company, '
                                                                          'maintenance notes, date and time.',
 '• الجرد: الصنف، الكمية، الرصيد السابق، الوارد، والاستهلاك (السابق + الوارد - الجرد الجديد). والوارد للمخازن له تقرير مستقل.': '• '
                                                                                                                                'Count: '
                                                                                                                                'item, '
                                                                                                                                'qty, '
                                                                                                                                'previous '
                                                                                                                                'balance, '
                                                                                                                                'stock-in '
                                                                                                                                'and '
                                                                                                                                'consumption '
                                                                                                                                '(previous '
                                                                                                                                '+ '
                                                                                                                                'stock-in '
                                                                                                                                '- '
                                                                                                                                'new '
                                                                                                                                'count). '
                                                                                                                                'Stock-in '
                                                                                                                                'has '
                                                                                                                                'its '
                                                                                                                                'own '
                                                                                                                                'report.',
 '• ملاحظات وأخطاء العاملين: تفاصيل كل بيان + ورقة ملخص بعدد ملاحظات وأخطاء كل عامل.': '• Employee notes & errors: '
                                                                                       'details of each entry + a '
                                                                                       'summary sheet of notes and '
                                                                                       'errors per employee.',
 '• إذونات الخروج والغياب والتأخير: تفاصيل كل بيان + ملخص بعدد إذونات وغياب وتأخير كل عامل.': '• Exit permits, absence '
                                                                                              '& lateness: details of '
                                                                                              'each entry + a summary '
                                                                                              'per employee.',
 '• كل تقرير مكتوب فيه: خدمات مصر - فرع الإسكندرية.': '• Every report shows: Khadamat Misr - Alexandria Branch.',
 '• 🖨️ الطباعة: تفتح صفحة منسقة بلوجو خدمات مصر وخانات التوقيع وتبدأ نافذة الطباعة (ويمكن حفظها PDF).': '• 🖨️ Print: '
                                                                                                        'opens a '
                                                                                                        'styled page '
                                                                                                        'with the '
                                                                                                        'Khadamat Misr '
                                                                                                        'logo and '
                                                                                                        'signature '
                                                                                                        'boxes and '
                                                                                                        'starts the '
                                                                                                        'print dialog '
                                                                                                        '(or save as '
                                                                                                        'PDF).',
 '• اضغط «معاينة الأعداد» لترى عدد السجلات في كل قسم قبل الاستخراج.': '• Press “Preview counts” to see records per '
                                                                      'section before exporting.',
 'مقابل': 'vs.',
 'جديد - لم يكن في الفترة السابقة': 'new - not in the previous period',
 'أكثر العاملين أخطاءً في الفترة الحالية': 'Most errors in the current period',
 'الأكثر إشادةً': 'Most commended',
 'أكثر قسم أخطاءً': 'Department with most errors',
 'أكثر موظف أخطاءً': 'Staff member with most errors',
 'عمليات الجرد': 'Stock counts',
 'أعمال الصيانة': 'Maintenance jobs',
 'أيام الغياب': 'Absence days',
 'مرات التأخير': 'Late occurrences',
 'التنبيهات الشفهية': 'Verbal warnings',
 'الإشادات': 'Commendations',
 'الأخطاء': 'Errors',
 'الملاحظات': 'Notes',
 'الزيارات': 'Visits',
 'الجهات الثماني': 'the eight entities',
 'ملخص ذكي': 'smart summary',
 'رسم بياني': 'chart',
 'اختر': 'Select',
 'هذه السنة': 'This year',
 'هذا الأسبوع': 'This week',
 'الشهر مقابل الماضي': 'Month vs. previous',
 'الأسبوع مقابل الماضي': 'Week vs. previous',
 'السنة مقابل الماضية': 'Year vs. previous',
 'اليوم فقط': 'Today only',
 'وارد المخازن': 'Warehouse stock-in',
 'قاعدة العاملين': 'Employees database',
 'صورة': 'Photo',
 'صورة الموظف': 'Employee photo',
 'صورة البطاقة': 'ID photo',
 'اختيار': 'Choose',
 'عرض': 'View',
 'إزالة': 'Remove',
 'كلمة المرور الجديدة': 'New password',
 'تأكيد كلمة المرور الجديدة': 'Confirm new password',
 'كلمة المرور الحالية': 'Current password',
 'زيارات اليوم': "Today's visits",
 'اليوم': 'Today',
 'الجديدة': 'new',
 'الاسم': 'Name',
 'للمخازن': 'for warehouses',
 'المخازن': 'warehouses',
 'الحساب': 'Account',
 'تفعيل': 'enable',
 'إيقاف': 'disable'}

_ARABIC = _re.compile(r"[\u0600-\u06FF]")
_AR_LETTER = "\u0621-\u064A"

# ------------------------------------------------------------ Arabic shaping
# Tk on X11 (Linux) does not shape or reorder Arabic: letters show up
# disconnected and the line reads left-to-right.  Windows Tk uses Uniscribe
# and already renders Arabic correctly, so this workaround is enabled only
# where it is needed - otherwise the text would be shaped twice.
_NEEDS_SHAPING = not _sys.platform.startswith("win")
try:
    import arabic_reshaper as _reshaper
    from bidi.algorithm import get_display as _bidi_display
except Exception:                                   # optional dependency
    _reshaper = None
    _bidi_display = None


def _shape(text):
    """Join Arabic letters into their contextual forms (display only)."""
    if not _NEEDS_SHAPING or _reshaper is None:
        return text
    if not isinstance(text, str) or not _ARABIC.search(text):
        return text
    try:
        return _bidi_display(_reshaper.reshape(text))
    except Exception:
        return text

# ------------------------------------------------------- إضافات: المظهر وتقارير الجهات
_DICT.update({
    '🎨 الألوان والمظهر': '🎨 Colors & Appearance',
    '🎨  الألوان والمظهر': '🎨  Colors & Appearance',
    'الألوان والمظهر   🎨': 'Colors & Appearance   🎨',
    'الألوان والمظهر': 'Colors & Appearance',
    'اختر ألوان الواجهة وشكل العرض — التغيير يُطبَّق ويُحفظ فورًا، والوضع النهاري هو الأساس للعمل المكتبي.':
        'Pick interface colors and layout style — changes apply and save instantly.',
    'ألوان الواجهة (الوضع النهاري + الليلي)': 'Interface colors (light + dark)',
    'شكل الواجهة': 'Interface style',
    'ملاحظة: كل الألوان فاتحة ومناسبة للعمل المكتبي، والوضع الليلي يبقى متاحًا من زر القمر أعلى الشاشة.':
        'Note: all palettes are light and office-friendly; dark mode stays available from the moon button.',
    'مريح': 'Comfortable',
    'مضغوط': 'Compact',
    'كلاسيكي': 'Classic',
    'أزرق مؤسسي': 'Corporate blue',
    'حكومي كلاسيكي': 'Government classic',
    'أخضر مؤسسي': 'Corporate green',
    'أزرق ملكي': 'Royal blue',
    'تركوازي بحري': 'Marine teal',
    'رمادي مكتبي': 'Office gray',
    'بيج دافئ': 'Warm beige',
    'الوضع الليلي': 'Dark mode',
    '✔ مختار': '✔ Selected',
    'تطبيق': 'Apply',
    '🖨️ طباعة تقارير': '🖨️ Print reports',
    '🖨️ تقارير حضور وانصراف موظفي الجهات': '🖨️ Entity staff attendance reports',
    'تقرير حضور وانصراف موظفي الجهات': 'Entity staff attendance report',
    'حدد المدة، واختر جهة أو موظفًا (أو اترك الكل)، ثم اطبع أو استخرج Excel.':
        'Pick the period and an entity or employee (or leave all), then print or export Excel.',
    'اترك «كل الجهات» و«كل الموظفين» لعمل تقرير شامل.': 'Leave “All entities” and “All employees” for a full report.',
    'كل الجهات': 'All entities',
    'كل الموظفين': 'All employees',
    'محتوى التقرير:': 'Report content:',
    'تفاصيل الحضور': 'Attendance details',
    'ملخص بالموظف والجهة': 'Summary by employee & entity',
    'التفاصيل + الملخص': 'Details + summary',
    'التقرير يطبع على ورق A4 منسق بلوجو خدمات مصر وخانات التوقيع، ويمكن حفظه PDF.':
        'Printed on A4 with the Khadamat Misr logo and signature fields; can be saved as PDF.',
    'لا توجد بيانات': 'No data',
    'لا توجد سجلات حضور في المدة/الاختيار المحدد.': 'No attendance records for the selected period/filter.',
    'حضور موظفي الجهات': 'Entity staff attendance',
    'قاعدة موظفي الجهات': 'Entity staff database',
    'ملخص حضور الموظفين': 'Staff attendance summary',
    'ملخص الجهات': 'Entities summary',
    'تفاصيل حضور وانصراف موظفي الجهات': 'Entity staff check-in/out details',
    'ملخص حضور كل موظف': 'Attendance summary per employee',
    'ملخص الحضور حسب الجهة': 'Attendance summary by entity',
    'أيام مسجلة': 'Recorded days',
    'أيام مكتملة': 'Complete days',
    'أيام بدون انصراف': 'Days without check-out',
    'إجمالي الساعات': 'Total hours',
    'متوسط الساعات/يوم': 'Avg hours/day',
    'عدد الموظفين المسجلين': 'Registered employees',
    'عدد السجلات': 'Records',
    'بالمدة أو بالجهة أو بالاسم': 'by period, entity or name',
    'تسجيل سريع بساعة الجهاز:': 'Quick registration with the device clock:',
    '⏱️ الوقت بيتاخد من ساعة الجهاز لحظة الضغط (والتاريخ المختار فوق)':
        '⏱️ Time is taken from the device clock on click (for the selected date)',
    '✅ حضور الآن': '✅ Check-in now',
    '🚪 انصراف الآن': '🚪 Check-out now',
    'اختر العامل': 'Pick an employee',
    'اختر اسم العامل من القائمة أولًا، ثم اضغط الحضور أو الانصراف.':
        'Pick the employee from the list first, then press check-in or check-out.',
    'تاريخ مختلف': 'Different date',
    'مسجل بالفعل': 'Already recorded',
    'مفيش حضور مسجل': 'No check-in recorded',
    'تم تسجيل الحضور': 'Check-in recorded',
    'تم تسجيل الانصراف': 'Check-out recorded',
})

# Phrase fallback: used for composed strings (f-strings with data inside).
# Longer phrases are tried first so words are not translated piecemeal.
_SKIP = {"خ", "تم", "من", "إلى", "في", "حفظ", "خطأ", "عدد", "اسم", "وارد", "جرد", "رصيد"}
_PHRASES = sorted(
    ((k, v) for k, v in _DICT.items()
     if len(k) >= 3 and k not in _SKIP and _ARABIC.search(k)),
    key=lambda kv: -len(kv[0]))


def is_en():
    return LANG == "en"


def T(text):
    """Translate a UI string to the active language."""
    if not isinstance(text, str) or not text:
        return text
    if LANG != "en":
        return _shape(text)
    hit = _DICT.get(text)
    if hit is not None:
        return hit
    if not _ARABIC.search(text):
        return text
    out = text
    for ar, en in _PHRASES:
        if ar in out:
            if " " in ar:
                out = out.replace(ar, en)
            else:
                out = _re.sub(r"(?<![" + _AR_LETTER + r"])" + _re.escape(ar) +
                              r"(?![" + _AR_LETTER + r"])", en, out)
    out = _re.sub(r"([0-9]{1,2}:[0-9]{2})\s*\u0635", r"\1 AM", out)
    out = _re.sub(r"([0-9]{1,2}:[0-9]{2})\s*\u0645", r"\1 PM", out)
    out = out.replace("\u060c", ",").replace("\u061b", ";")
    return out


def set_lang(code):
    global LANG, MONTHS
    LANG = code if code in ("ar", "en") else "ar"
    MONTHS = MONTHS_EN if LANG == "en" else MONTHS_AR


MONTHS_AR = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر']
MONTHS_EN = ['January',
 'February',
 'March',
 'April',
 'May',
 'June',
 'July',
 'August',
 'September',
 'October',
 'November',
 'December']
MONTHS = MONTHS_AR


def month_index(name):
    if name in MONTHS_AR:
        return MONTHS_AR.index(name) + 1
    if name in MONTHS_EN:
        return MONTHS_EN.index(name) + 1
    return 1


# ------------------------------------------------------------------- tk hooks
_SWAP = {"left": "right", "right": "left", "e": "w", "w": "e",
         "ne": "nw", "nw": "ne", "se": "sw", "sw": "se"}


def _side(v):
    return _SWAP.get(v, v) if LANG == "en" else v


def _anchor(v):
    return _SWAP.get(v, v) if LANG == "en" else v


def _sticky(v):
    if LANG != "en" or not isinstance(v, str):
        return v
    return v.translate(str.maketrans({"e": "w", "w": "e"}))


def _fix(kw):
    """Translate text-ish keywords and mirror alignment ones."""
    if "text" in kw and isinstance(kw["text"], str):
        kw["text"] = T(kw["text"])
    if "label" in kw and isinstance(kw["label"], str):
        kw["label"] = T(kw["label"])
    if LANG == "en":
        for k in ("anchor", "justify"):
            if k in kw and isinstance(kw[k], str):
                kw[k] = _anchor(kw[k])
        if "side" in kw and isinstance(kw["side"], str):
            kw["side"] = _side(kw["side"])
        if "sticky" in kw and isinstance(kw["sticky"], str):
            kw["sticky"] = _sticky(kw["sticky"])


_installed = False


def install():
    """Patch tkinter so every widget text is translated and mirrored."""
    global _installed
    if _installed:
        return
    _installed = True

    def wrap_init(cls):
        orig = cls.__init__

        def init(self, *a, **kw):
            if a and isinstance(a[0], dict):
                a = (dict(a[0]),) + a[1:]
                _fix(a[0])
            _fix(kw)
            orig(self, *a, **kw)
        cls.__init__ = init

    for cls in (_tk.Label, _tk.Button, _tk.Checkbutton, _tk.Radiobutton,
                _tk.Menubutton, _tk.Message, _tk.LabelFrame,
                _ttk.Label, _ttk.Button, _ttk.Checkbutton, _ttk.Radiobutton,
                _ttk.LabelFrame):
        wrap_init(cls)

    # widget.configure(text=...) / config(...)
    _orig_config = _tk.Misc.configure

    def configure(self, cnf=None, **kw):
        if isinstance(cnf, dict):
            cnf = dict(cnf)
            _fix(cnf)
        _fix(kw)
        return _orig_config(self, cnf, **kw)
    _tk.Misc.configure = configure
    _tk.Misc.config = configure

    # geometry managers: mirror side / anchor / sticky
    _orig_pack = _tk.Pack.pack
    _orig_grid = _tk.Grid.grid

    def pack(self, cnf={}, **kw):
        if isinstance(cnf, dict):
            cnf = dict(cnf)
            _fix(cnf)
        _fix(kw)
        return _orig_pack(self, cnf, **kw)

    def grid(self, cnf={}, **kw):
        if isinstance(cnf, dict):
            cnf = dict(cnf)
            _fix(cnf)
        _fix(kw)
        return _orig_grid(self, cnf, **kw)
    _tk.Pack.pack = pack
    _tk.Grid.grid = grid

    # window / dialog titles
    _orig_title = _tk.Wm.title

    def title(self, string=None):
        if string is not None:
            string = T(string)
        return _orig_title(self, string)
    _tk.Wm.title = title

    # canvas text (header clock, stat cards, section icons)
    _orig_ct = _tk.Canvas.create_text
    _orig_ic = _tk.Canvas.itemconfigure

    def create_text(self, *a, **kw):
        if "text" in kw:
            kw["text"] = T(kw["text"])
        return _orig_ct(self, *a, **kw)

    def itemconfigure(self, tag, cnf=None, **kw):
        if isinstance(cnf, dict):
            cnf = dict(cnf)
            if "text" in cnf:
                cnf["text"] = T(cnf["text"])
        if "text" in kw:
            kw["text"] = T(kw["text"])
        return _orig_ic(self, tag, cnf, **kw)
    _tk.Canvas.create_text = create_text
    _tk.Canvas.itemconfigure = itemconfigure
    _tk.Canvas.itemconfig = itemconfigure

    # tree headings / notebook tabs / menus
    _orig_heading = _ttk.Treeview.heading

    def heading(self, column, option=None, **kw):
        if isinstance(option, dict) and "text" in option:
            option = dict(option)
            option["text"] = T(option["text"])
        if "text" in kw:
            kw["text"] = T(kw["text"])
        return _orig_heading(self, column, option, **kw)
    _ttk.Treeview.heading = heading

    # tree cell values: shape for display, keep the raw text so data read back
    # (inventory save, inline cell edit, entity keys) stays clean.
    _orig_tv_insert = _ttk.Treeview.insert
    _orig_tv_item = _ttk.Treeview.item
    _orig_tv_set = _ttk.Treeview.set
    _orig_tv_delete = _ttk.Treeview.delete

    def _shape_cells(values):
        return [_shape(v) if isinstance(v, str) else v for v in values]

    def _raw_store(tree):
        store = getattr(tree, "_khm_raw_cells", None)
        if store is None:
            store = {}
            tree._khm_raw_cells = store
        return store

    def _col_index(tree, column):
        if isinstance(column, str) and column.startswith("#") and column[1:].isdigit():
            return int(column[1:]) - 1
        cols = tree["columns"]
        if isinstance(cols, str):
            cols = [cols]
        try:
            return list(cols).index(column)
        except (ValueError, TypeError):
            return None

    def tv_insert(self, parent, index, iid=None, **kw):
        vals = kw.get("values")
        raw = list(vals) if vals is not None else None
        if vals is not None:
            kw["values"] = _shape_cells(vals)
        new_iid = _orig_tv_insert(self, parent, index, iid, **kw)
        if raw is not None:
            _raw_store(self)[new_iid] = raw
        return new_iid

    def tv_item(self, item, option=None, **kw):
        store = _raw_store(self)
        if kw.get("values") is not None:
            raw = list(kw["values"])
            store[item] = raw
            kw["values"] = _shape_cells(raw)
        res = _orig_tv_item(self, item, option, **kw)
        if option == "values" and item in store:
            return tuple(store[item])
        if option is None and isinstance(res, dict) and item in store:
            res = dict(res)
            res["values"] = tuple(store[item])
        return res

    def tv_set(self, item, column=None, value=None):
        store = _raw_store(self)
        if value is None:
            idx = _col_index(self, column)
            res = _orig_tv_set(self, item, column)
            if idx is not None and item in store and 0 <= idx < len(store[item]):
                return store[item][idx]
            return res
        idx = _col_index(self, column)
        if idx is not None:
            row = store.setdefault(item, [])
            while len(row) <= idx:
                row.append("")
            row[idx] = value
        return _orig_tv_set(self, item, column, _shape(value) if isinstance(value, str) else value)

    def tv_delete(self, *items):
        store = _raw_store(self)
        for it in items:
            store.pop(it, None)
        return _orig_tv_delete(self, *items)

    _ttk.Treeview.insert = tv_insert
    _ttk.Treeview.item = tv_item
    _ttk.Treeview.set = tv_set
    _ttk.Treeview.delete = tv_delete

    _orig_nb_add = _ttk.Notebook.add

    def nb_add(self, child, **kw):
        if "text" in kw:
            kw["text"] = T(kw["text"])
        return _orig_nb_add(self, child, **kw)
    _ttk.Notebook.add = nb_add

    for name in ("add_command", "add_cascade", "add_checkbutton", "add_radiobutton"):
        orig = getattr(_tk.Menu, name)

        def make(orig=orig):
            def add(self, cnf={}, **kw):
                if isinstance(cnf, dict):
                    cnf = dict(cnf)
                    if "label" in cnf:
                        cnf["label"] = T(cnf["label"])
                if "label" in kw:
                    kw["label"] = T(kw["label"])
                return orig(self, cnf, **kw)
            return add
        setattr(_tk.Menu, name, make())

    # message boxes
    for name in ("showinfo", "showwarning", "showerror", "askyesno",
                 "askokcancel", "askquestion"):
        orig = getattr(_mb, name)

        def make(orig=orig):
            def box(title=None, message=None, **opts):
                return orig(T(title), T(message), **opts)
            return box
        setattr(_mb, name, make())

    # file dialogs
    for name in ("askopenfilename", "asksaveasfilename", "askdirectory"):
        orig = getattr(_fd, name)

        def make(orig=orig):
            def dlg(**kw):
                if "title" in kw:
                    kw["title"] = T(kw["title"])
                return orig(**kw)
            return dlg
        setattr(_fd, name, make())


def sync_months():
    """Keep MONTHS pointing at the list for the active language."""
    global MONTHS
    MONTHS = MONTHS_EN if LANG == "en" else MONTHS_AR
