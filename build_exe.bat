@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   خدمات مصر - بناء نسخة exe بدون Python
echo ============================================
echo.
where py >nul 2>nul
if errorlevel 1 (
  where python >nul 2>nul
  if errorlevel 1 (
    echo Python غير مثبت. حمله من https://www.python.org/downloads/
    echo وعلّم على "Add Python to PATH" أثناء التثبيت.
    pause
    exit /b 1
  )
  set PY=python
) else (
  set PY=py -3
)
echo [1/3] تثبيت المتطلبات...
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto fail
echo.
echo [2/3] بناء الملف التنفيذي (قد يأخذ دقيقتين)...
%PY% -m PyInstaller --noconfirm --clean KhadamatMisr.spec
if errorlevel 1 goto fail
echo.
echo [3/3] تم بنجاح.
echo الملف الجاهز: dist\KhadamatMisr.exe
echo.
echo ملاحظة: قاعدة البيانات center_management.db تُنشأ بجوار الملف التنفيذي
echo أول مرة، ومعه مجلد photos لحفظ صور العاملين والمستندات.
pause
exit /b 0

:fail
echo.
echo فشل البناء. اقرأ الرسائل بالأعلى.
pause
exit /b 1
