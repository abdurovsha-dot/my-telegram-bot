@echo off
REM ===========================================================
REM   SysOne Mahalla - BUILD (PyInstaller)
REM   (c) SysOne Digital Solutions
REM ===========================================================

REM --- Admin huquqlarini so'rash (UAC) ---
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Admin huquqlari so'ralmoqda...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

REM --- Skript joylashgan papkaga o'tish ---
cd /d "%~dp0"

echo ===========================================================
echo    SysOne Mahalla - BUILD
echo ===========================================================
echo.

REM --- Python bor-yo'qligini tekshirish ---
set "PYEXE="
where py >nul 2>&1 && set "PYEXE=py"
if not defined PYEXE where python >nul 2>&1 && set "PYEXE=python"

if not defined PYEXE goto INSTALL_PYTHON
goto HAVE_PYTHON

:INSTALL_PYTHON
echo [!] Python topilmadi. Yuklab o'rnatilmoqda...
powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe' -OutFile '%TEMP%\python-setup.exe'"
if not exist "%TEMP%\python-setup.exe" (
    echo [XATO] Python yuklab bo'lmadi. Internetni tekshiring.
    pause
    exit /b 1
)
echo Python o'rnatilmoqda (jim rejimda, bir oz kuting)...
"%TEMP%\python-setup.exe" /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1
echo.
echo [OK] Python o'rnatildi.
echo     PATH yangilanishi uchun bu oynani YOPING va
echo     build.bat ni QAYTA ishga tushiring.
pause
exit /b

:HAVE_PYTHON
echo [OK] Python topildi:
%PYEXE% --version
echo.

REM --- pip yangilash ---
echo [1/3] pip yangilanmoqda...
%PYEXE% -m pip install --upgrade pip

REM --- Kutubxonalar ---
echo.
echo [2/3] Kutubxonalar o'rnatilmoqda...
%PYEXE% -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [XATO] Kutubxona o'rnatishda muammo.
    pause
    exit /b 1
)

REM --- Icon tekshiruvi ---
set "ICONOPT="
if exist sysone.ico (
    set "ICONOPT=--icon=sysone.ico"
    echo [OK] Icon topildi: sysone.ico
) else (
    echo [!] sysone.ico topilmadi - exe standart icon bilan chiqadi.
)

REM --- Build (bitta qatorda, caret yo'q) ---
echo.
echo [3/3] Build boshlandi...
%PYEXE% -m PyInstaller --noconfirm --clean --onefile --console --name "SysOne Mahalla" %ICONOPT% --version-file version_info.txt --collect-all selenium --collect-all trio --collect-all trio_websocket --collect-all outcome --hidden-import openpyxl --hidden-import openpyxl.cell._writer --hidden-import pynput.keyboard._win32 --hidden-import pynput.mouse._win32 "SysOne_Mahalla.py"

if %errorlevel% neq 0 (
    echo [XATO] Build muvaffaqiyatsiz. Yuqoridagi xabarni o'qing.
    pause
    exit /b 1
)

echo.
echo ===========================================================
echo    TAYYOR!  dist\SysOne Mahalla.exe
echo.
echo    Exe yoniga ro'yhat.xlsx ni qo'ying.
echo    Birinchi ishga tushishda internet kerak (ChromeDriver).
echo    Google Chrome o'rnatilgan bo'lsin.
echo ===========================================================
echo.
pause
