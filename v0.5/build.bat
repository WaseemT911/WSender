@echo off
REM ==================================================================
REM  WhatsApp Messages Sender - build portable Windows executable
REM
REM  Run this on a Windows PC with:
REM    - Python 3.8.x installed (LAST version supporting Windows 7;
REM      use the 32-bit installer if your target PCs are 32-bit)
REM    - pip install -r requirements.txt
REM    - pip install pyinstaller
REM
REM  Result: dist\WhatsAppSender.exe  (runs on Windows 7 and 10,
REM  no Python or libraries needed on the target PC)
REM ==================================================================
cd /d "%~dp0"

pyinstaller --noconfirm --onefile --name "WhatsAppSender" --collect-all customtkinter main.py
if errorlevel 1 (
    echo.
    echo Build FAILED. Make sure Python 3.8 and PyInstaller are installed:
    echo     pip install pyinstaller
    echo     pip install -r requirements.txt
    pause
    exit /b 1
)

REM Put the runtime files next to the exe (config/data/message/image).
REM If missing, the program also creates them automatically on first run.
if exist config.ini copy /Y config.ini dist\ >nul
if exist message.txt copy /Y message.txt dist\ >nul
if exist data.xlsx copy /Y data.xlsx dist\ >nul
if exist images xcopy /E /I /Y images dist\images >nul

echo.
echo ============================================================
echo  Build OK:  dist\WSender.exe
REM
echo  To install on a PC: copy these onto it (any folder):
echo      WSender.exe
echo      config.ini  message.txt  data.xlsx  images\
echo ============================================================
pause
