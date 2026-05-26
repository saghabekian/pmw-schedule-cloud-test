@echo off
setlocal
pushd "%~dp0"

echo.
echo ==========================================
echo PMW Ticket + Fabrication APP v26
echo Desktop Arrows + Cloud Prep
echo ==========================================
echo.
echo This does NOT map network drives and does NOT create Windows users.
echo Best location: C:\PMW_APP
echo.

echo Installing required packages if needed...
py -m pip install -r "%CD%\requirements.txt"
if errorlevel 1 (
    python -m pip install -r "%CD%\requirements.txt"
)

echo.
echo Your PC network addresses are listed below.
echo On your iPhone, use one that starts with 192.168, 10., or 172.
echo Add :5050 after it.
echo.
ipconfig | findstr /i "IPv4"

echo.
echo Examples:
echo   This PC: http://127.0.0.1:5050
echo   iPhone:  http://YOUR-PC-IP:5050
echo.
echo Starting PMW APP v26...
echo Keep this black window open while testing.
echo.

py "%CD%\app.py"
if errorlevel 1 (
    python "%CD%\app.py"
)

popd
pause
