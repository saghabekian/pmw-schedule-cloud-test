@echo off
setlocal
pushd "%~dp0"
echo PMW Cloud Test v1 - Local Test
echo Installing packages...
py -m pip install -r requirements.txt
if errorlevel 1 (
    python -m pip install -r requirements.txt
)
echo Starting local server at http://127.0.0.1:5050
py app.py
if errorlevel 1 (
    python app.py
)
popd
pause
