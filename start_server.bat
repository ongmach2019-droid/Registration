@echo off
pushd "%~dp0"
echo Starting Conference Registration Server (Excel Database)...
echo Database file: %CD%\registrations.xlsx
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://localhost:8000"
python server.py
if errorlevel 1 (
    echo.
    echo Server stopped with an error. Make sure Python and openpyxl are installed:
    echo   pip install -r requirements.txt
    pause
)
popd
