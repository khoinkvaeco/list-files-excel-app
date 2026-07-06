@echo off
chcp 65001 >nul
title Cong cu kiem tra thue
cd /d "%~dp0"

echo ============================================
echo   CONG CU HO TRO KIEM TRA THUE
echo ============================================
echo.

REM 1) Kiem tra da cai Python chua
where python >nul 2>nul
if errorlevel 1 (
    echo [LOI] Chua cai dat Python.
    echo Vui long tai va cai Python tai: https://www.python.org/downloads/
    echo Khi cai nho tich chon "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

REM 2) Tao moi truong ao lan dau
if not exist ".venv" (
    echo Dang chuan bi lan dau, vui long doi mot chut...
    python -m venv .venv
)

call ".venv\Scripts\activate.bat"

REM 3) Cai thu vien can thiet
echo Dang kiem tra / cai dat thu vien...
python -m pip install --quiet --upgrade pip
python -m pip install --quiet -r requirements.txt

REM 4) Mo ung dung tren trinh duyet
echo.
echo Dang mo ung dung tren trinh duyet tai http://localhost:8502
echo (De dong: dong cua so nay)
echo.
streamlit run app.py --server.port 8502

pause
