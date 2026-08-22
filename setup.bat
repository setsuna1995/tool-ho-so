@echo off
chcp 65001 >nul
echo ============================================
echo  Cai dat moi truong cho cong cu tao ho so
echo ============================================

where python >nul 2>nul
if errorlevel 1 (
    echo [LOI] Khong tim thay Python tren may nay.
    echo Vui long cai Python tu https://www.python.org/downloads/
    echo ^(nho tick "Add python.exe to PATH" khi cai^) roi chay lai setup.bat nay.
    pause
    exit /b 1
)

echo Da tim thay Python, dang cai thu vien can thiet...
python -m pip install --user -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo [LOI] Cai thu vien that bai. Kiem tra ket noi mang roi thu lai.
    pause
    exit /b 1
)

echo.
python "%~dp0check_backend.py"

echo.
echo ============================================
echo  Cai dat xong. Co the chay tao_ho_so_moi.py
echo ============================================
pause
