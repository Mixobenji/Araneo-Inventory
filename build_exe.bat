@echo off
REM ============================================================
REM  Araneo Inventory v1.0.5 - Windows EXE Builder
REM ============================================================
echo [1/4] Checking Python...
python --version >nul 2>&1 || (echo ERROR: Python not found. & pause & exit /b 1)
echo [2/4] Installing dependencies...
python -m pip install openpyxl pyinstaller pillow --quiet
echo [3/4] Building executable...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "Araneo_Inventory" ^
    --icon "araneo_inventory.ico" ^
    araneo_inventory.py
echo [4/4] Done!
if exist dist\Araneo_Inventory.exe (
    copy dist\Araneo_Inventory.exe .
    echo SUCCESS: Araneo_Inventory.exe is ready.
) else (
    echo ERROR: Build failed.
)
pause
