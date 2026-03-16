@echo off
chcp 65001 > /dev/null

set PYTHON=C:\Users\uscp9a\AppData\Local\Python\bin\python.exe

echo ============================================
echo  Construyendo Distribucion_OBS.exe
echo ============================================
echo.

"%PYTHON%" -m PyInstaller ^
    --onefile ^
    --console ^
    --name Distribucion_OBS ^
    --noconfirm ^
    --collect-data country_converter ^
    --collect-data openpyxl ^
    --collect-all pulp ^
    --hidden-import=openpyxl ^
    --hidden-import=xlsxwriter ^
    --hidden-import=pandas ^
    --hidden-import=numpy ^
    launcher.py

echo.
if %ERRORLEVEL% EQU 0 (
    echo [OK] Ejecutable generado en: dist\Distribucion_OBS.exe
) else (
    echo [ERROR] La compilacion fallo. Revisa los mensajes anteriores.
)

echo.
pause
