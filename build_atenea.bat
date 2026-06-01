@echo off
chcp 65001 > nul

set PYTHON=C:\Users\uscp9a\AppData\Local\Python\bin\python.exe

echo ============================================
echo  Construyendo Distribucion_OBS_Atenea.exe
echo ============================================
echo.

"%PYTHON%" -m PyInstaller ^
    --onefile ^
    --console ^
    --name Distribucion_OBS_Atenea ^
    --noconfirm ^
    --collect-data country_converter ^
    --collect-data openpyxl ^
    --collect-all pulp ^
    --hidden-import=openpyxl ^
    --hidden-import=xlsxwriter ^
    --hidden-import=pandas ^
    --hidden-import=numpy ^
    --hidden-import=msal ^
    --hidden-import=requests ^
    launcher_atenea.py

echo.
if %ERRORLEVEL% EQU 0 (
    echo [OK] Ejecutable generado en: dist\Distribucion_OBS_Atenea.exe
) else (
    echo [ERROR] La compilacion fallo. Revisa los mensajes anteriores.
)

echo.
pause
