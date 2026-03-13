@echo off
chcp 65001 > /dev/null
echo ============================================
echo  Construyendo Distribucion_OBS.exe
echo ============================================
echo.

pyinstaller ^
    --onefile ^
    --console ^
    --name Distribucion_OBS ^
    --noconfirm ^
    --hidden-import=openpyxl ^
    --hidden-import=xlsxwriter ^
    --hidden-import=pandas ^
    --hidden-import=numpy ^
    launcher.py

echo.
if %ERRORLEVEL% EQU 0 (
    echo [OK] Ejecutable generado en: dist\Distribucion_OBS.exe
    echo.
    echo Recuerda copiar junto al .exe:
    echo   - Areas_Paises.xlsx
    echo   - SUDOKU.xlsx
) else (
    echo [ERROR] La compilacion fallo. Revisa los mensajes anteriores.
)

echo.
pause
