@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Устанавливаю нужные библиотеки...
python -m pip install --upgrade requests openpyxl pypdf reportlab pillow pyinstaller
if errorlevel 1 goto fail
echo Собираю BLAGO OTGRUZKI.exe ...
python -m PyInstaller --noconfirm --onefile --windowed --icon blago.ico --add-data "blago.ico;." --name "BLAGO OTGRUZKI" blago_otgruzki.py
if errorlevel 1 goto fail
copy /y "dist\BLAGO OTGRUZKI.exe" "BLAGO OTGRUZKI.exe" >nul
echo.
echo Готово: файл "BLAGO OTGRUZKI.exe" лежит в этой папке.
pause
exit /b 0
:fail
echo.
echo Не получилось. Сфотографируй это окно и пришли.
pause
exit /b 1
