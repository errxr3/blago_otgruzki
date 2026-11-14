@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Ищем настоящий Python: сначала py (лаунчер), потом python. Заглушка из Microsoft Store не годится.
set "PY="
py -3 -c "import sys" >nul 2>&1 && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>&1 && set "PY=python"
if not defined PY goto nopy
echo Python: %PY%
echo Устанавливаю нужные библиотеки...
%PY% -m pip install --upgrade requests openpyxl pypdf reportlab pillow pyinstaller
if errorlevel 1 goto fail
echo Собираю BLAGO OTGRUZKI.exe ...
%PY% -m PyInstaller --noconfirm --onefile --windowed --icon blago.ico --splash splash.png --add-data "blago.ico;." --name "BLAGO OTGRUZKI" blago_otgruzki.py
if errorlevel 1 goto fail
copy /y "dist\BLAGO OTGRUZKI.exe" "BLAGO OTGRUZKI.exe" >nul
echo.
echo Готово: файл "BLAGO OTGRUZKI.exe" лежит в этой папке.
pause
exit /b 0
:nopy
echo Не нашёл Python. Тот, которым ты запускаешь blago_otgruzki.py, должен отвечать на команду py.
:fail
echo.
echo Не получилось. Сфотографируй это окно и пришли.
pause
exit /b 1
