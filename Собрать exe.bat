@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "blago_otgruzki.py" goto nofile
if not exist "blago_ui.html" goto nofile
rem Ищем настоящий Python: сначала py (лаунчер), потом python. Заглушка из Microsoft Store не годится.
set "PY="
py -3 -c "import sys" >nul 2>&1 && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>&1 && set "PY=python"
if not defined PY goto nopy
echo Python: %PY%
echo Устанавливаю нужные библиотеки...
%PY% -m pip install --upgrade requests openpyxl pypdf reportlab pillow pywebview pyinstaller
if errorlevel 1 goto fail
rem Открытую программу закрываем: иначе Windows не даст заменить exe
taskkill /im "BLAGO OTGRUZKI.exe" /f >nul 2>&1
echo Собираю BLAGO OTGRUZKI.exe ...
%PY% -m PyInstaller --noconfirm --onefile --windowed --icon blago.ico --splash splash.png --add-data "blago_ui.html;." --name "BLAGO OTGRUZKI" blago_otgruzki.py
if errorlevel 1 goto fail
copy /y "dist\BLAGO OTGRUZKI.exe" "BLAGO OTGRUZKI.exe" >nul
if errorlevel 1 goto busy
echo.
echo Готово: файл "BLAGO OTGRUZKI.exe" лежит в этой папке.
findstr /b "APP_VERSION" blago_otgruzki.py
pause
exit /b 0
:busy
echo.
echo Новый exe собран, но старый не заменился - программа была открыта.
echo Закрой BLAGO OTGRUZKI и запусти этот файл ещё раз.
pause
exit /b 1
:nofile
echo Рядом с этим файлом нет blago_otgruzki.py или blago_ui.html.
echo Похоже, сборщик запущен прямо из архива. Сначала распакуй архив целиком:
echo правой кнопкой по zip - "Извлечь все", потом запусти "Собрать exe.bat" из распакованной папки.
pause
exit /b 1
:nopy
echo Не нашёл Python. Тот, которым ты запускаешь blago_otgruzki.py, должен отвечать на команду py.
:fail
echo.
echo Не получилось. Сфотографируй это окно и пришли.
pause
exit /b 1
