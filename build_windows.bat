@echo off
setlocal

where py >nul 2>&1
if errorlevel 1 (
  echo Python is niet gevonden. Gebruik GitHub Actions of installeer Python 3.12.
  pause
  exit /b 1
)

py -m pip install --upgrade pip
if errorlevel 1 goto :error
py -m pip install -r requirements.txt
if errorlevel 1 goto :error
py -m py_compile app.py
if errorlevel 1 goto :error
py -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name FashionImageRenamer_v5 ^
  --icon "app_icon.ico" ^
  --add-data "fotocategorieen.json;." --add-data "matchconfig.json;." ^
  app.py
if errorlevel 1 goto :error

if not exist "dist\FashionImageRenamer_v5.exe" goto :error

echo.
echo Klaar: dist\FashionImageRenamer_v5.exe
pause
exit /b 0

:error
echo.
echo De build is mislukt. Bekijk de foutmelding hierboven.
pause
exit /b 1
