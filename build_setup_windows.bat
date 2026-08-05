@echo off
setlocal

call build_windows.bat
if errorlevel 1 exit /b 1

set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (
  echo.
  echo Inno Setup 6 is niet gevonden.
  echo Installeer Inno Setup 6 of gebruik de GitHub-workflow Build Setup.exe.
  pause
  exit /b 1
)

"%ISCC%" installer.iss
if errorlevel 1 goto :error
if not exist "installer_output\FashionImageRenamer_v5_Setup.exe" goto :error

echo.
echo Klaar: installer_output\FashionImageRenamer_v5_Setup.exe
pause
exit /b 0

:error
echo.
echo De Setup-build is mislukt. Bekijk de foutmelding hierboven.
pause
exit /b 1
