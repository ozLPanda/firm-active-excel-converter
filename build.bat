@echo off
setlocal

cd /d "%~dp0"

python -m pip install -r requirements.txt
if errorlevel 1 goto :error

python -m pip install pyinstaller
if errorlevel 1 goto :error

python -m PyInstaller --clean --onefile --windowed --name satu-converter ^
  --add-data "group_aliases.json;." ^
  --add-data "satu_schema.json;." ^
  gui.py
if errorlevel 1 goto :error

echo.
echo Build completed:
echo %~dp0dist\satu-converter.exe
echo.
pause

endlocal
exit /b 0

:error
set "EXIT_CODE=%errorlevel%"
echo.
echo Build failed with error code %EXIT_CODE%.
echo Check the messages above for details.
echo.
pause
endlocal
exit /b %EXIT_CODE%
