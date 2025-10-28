@echo off
REM Build Windows EXE locally using PyInstaller. Run this on a Windows machine.
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

nREM Adjust --add-data as needed for your templates and static files
python -m PyInstaller --noconfirm --clean --onefile ^
  --add-data "templates;templates" ^
  --add-data "PLV LOGO.png;." ^
  --name dynamic_payroll main.py

echo Build complete. Output in dist\
REM Optionally build an installer with Inno Setup if iscc is available
where iscc >nul 2>nul
if %errorlevel%==0 (
  echo Found Inno Setup (iscc). Building installer...
  if not exist installer_output mkdir installer_output
  iscc installer\installer.iss
  if exist installer_output\*.exe (
    echo Installer created in installer_output\
  ) else (
    echo Installer built; check the Inno Setup Output directory.
  )
) else (
  echo "Inno Setup (iscc) not found. To create a Windows installer, install Inno Setup and re-run this script."
)
pause
