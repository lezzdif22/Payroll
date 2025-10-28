Building an EXE - guidance

This project includes a GitHub Actions workflow and local scripts to build executables.

1) Get the Windows EXE (recommended)
- Push to the `main` branch or run the workflow manually from the Actions tab.
- Open the workflow run and download the `dynamic_payroll_exe` artifact. The artifact contains `dynamic_payroll.exe` (single-file PyInstaller build).

2) Build locally on Windows
- On a Windows machine with Python 3.11 installed, run `build_windows.bat` from the project root.
- The script installs dependencies and runs PyInstaller. Output will be in `dist\`.

3) Build on Linux
- Run `./build_linux.sh` to produce a Linux onefile binary in `dist/`.
- This will not create a Windows exe. To make a Windows exe, either build on Windows or use the GitHub Actions workflow.

Notes and tips
- Data files: PyInstaller doesn't automatically include non-Python files. Use `--add-data "src;dest"` for each file or folder. The included scripts pass examples for `templates/` and `PLV LOGO.png`.
- Hidden imports: If PyInstaller misses imports, add them with `--hidden-import modulename`.
- GUI apps: If any GUI frameworks are used, ensure you set proper hooks or include required DLLs.
- Signing: For distribution outside internal use, consider code signing on Windows.

Creating an installer (Windows)

This repo includes an Inno Setup script at `installer/installer.iss`. The GitHub Actions workflow installs Inno Setup (via Chocolatey) and runs `iscc installer\installer.iss`. The workflow will attempt to upload the resulting installer as an artifact named `dynamic_payroll_installer`.

Locally on Windows:
- Install Inno Setup (https://jrsoftware.org/isinfo.php) or use Chocolatey: `choco install innosetup`.
- After running `build_windows.bat`, if `iscc` is available the script will run the Inno Setup compiler. The installer executable will be placed into `installer_output/` (or check Inno Setup's Output folder if different).

Note: The Inno script assumes `dist\dynamic_payroll.exe` exists (PyInstaller onefile). Adjust `installer.iss` if you use onefolder builds or additional files.

Troubleshooting
- If the exe crashes on startup, test the app on the same Windows machine using `python main.py` to reproduce errors before packaging.
- Collect the `--debug` PyInstaller logs (remove `--noconfirm --clean --onefile` flags appropriately) to diagnose missing files.
