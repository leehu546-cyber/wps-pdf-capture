@echo off
cd /d "%~dp0"
if exist "%~dp0PDF截图助手.exe" (
    start "" "%~dp0PDF截图助手.exe"
    exit /b 0
)
if exist "%~dp0dist\PDF截图助手\PDF截图助手.exe" (
    start "" "%~dp0dist\PDF截图助手\PDF截图助手.exe"
    exit /b 0
)
if exist "D:\Python313\pythonw.exe" (
    start "" "D:\Python313\pythonw.exe" "%~dp0_wps_pdf_capture_test.py"
    exit /b 0
)
echo 未找到 PDF截图助手.exe，也未找到 Python。请先运行 build_release.bat 打包。
pause
