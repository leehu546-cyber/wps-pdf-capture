@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "PY=D:\Python313\python.exe"
if not exist "%PY%" (
    echo [ERROR] 未找到 Python: %PY%
    echo 请修改 build_release.bat 里的 PY 变量。
    pause
    exit /b 1
)

set "VENV=%~dp0.venv_build"
set "VENV_PY=%VENV%\Scripts\python.exe"

echo === 1/5 准备干净虚拟环境 ===
if not exist "%VENV_PY%" (
    "%PY%" -m venv "%VENV%"
    if errorlevel 1 (
        echo [ERROR] 创建虚拟环境失败
        pause
        exit /b 1
    )
)

echo === 2/5 安装打包依赖 ===
"%VENV_PY%" -m pip install -q --upgrade pip
"%VENV_PY%" -m pip install -q pyinstaller pymupdf pywin32 pillow

echo === 3/5 PyInstaller 打包 ===
"%VENV_PY%" -m PyInstaller --noconfirm --clean "PDF截图助手.spec"
if errorlevel 1 (
    echo [ERROR] PyInstaller 打包失败
    pause
    exit /b 1
)

if not exist "dist\PDF截图助手\PDF截图助手.exe" (
    echo [ERROR] 未生成 dist\PDF截图助手\PDF截图助手.exe
    pause
    exit /b 1
)

echo === 4/5 生成便携版 zip ===
if not exist "release" mkdir "release"
copy /Y "安装.bat" "dist\PDF截图助手\安装.bat" >nul
powershell -NoProfile -Command "Compress-Archive -Path 'dist\PDF截图助手\*' -DestinationPath 'release\PDF截图助手_便携版.zip' -Force"
if errorlevel 1 (
    echo [ERROR] 便携版 zip 生成失败
    pause
    exit /b 1
)

echo === 5/5 Inno Setup 安装包 ===
set "ISCC="
for %%I in (
    "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
    "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
    "%ProgramFiles%\Inno Setup 6\ISCC.exe"
) do (
    if exist %%I set "ISCC=%%~I"
)

if defined ISCC (
    "%ISCC%" "installer.iss"
    if errorlevel 1 (
        echo [ERROR] Inno Setup 编译失败
        pause
        exit /b 1
    )
    echo [OK] 安装包: release\PDF截图助手_setup.exe
) else (
    echo [WARN] 未安装 Inno Setup 6，跳过安装包。
    echo       下载: https://jrsoftware.org/isinfo.php
    echo       安装后重新运行本脚本即可生成 setup.exe
)

echo.
echo [OK] 便携版: release\PDF截图助手_便携版.zip
echo [OK] 程序目录: dist\PDF截图助手\
for %%A in ("release\PDF截图助手_便携版.zip") do echo [OK] zip 大小: %%~zA 字节
echo.
pause
