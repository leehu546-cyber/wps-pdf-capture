@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "DEST=%LOCALAPPDATA%\PDF截图助手"
echo 正在安装 PDF截图助手 到:
echo   %DEST%
echo.

if not exist "%~dp0PDF截图助手.exe" (
    echo [ERROR] 请在本文件夹内运行安装（需与 PDF截图助手.exe 同级）。
    pause
    exit /b 1
)

mkdir "%DEST%" 2>nul
xcopy /E /I /Y /Q "%~dp0*" "%DEST%\" >nul
if errorlevel 1 (
    echo [ERROR] 复制文件失败
    pause
    exit /b 1
)

powershell -NoProfile -Command ^
  "$dest='%DEST%';" ^
  "$lnk=Join-Path ([Environment]::GetFolderPath('Desktop')) 'PDF截图助手.lnk';" ^
  "$s=(New-Object -COM WScript.Shell).CreateShortcut($lnk);" ^
  "$s.TargetPath=Join-Path $dest 'PDF截图助手.exe';" ^
  "$s.WorkingDirectory=$dest;" ^
  "$s.Description='PDF截图助手';" ^
  "$s.Save()"

echo.
echo [OK] 安装完成
echo [OK] 桌面已创建快捷方式「PDF截图助手」
echo [OK] 程序目录: %DEST%
echo.
set /p RUN=是否现在启动？(Y/N): 
if /i "%RUN%"=="Y" start "" "%DEST%\PDF截图助手.exe"
exit /b 0
