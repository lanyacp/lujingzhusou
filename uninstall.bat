@echo off
chcp 65001 >nul
setlocal
set "SRC=%~dp0"
set "DEST=%APPDATA%\kingsoft\wps\jsaddons\ABPathHelper_1.0.0"

echo 卸载「路径助手」加载项...
if exist "%DEST%" rmdir /s /q "%DEST%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SRC%tools\merge-publish.ps1" -Mode uninstall
echo.
echo 卸载完成，请重启 WPS 生效。
pause
