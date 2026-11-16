@echo off
setlocal
echo 卸载 路径助手 PowerPoint 常驻版...
reg delete "HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\ABPathHelper" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Trusted Locations\ABPathHelper" /f >nul 2>&1
if exist "%APPDATA%\ABPathHelper" rmdir /s /q "%APPDATA%\ABPathHelper"
echo.
echo 卸载完成, 重启 PowerPoint 生效.
pause