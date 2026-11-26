@echo off
setlocal
echo 卸载 路径助手 PowerPoint 常驻版...
echo.

rem PowerPoint 开着时加载项文件被占用, 删不掉 —— 先提示退出
tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if not errorlevel 1 goto pp_running

reg delete "HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\ABPathHelper" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\PowerPoint\TrustCenter\Trusted Locations\Location88" /f >nul 2>&1
rem 旧版脚本写在错误路径下的键, 一并清理
reg delete "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Trusted Locations\Location88" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Trusted Locations\ABPathHelper" /f >nul 2>&1
if exist "%APPDATA%\ABPathHelper" rmdir /s /q "%APPDATA%\ABPathHelper"
if exist "%APPDATA%\ABPathHelper" goto nodel
echo 卸载完成, 重新打开 PowerPoint 生效.
echo.
pause
exit /b 0

:pp_running
echo [未卸载] 检测到 PowerPoint 正在运行.
echo        请先完全退出 PowerPoint, 再双击本脚本.
echo.
pause
exit /b 1

:nodel
echo 注册信息已清除, 但文件夹 "%APPDATA%\ABPathHelper" 没删干净, 手动删除即可.
pause
exit /b 0