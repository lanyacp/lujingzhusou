@echo off
setlocal
set "DEST=%APPDATA%\ABPathHelper"
set "SRC=%~dp0"

echo ============================================
echo   路径助手 PowerPoint 常驻版 安装
echo ============================================
echo.

if not exist "%SRC%PathHelper.ppam" goto nopam

rem PowerPoint 开着时加载项文件被占用, 覆盖会静默失败 —— 先提示退出
tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if not errorlevel 1 goto pp_running

if not exist "%DEST%" mkdir "%DEST%"
copy /y "%SRC%PathHelper.ppam" "%DEST%\PathHelper.ppam" >nul
if errorlevel 1 goto nocopy

reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\ABPathHelper" /v Path /t REG_SZ /d "%DEST%\PathHelper.ppam" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\ABPathHelper" /v Autoload /t REG_DWORD /d 1 /f >nul
rem 受信任位置: Office 实际读取的路径带 TrustCenter 一级(少了它等于没设置)
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\TrustCenter\Trusted Locations\Location88" /v Path /t REG_SZ /d "%DEST%" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\TrustCenter\Trusted Locations\Location88" /v Description /t REG_SZ /d "路径助手 PowerPoint 加载项" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\TrustCenter\Trusted Locations\Location88" /v AllowSubfolders /t REG_DWORD /d 1 /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\TrustCenter\Trusted Locations\Location88" /v Date /t REG_SZ /d "2026-11-16" /f >nul
rem 清掉旧版脚本写在错误路径下的键
reg delete "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Trusted Locations\Location88" /f >nul 2>&1

echo 安装完成!
echo.
echo 下一步: 打开 PowerPoint, 顶部(功能区下方)会出现 [路径助手面板] 按钮,
echo         点它弹出面板, 对任何 PPT 都有效.
echo.
echo 若提示宏被禁用: 文件-选项-信任中心-信任中心设置-宏设置,
echo 选 [禁用所有宏并发出通知], 重开PowerPoint后点黄条 [启用内容].
echo.
pause
exit /b 0

:pp_running
echo [未安装] 检测到 PowerPoint 正在运行.
echo        请先完全退出 PowerPoint(保存好你的课件), 再双击本脚本.
echo.
pause
exit /b 1

:nocopy
echo [失败] 无法写入 "%DEST%\PathHelper.ppam", 请确认 PowerPoint 已完全退出后重试.
pause
exit /b 1

:nopam
echo [失败] 找不到 PathHelper.ppam, 请确认与本脚本同一文件夹.
pause
exit /b 1