@echo off
setlocal
set "DEST=%APPDATA%\ABPathHelper"
set "SRC=%~dp0"

echo ============================================
echo   路径助手 PowerPoint 常驻版 安装
echo ============================================
echo.

if not exist "%SRC%PathHelper.ppam" goto nopam
if not exist "%DEST%" mkdir "%DEST%"
copy /y "%SRC%PathHelper.ppam" "%DEST%\PathHelper.ppam" >nul

reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\ABPathHelper" /v Path /t REG_SZ /d "%DEST%\PathHelper.ppam" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\ABPathHelper" /v Autoload /t REG_DWORD /d 1 /f >nul
rem Office 只识别 LocationN 形式的信任位置键, 用高编号避免占用系统默认位置
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Trusted Locations\Location88" /v Path /t REG_SZ /d "%DEST%" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Trusted Locations\Location88" /v Description /t REG_SZ /d "路径助手 PowerPoint 加载项" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Trusted Locations\Location88" /v AllowSubfolders /t REG_DWORD /d 1 /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Trusted Locations\Location88" /v Date /t REG_SZ /d "2026-11-16" /f >nul

echo 安装完成!
echo.
echo 下一步: 完全退出 PowerPoint 再重新打开,
echo 顶部会出现 [路径助手面板] 按钮, 点它弹出面板, 任何PPT都在.
echo.
echo 若提示宏被禁用: 文件-选项-信任中心-信任中心设置-宏设置,
echo 选 [禁用所有宏并发出通知], 重开PowerPoint后点黄条 [启用内容].
echo.
pause
exit /b 0

:nopam
echo [失败] 找不到 PathHelper.ppam, 请确认与本脚本同一文件夹.
pause
exit /b 1