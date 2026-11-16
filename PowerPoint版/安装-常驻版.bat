@echo off
chcp 65001 >nul
setlocal
set "DEST=%APPDATA%\ABPathHelper"
set "SRC=%~dp0"

echo ============================================
echo   路径助手 · PowerPoint 常驻版 安装
echo ============================================
echo.

if not exist "%SRC%PathHelper.ppam" (
    echo [失败] 找不到 PathHelper.ppam，请确认本脚本与它在同一文件夹。
    pause
    exit /b 1
)

if not exist "%DEST%" mkdir "%DEST%"
copy /y "%SRC%PathHelper.ppam" "%DEST%\PathHelper.ppam" >nul

reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\ABPathHelper" /v Path /t REG_SZ /d "%DEST%\PathHelper.ppam" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\ABPathHelper" /v Autoload /t REG_DWORD /d 1 /f >nul
rem 受信任位置: 让加载项免宏提示直接运行
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Trusted Locations\ABPathHelper" /v Path /t REG_SZ /d "%DEST%" /f >nul
if errorlevel 1 (
    echo [失败] 写注册表出错。
    pause
    exit /b 1
)

echo 安装完成！
echo.
echo 下一步：完全退出 PowerPoint（若开着）→ 重新打开 PowerPoint
echo 启动时会自动出现「路径助手」悬浮面板和工具条按钮。
echo.
echo 若提示宏被禁用：文件→选项→信任中心→信任中心设置→宏设置
echo 选「禁用所有宏，并发出通知」→ 重开 PowerPoint → 点黄条「启用内容」
echo.
pause
