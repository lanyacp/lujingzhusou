@echo off
chcp 65001 >nul
setlocal
set "SRC=%~dp0"
set "JSADDONS=%APPDATA%\kingsoft\wps\jsaddons"
set "DEST=%JSADDONS%\ABPathHelper_1.0.0"

echo ============================================
echo   路径助手 (WPS 演示精准路径动画加载项) 安装
echo ============================================
echo.

if not exist "%JSADDONS%" mkdir "%JSADDONS%"
if exist "%DEST%" rmdir /s /q "%DEST%"
mkdir "%DEST%"

copy /y "%SRC%index.html" "%DEST%\" >nul
copy /y "%SRC%main.js" "%DEST%\" >nul
copy /y "%SRC%ribbon.xml" "%DEST%\" >nul
xcopy "%SRC%js" "%DEST%\js\" /e /i /y >nul
xcopy "%SRC%ui" "%DEST%\ui\" /e /i /y >nul
xcopy "%SRC%images" "%DEST%\images\" /e /i /y >nul

if errorlevel 1 (
    echo [失败] 文件复制出错，请检查目录权限。
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%SRC%tools\merge-publish.ps1" -Mode install
if errorlevel 1 (
    echo [失败] 写入 publish.xml 出错。
    pause
    exit /b 1
)

echo.
echo 安装完成！请【完全退出 WPS】(包括托盘图标/后台进程) 后重新打开 WPS 演示，
echo 功能区会出现「路径助手」选项卡。
echo 若未出现，请查阅 README.md 的「故障排查」一节。
echo.
pause
