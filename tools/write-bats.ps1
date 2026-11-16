# 以 GBK(936) 编码重写安装/卸载 bat, 修复 cmd 中文乱码与语法损坏
$gbk = [System.Text.Encoding]::GetEncoding(936)
$dir = 'D:\projects\wps插件\PowerPoint版'

$install = @'
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
reg add "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Trusted Locations\ABPathHelper" /v Path /t REG_SZ /d "%DEST%" /f >nul

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
'@

$uninstall = @'
@echo off
setlocal
echo 卸载 路径助手 PowerPoint 常驻版...
reg delete "HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\ABPathHelper" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Trusted Locations\ABPathHelper" /f >nul 2>&1
if exist "%APPDATA%\ABPathHelper" rmdir /s /q "%APPDATA%\ABPathHelper"
echo.
echo 卸载完成, 重启 PowerPoint 生效.
pause
'@

# cmd 需要 CRLF 行尾, 统一转换
$install = ($install -replace "`r`n", "`n") -replace "`n", "`r`n"
$uninstall = ($uninstall -replace "`r`n", "`n") -replace "`n", "`r`n"
[IO.File]::WriteAllText((Join-Path $dir '安装-常驻版.bat'), $install, $gbk)
[IO.File]::WriteAllText((Join-Path $dir '卸载.bat'), $uninstall, $gbk)
Write-Host 'bats rewritten in GBK + CRLF'
