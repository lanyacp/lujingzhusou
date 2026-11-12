# merge-publish.ps1 — 在 %APPDATA%\kingsoft\wps\jsaddons\publish.xml 中合并/移除本加载项的注册条目
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File merge-publish.ps1 -Mode install|uninstall
param(
    [Parameter(Position = 0)][ValidateSet('install', 'uninstall')][string]$Mode = 'install',
    [string]$Name = 'ABPathHelper',
    [string]$Url = 'ABPathHelper_1.0.0',
    [string]$Version = '1.0.0'
)

$ErrorActionPreference = 'Stop'
$jsaddons = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
$pub = Join-Path $jsaddons 'publish.xml'
$entry = "<jsplugin name=`"$Name`" type=`"wpp`" url=`"$Url`" version=`"$Version`" enable=`"enable_dev`" install=`"null`"/>"

if (-not (Test-Path $jsaddons)) {
    New-Item -ItemType Directory -Force -Path $jsaddons | Out-Null
    Write-Host "created jsaddons dir: $jsaddons"
}

function Remove-Entry([string]$xml) {
    # 移除已有的同名条目(自闭合); 兼容属性顺序任意
    return [regex]::Replace($xml, "<jsplugin\b[^>]*\bname=`"$Name`"[^>]*/>\s*", '')
}

if ($Mode -eq 'install') {
    if (Test-Path $pub) {
        $xml = [IO.File]::ReadAllText($pub)
        $xml = Remove-Entry $xml
        if ($xml -match '</jsplugins>') {
            $xml = $xml -replace '</jsplugins>', ($entry + "`r`n</jsplugins>")
        } else {
            $xml = "<jsplugins>`r`n" + $entry + "`r`n</jsplugins>"
        }
    } else {
        $xml = "<jsplugins>`r`n" + $entry + "`r`n</jsplugins>"
    }
    [IO.File]::WriteAllText($pub, $xml, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "publish.xml updated (install): $pub"
    Write-Host $entry
} else {
    if (Test-Path $pub) {
        $xml = [IO.File]::ReadAllText($pub)
        $xml = Remove-Entry $xml
        [IO.File]::WriteAllText($pub, $xml, (New-Object System.Text.UTF8Encoding($false)))
        Write-Host "publish.xml updated (uninstall): $pub"
    } else {
        Write-Host "publish.xml not found, skip"
    }
}
