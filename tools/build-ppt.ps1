# build-ppt.ps1 - Build the PowerPoint .pptm tool file via COM + VBIDE injection
# ASCII only; all Chinese strings are read from UTF-8 side files.
# Output: PowerPoint Version\PathHelper.pptm + self-test result printed.

$ErrorActionPreference = 'Stop'
$srcDir = Join-Path $PSScriptRoot 'ppt-src'
$outDir = Join-Path $PSScriptRoot '..\PowerPoint版'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$outPptm = Join-Path $outDir '路径助手.pptm'
if (Test-Path $outPptm) { Remove-Item -Force $outPptm }   # 只清旧 pptm, 保留说明文件
$outPptm = Join-Path $outDir '路径助手.pptm'
$utf8 = [System.Text.Encoding]::UTF8

# 1) read sources
$modCode = [IO.File]::ReadAllText((Join-Path $srcDir 'modPathHelper.bas'), $utf8)
$modCode = ($modCode -split '\r?\n' | Where-Object { $_ -notmatch '^Attribute\s' }) -join "`r`n"
$formCode = [IO.File]::ReadAllText((Join-Path $srcDir 'form-code.txt'), $utf8)
$ctrlLines = [IO.File]::ReadAllLines((Join-Path $srcDir 'form-controls.txt'), $utf8)

# 2) PowerPoint COM
$app = New-Object -ComObject PowerPoint.Application
try {
    $app.AutomationSecurity = 1  # msoAutomationSecurityLow (run our own SelfTest)
    try { $app.Visible = $true } catch {}
    Start-Sleep -Milliseconds 600
    $pres = $app.Presentations.Add(-1)  # with window

    # 3) inject module FIRST (macro must exist before binding RunMacro action)
    $vbp = $pres.VBProject
    $mod = $vbp.VBComponents.Add(1)   # vbext_ct_StdModule
    $mod.Name = 'modPathHelper'
    $mod.CodeModule.AddFromString($modCode)

    # 4) build UserForm
    $frm = $vbp.VBComponents.Add(3)   # vbext_ct_MSForm
    $frm.Name = 'frmPanel'
    $d = $frm.Designer
    try {
        $d.Properties.Item('Caption').Value = '路径助手'
        $d.Properties.Item('Width').Value = 258
        $d.Properties.Item('Height').Value = 336
    } catch {}
    foreach ($ln in $ctrlLines) {
        if ([string]::IsNullOrWhiteSpace($ln)) { continue }
        $p = $ln -split '\|'
        $c = $d.Controls.Add($p[1])
        $c.Name = $p[0]
        if ($p[2] -ne '') { $c.Caption = $p[2] }
        $c.Left = [int]$p[3]; $c.Top = [int]$p[4]; $c.Width = [int]$p[5]; $c.Height = [int]$p[6]
    }
    $frm.CodeModule.AddFromString($formCode)

    # 5) slide 1: cover + launch button
    $s1 = $pres.Slides.Add(1, 12)   # ppLayoutBlank
    $null = $s1.Shapes.AddTextbox(1, 120, 100, 480, 60)   # msoTextOrientationHorizontal
    $s1.Shapes.Item($s1.Shapes.Count).TextFrame.TextRange.Text = '路径助手 · PowerPoint 版'
    $s1.Shapes.Item($s1.Shapes.Count).TextFrame.TextRange.Font.Size = 36
    $s1.Shapes.Item($s1.Shapes.Count).TextFrame.TextRange.Font.Bold = $true
    $null = $s1.Shapes.AddTextbox(1, 120, 180, 480, 120)
    $s1.Shapes.Item($s1.Shapes.Count).TextFrame.TextRange.Text = "三步用法:`r1. 打开你的课件(保持本文件也开着)`r2. 选中图形 → 点下方启动按钮 → 面板里点「开始」`r3. 拖标记 → 「生成路径动画」`r`r【点击无效时】宏被禁用了: 文件→选项→信任中心→信任中心设置→宏设置→选「禁用所有宏，并发出通知」→确定 → 关闭重开本文件 → 点顶部黄条「启用内容」→ 再点按钮; 也可 Alt+F8 运行 ShowPanel"
    $s1.Shapes.Item($s1.Shapes.Count).TextFrame.TextRange.Font.Size = 14
    $btn = $s1.Shapes.AddShape(5, 240, 330, 240, 70)      # msoShapeRoundedRectangle
    $btn.Name = 'btnLaunch'
    $btn.TextFrame.TextRange.Text = '▶ 启动路径助手'
    $btn.TextFrame.TextRange.Font.Size = 22
    $btn.TextFrame.TextRange.Font.Bold = $true
    $btn.Fill.ForeColor.RGB = 16740352                    # BGR blue RGB(0,112,255)
    $btn.TextFrame.TextRange.Font.Color.RGB = 16777215    # white
    $as = $btn.ActionSettings.Item(1)   # ppMouseClick
    $as.Action = 8                      # ppActionRunMacro
    $as.Run = 'ShowPanel'

    # 6) self-test (macro runs inside this presentation)
    $resultFile = Join-Path $env:TEMP 'abp_ppt_selftest.txt'
    if (Test-Path $resultFile) { Remove-Item -Force $resultFile }
    $null = $app.GetType().InvokeMember('Run', 'InvokeMethod', $null, $app, @('SelfTest'))
    Start-Sleep -Milliseconds 500
    if (Test-Path $resultFile) {
        Write-Host ('SELFTEST: ' + (Get-Content $resultFile -Raw))
    } else {
        Write-Host 'SELFTEST: no result file (macro did not run?)'
    }

    # 7) save as .pptm (tool file) AND .ppam (resident add-in)
    $pres.SaveAs($outPptm, 25)
    Write-Host ('SAVED: ' + $outPptm)
    $outPpam = Join-Path $outDir 'PathHelper.ppam'
    if (Test-Path $outPpam) { Remove-Item -Force $outPpam }
    $pres.SaveAs($outPpam, 25)
    $pres.Saved = -1
    $pres.Close()
    # 修正内容类型: pptm -> addin (否则 PowerPoint 拒绝按加载项加载)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::Open($outPpam, 'Update')
    $e = $zip.GetEntry('[Content_Types].xml')
    $r = New-Object IO.StreamReader($e.Open()); $xml = $r.ReadToEnd(); $r.Close()
    $xml = $xml -replace 'application/vnd\.ms-powerpoint\.presentation\.macroEnabled\.main\+xml', 'application/vnd.ms-powerpoint.addin.macroEnabled.main+xml'
    $e.Delete()
    $ne = $zip.CreateEntry('[Content_Types].xml')
    $w = New-Object IO.StreamWriter($ne.Open()); $w.Write($xml); $w.Close()
    $zip.Dispose()
    Write-Host ('SAVED(addin): ' + $outPpam)
} finally {
    try { $app.Quit() } catch {}
    Start-Sleep -Milliseconds 800
}
Write-Host 'BUILD DONE'
