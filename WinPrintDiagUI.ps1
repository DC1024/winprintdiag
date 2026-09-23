# WinPrintDiag UI - graphical front end for WinPrintDiag.ps1
param(
    [switch]$SelfTest,
    [switch]$SelfTestDeep
)

$ErrorActionPreference = 'Continue'
$script:Version = '1.0'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Resolve-ScriptRootPath {
    $defPath = [string]$MyInvocation.MyCommand.Definition
    $cmdPath = [string]$MyInvocation.MyCommand.Path
    $root = ''
    if (-not [string]::IsNullOrWhiteSpace($cmdPath)) { $root = Split-Path -Parent $cmdPath }
    if ([string]::IsNullOrWhiteSpace($root) -and $PSScriptRoot) { $root = $PSScriptRoot }
    if ([string]::IsNullOrWhiteSpace($root) -and $defPath.Length -lt 1024 -and $defPath -notmatch "`n") {
        try { $root = Split-Path -Parent $defPath } catch { $root = '' }
    }
    if ([string]::IsNullOrWhiteSpace($root)) { $root = [System.IO.Directory]::GetCurrentDirectory() }
    return $root
}

$RootDir = Resolve-ScriptRootPath
$CoreScript = Join-Path $RootDir 'WinPrintDiag.ps1'
$ReportDir = Join-Path $RootDir 'reports'
$script:LastReport = $null

$id = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($id)
$script:IsAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$form = New-Object System.Windows.Forms.Form
$form.Text = 'WinPrintDiag - 打印子系统体检'
$form.Size = New-Object System.Drawing.Size(940, 660)
$form.MinimumSize = New-Object System.Drawing.Size(760, 520)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$form.BackColor = [System.Drawing.Color]::White

$pnlTop = New-Object System.Windows.Forms.Panel
$pnlTop.Dock = 'Top'
$pnlTop.Height = 58
$pnlTop.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
$form.Controls.Add($pnlTop)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'WinPrintDiag'
$lblTitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 13, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(24, 80, 140)
$lblTitle.Location = New-Object System.Drawing.Point(14, 8)
$lblTitle.AutoSize = $true
$pnlTop.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = '打印子系统体检工具 - v' + $script:Version
$lblSub.ForeColor = [System.Drawing.Color]::FromArgb(110, 118, 130)
$lblSub.Location = New-Object System.Drawing.Point(16, 34)
$lblSub.AutoSize = $true
$pnlTop.Controls.Add($lblSub)

$lblElev = New-Object System.Windows.Forms.Label
if ($script:IsAdmin) {
    $lblElev.Text = '运行级别：管理员'
    $lblElev.ForeColor = [System.Drawing.Color]::FromArgb(160, 40, 40)
}
else {
    $lblElev.Text = '运行级别：标准用户'
    $lblElev.ForeColor = [System.Drawing.Color]::FromArgb(90, 100, 112)
}
$lblElev.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9, [System.Drawing.FontStyle]::Bold)
$lblElev.AutoSize = $true
$lblElev.Location = New-Object System.Drawing.Point(0, 22)
$pnlTop.Controls.Add($lblElev)

$pnlTool = New-Object System.Windows.Forms.Panel
$pnlTool.Dock = 'Top'
$pnlTool.Height = 50
$pnlTool.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($pnlTool)

function New-ToolButton {
    param([string]$Text, [int]$X, [int]$W, [switch]$Primary)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, 11)
    $b.Size = New-Object System.Drawing.Size($W, 30)
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderSize = 1
    if ($Primary) {
        $b.BackColor = [System.Drawing.Color]::FromArgb(30, 95, 160)
        $b.ForeColor = [System.Drawing.Color]::White
    }
    else {
        $b.BackColor = [System.Drawing.Color]::FromArgb(250, 251, 253)
        $b.ForeColor = [System.Drawing.Color]::FromArgb(40, 46, 56)
    }
    $b.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    return $b
}

$btnScan = New-ToolButton -Text '开始体检' -X 14 -W 110 -Primary
$btnFix = New-ToolButton -Text '修复（需管理员）' -X 132 -W 140
$btnClear = New-ToolButton -Text '清理队列' -X 280 -W 110
$btnOpen = New-ToolButton -Text '打开报告目录' -X 398 -W 140
$btnSave = New-ToolButton -Text '另存报告' -X 546 -W 110
$pnlTool.Controls.Add($btnScan)
$pnlTool.Controls.Add($btnFix)
$pnlTool.Controls.Add($btnClear)
$pnlTool.Controls.Add($btnOpen)
$pnlTool.Controls.Add($btnSave)

$pnlSummary = New-Object System.Windows.Forms.Panel
$pnlSummary.Dock = 'Top'
$pnlSummary.Height = 38
$pnlSummary.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$form.Controls.Add($pnlSummary)

$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Text = '点击“开始体检”执行一次只读诊断（约 5 秒）。'
$lblSummary.Location = New-Object System.Drawing.Point(16, 10)
$lblSummary.AutoSize = $true
$lblSummary.ForeColor = [System.Drawing.Color]::FromArgb(70, 76, 86)
$pnlSummary.Controls.Add($lblSummary)

# 报告窗格需要“等宽 + 自带中文字形”的字体，否则中文会显示成方框。
# 注意：SimSun / NSimSun 在 GDI+ 的字体族枚举里是以本地化名（宋体 / 新宋体）出现的，
# 所以这里按名创建字体、再核对解析结果，而不是去枚举 FontFamily.Families。
# 新宋体 / 宋体的 ASCII 是半角定宽、汉字整宽，正好与报告表格的对齐假设一致。
$fontCandidates = @(
    @{ Want = 'NSimSun'; Names = @('NSimSun', '新宋体') },
    @{ Want = 'SimSun'; Names = @('SimSun', '宋体') },
    @{ Want = 'Consolas'; Names = @('Consolas') },
    @{ Want = 'Microsoft YaHei UI'; Names = @('Microsoft YaHei UI', '微软雅黑 UI') }
)
$reportFontName = ''
foreach ($cand in $fontCandidates) {
    try {
        $probe = New-Object System.Drawing.Font($cand.Want, 10)
        if ($cand.Names -contains $probe.Name) {
            $reportFontName = $probe.Name
            $probe.Dispose()
            break
        }
        $probe.Dispose()
    }
    catch {
    }
}
if ([string]::IsNullOrWhiteSpace($reportFontName)) { $reportFontName = 'Microsoft Sans Serif' }

$txtReport = New-Object System.Windows.Forms.TextBox
$txtReport.Multiline = $true
$txtReport.ReadOnly = $true
$txtReport.ScrollBars = 'Both'
$txtReport.WordWrap = $false
$txtReport.Dock = 'Fill'
$txtReport.BorderStyle = 'FixedSingle'
$txtReport.BackColor = [System.Drawing.Color]::White
$txtReport.ForeColor = [System.Drawing.Color]::FromArgb(32, 34, 38)
$txtReport.Font = New-Object System.Drawing.Font($reportFontName, 10)
$txtReport.Text = '尚无报告。点击“开始体检”开始。'
$form.Controls.Add($txtReport)

$pnlStatus = New-Object System.Windows.Forms.Panel
$pnlStatus.Dock = 'Bottom'
$pnlStatus.Height = 30
$pnlStatus.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
$form.Controls.Add($pnlStatus)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = '就绪'
$lblStatus.Location = New-Object System.Drawing.Point(14, 8)
$lblStatus.AutoSize = $true
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(100, 108, 120)
$pnlStatus.Controls.Add($lblStatus)


function Set-Busy {
    param([bool]$Busy)
    $btnScan.Enabled = -not $Busy
    $btnFix.Enabled = -not $Busy
    $btnClear.Enabled = -not $Busy
    $btnSave.Enabled = -not $Busy
}

function Set-Summary {
    param([string]$Text, [int]$Mode)
    $lblSummary.Text = $Text
    if ($Mode -eq 0) { $lblSummary.ForeColor = [System.Drawing.Color]::FromArgb(30, 130, 70) }
    if ($Mode -eq 1) { $lblSummary.ForeColor = [System.Drawing.Color]::FromArgb(180, 110, 20) }
    if ($Mode -eq 2) { $lblSummary.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 40) }
}

function Start-CoreJob {
    param([string[]]$CoreArgs)
    if (-not (Test-Path -LiteralPath $CoreScript)) {
        [System.Windows.Forms.MessageBox]::Show(('未找到核心脚本：' + [Environment]::NewLine + $CoreScript), 'WinPrintDiag', 'OK', 'Error') | Out-Null
        return $false
    }
    if (-not (Test-Path -LiteralPath $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null }
    Remove-Item (Join-Path $ReportDir 'WinPrintDiag_*.txt') -Force -ErrorAction SilentlyContinue
    $argString = '-OutDir "' + $ReportDir + '"'
    foreach ($a in $CoreArgs) {
        if (-not [string]::IsNullOrWhiteSpace($a)) { $argString = $argString + ' ' + $a }
    }
    $cmdLine = '& "' + $CoreScript + '" ' + $argString
    Set-Busy $true
    $form.Cursor = 'WaitCursor'
    [System.Windows.Forms.Application]::DoEvents()
    try {
        Invoke-Expression $cmdLine | Out-Null
        Update-ReportView
    }
    catch {
        Set-Busy $false
        Set-Summary ('诊断失败：' + $_.Exception.Message) 2
        $lblStatus.Text = '失败'
        $txtReport.Text = '诊断失败。' + [Environment]::NewLine + $_.Exception.Message
    }
    $form.Cursor = 'Default'
    return $true
}

function Update-ReportView {
    $found = Get-ChildItem $ReportDir -Filter 'WinPrintDiag_*.txt' -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
    Set-Busy $false
    if ($null -eq $found) {
        Set-Summary '体检结束，但未生成报告文件。' 2
        $lblStatus.Text = '无报告'
        return
    }
    $script:LastReport = $found.FullName
    $txtReport.Text = [System.IO.File]::ReadAllText($found.FullName, [System.Text.Encoding]::UTF8)
    $txtReport.SelectionStart = 0
    $txtReport.ScrollToCaret()
    $content = Get-Content $found.FullName
    $issueLines = @($content | Where-Object { $_ -match '^\s+\[(严重|警告|提示)\]' })
    if (@($content | Where-Object { $_ -match '^no obvious issues|未发现明显异常' }).Count -gt 0) {
        Set-Summary ('未发现问题。报告：' + $found.Name) 0
    }
    else {
        $critc = @($issueLines | Where-Object { $_ -match '严重' }).Count
        $warnc = @($issueLines | Where-Object { $_ -match '警告' }).Count
        $infoc = @($issueLines | Where-Object { $_ -match '提示' }).Count
        $mode = 2
        if ($critc -eq 0 -and $warnc -eq 0) { $mode = 1 }
        Set-Summary ('发现 ' + $issueLines.Count + ' 项：严重 ' + $critc + '、警告 ' + $warnc + '、提示 ' + $infoc + '。详见下方报告。') $mode
    }
    $lblStatus.Text = ('已完成 - ' + (Get-Date).ToString('HH:mm:ss'))
}

function Invoke-Scan {
    param([string[]]$ExtraArgs)
    $lblStatus.Text = '正在体检…'
    Set-Summary '正在执行诊断（约 5 秒）…' 1
    $txtReport.Text = '正在诊断，请稍候…'
    $argList = New-Object System.Collections.ArrayList
    foreach ($a in $ExtraArgs) { if ($a) { [void]$argList.Add($a) } }
    $started = Start-CoreJob -CoreArgs $argList.ToArray()
    if (-not $started) {
        $lblStatus.Text = '未能启动'
        Set-Summary '体检未能启动。' 2
    }
}

$btnScan.Add_Click({ Invoke-Scan -ExtraArgs @() })

function Invoke-Elevated {
    param([string]$SwitchName)
    if ($script:IsAdmin) {
        if ($SwitchName -eq '-Repair') {
            $lblStatus.Text = '正在修复…'
            Set-Summary '正在执行修复（从组件存储还原被替换的文件，旧文件会先隔离备份）…' 1
        }
        else {
            $lblStatus.Text = '正在清理队列…'
            Set-Summary '正在清理打印队列（先备份到带时间戳的目录，不删除任何文件）…' 1
        }
        $started = Start-CoreJob -CoreArgs @($SwitchName)
        if (-not $started) { $lblStatus.Text = '未能启动' }
        return
    }
    $msg = '该操作需要管理员权限。' + [Environment]::NewLine + [Environment]::NewLine + '请右键以管理员身份重新运行本工具，然后再点击此按钮。'
    [System.Windows.Forms.MessageBox]::Show($msg, 'WinPrintDiag', 'OK', 'Warning') | Out-Null
}

$btnFix.Add_Click({ Invoke-Elevated -SwitchName '-Repair' })
$btnClear.Add_Click({ Invoke-Elevated -SwitchName '-ClearQueue' })

$btnOpen.Add_Click({
    if (-not (Test-Path -LiteralPath $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null }
    Start-Process explorer -ArgumentList $ReportDir -ErrorAction SilentlyContinue
})

$btnSave.Add_Click({
    if ([string]::IsNullOrWhiteSpace($script:LastReport)) {
        [System.Windows.Forms.MessageBox]::Show('请先执行一次体检。', 'WinPrintDiag', 'OK', 'Information') | Out-Null
        return
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = '文本报告|*.txt|所有文件|*.*'
    $dlg.FileName = (Split-Path $script:LastReport -Leaf)
    if ($dlg.ShowDialog() -eq 'OK') {
        Copy-Item -LiteralPath $script:LastReport -Destination $dlg.FileName -Force
        $lblStatus.Text = ('已保存到 ' + $dlg.FileName)
    }
})

$form.Add_Resize({
    $w = $form.ClientSize.Width
    $lblElev.Location = New-Object System.Drawing.Point(($w - $lblElev.Width - 16), 22)
})
$form.Add_Shown({
    $w = $form.ClientSize.Width
    $lblElev.Location = New-Object System.Drawing.Point(($w - $lblElev.Width - 16), 22)
})

if ($SelfTestDeep) {
    $logPath = Join-Path $RootDir 'ui_selftest_deep.txt'
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('UI deep self-test ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
    [void]$lines.Add('core script: ' + $CoreScript + ' exists=' + (Test-Path -LiteralPath $CoreScript))
    Invoke-Scan -ExtraArgs @()
    $deadline = (Get-Date).AddSeconds(120)
    while ((Get-Date) -lt $deadline) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 300
        if (-not [string]::IsNullOrWhiteSpace($script:LastReport)) { break }
    }
    [void]$lines.Add('scan completed : ' + (-not [string]::IsNullOrWhiteSpace($script:LastReport)))
    [void]$lines.Add('report file    : ' + $script:LastReport)
    [void]$lines.Add('summary label  : ' + $lblSummary.Text)
    [void]$lines.Add('status label   : ' + $lblStatus.Text)
    [void]$lines.Add('scan btn state : enabled=' + $btnScan.Enabled)
    [void]$lines.Add('report font    : ' + $reportFontName + ' -> ' + $txtReport.Font.Name)
    $tl = $txtReport.Text
    [void]$lines.Add('textbox chars  : ' + $tl.Length)
    [void]$lines.Add('textbox head   : ' + (($tl -split "`n")[0]))
    [System.IO.File]::WriteAllLines($logPath, $lines, (New-Object System.Text.UTF8Encoding($true)))
    Write-Output ('deep self-test written: ' + $logPath)
    exit 0
}

if ($SelfTest) {
    $logPath = Join-Path $RootDir 'ui_selftest.txt'
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('UI self-test ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
    [void]$lines.Add('core script  : ' + $CoreScript + ' exists=' + (Test-Path -LiteralPath $CoreScript))
    [void]$lines.Add('report dir   : ' + $ReportDir)
    [void]$lines.Add('admin        : ' + $script:IsAdmin)
    [void]$lines.Add('controls     : ' + $form.Controls.Count + ' top-level, toolbar buttons ' + $pnlTool.Controls.Count)
    $names = @()
    foreach ($ctl in $pnlTool.Controls) { $names += $ctl.Text }
    [void]$lines.Add('buttons      : ' + ($names -join ' | '))
    [void]$lines.Add('textbox font : ' + $txtReport.Font.Name + ' ' + $txtReport.Font.Size)
    [System.IO.File]::WriteAllLines($logPath, $lines, (New-Object System.Text.UTF8Encoding($true)))
    Write-Output ('self-test written: ' + $logPath)
    exit 0
}

[void]$form.ShowDialog()
