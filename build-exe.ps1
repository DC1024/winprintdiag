# 重新打包 exe（把 .ps1 编译成单文件可执行程序）
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File build-exe.ps1
# 说明：任何对 WinPrintDiag.ps1 / WinPrintDiagUI.ps1 的改动，都必须重跑本脚本，
#       因为脚本内容是**嵌入在 exe 内部**的，不重打包的话 exe 还是旧逻辑。
# 注意：本文件必须保存为 UTF-8 带 BOM，否则 PowerShell 5.1 会按 ANSI 解码，中文全部变乱码。

$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($here)) { $here = Split-Path -Parent $MyInvocation.MyCommand.Definition }

# PS2EXE 需要先下载解包（PSGallery 的 Find-Module 在非交互模式下会卡 ShouldContinue，所以直下 nupkg）
$moduleRoot = Join-Path $env:TEMP 'ps2exe_extracted'
$manifest = Join-Path $moduleRoot 'ps2exe.psd1'

if (-not (Test-Path -LiteralPath $manifest)) {
    Write-Host 'PS2EXE 未就绪，正在下载...'
    $zip = Join-Path $env:TEMP 'ps2exe.zip'
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile('https://www.powershellgallery.com/api/v2/package/PS2EXE', $zip)
    Expand-Archive -LiteralPath $zip -DestinationPath $moduleRoot -Force
}

Import-Module $manifest -Force

# 控制台版：Subsystem = 3（Console），双击会开一个命令行窗口并保留输出
Invoke-ps2exe -inputFile (Join-Path $here 'WinPrintDiag.ps1') `
              -outputFile (Join-Path $here 'WinPrintDiag.exe') `
              -noConsole:$false -noVisualStyles:$true

# 图形版：Subsystem = 2（GUI），必须加 -STA，否则 WinForms 在部分场景下会异常
Invoke-ps2exe -inputFile (Join-Path $here 'WinPrintDiagUI.ps1') `
              -outputFile (Join-Path $here 'WinPrintDiagUI.exe') `
              -noConsole -STA

# 离线校验产物：读 PE 头确认子系统号，避免打错模式（把 GUI 打成 Console 或反之）
foreach ($f in @('WinPrintDiag.exe', 'WinPrintDiagUI.exe')) {
    $p = Join-Path $here $f
    if (-not (Test-Path -LiteralPath $p)) { throw ($f + ' 没有生成') }
    $b = [System.IO.File]::ReadAllBytes($p)
    $pe = [BitConverter]::ToInt32($b, 0x3C)
    $sub = [BitConverter]::ToUInt16($b, $pe + 0x5C)
    $mode = switch ($sub) { 2 { 'GUI' } 3 { 'Console' } default { '未知' } }
    Write-Host ('{0,-20} {1,7} 字节  subsystem={2} ({3})' -f $f, $b.Length, $sub, $mode)
}

Write-Host '打包完成。'
