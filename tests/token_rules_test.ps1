# 打印机识别规则回归测试
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File tests\token_rules_test.ps1
# 做法：从..\WinPrintDiag.ps1 里把 Get-PrinterTokens / Get-ConnKind 两个函数原文抽出来落盘再 dot-source，
#       保证测的是工具里的真代码，而不是复制出来的一份副本。
# 注意：本文件必须保存为 UTF-8 带 BOM，否则 PowerShell 5.1 会按 ANSI 解码，中文全部变乱码。

$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($here)) { $here = Split-Path -Parent $MyInvocation.MyCommand.Definition }
$toolPs1 = Join-Path (Split-Path -Parent $here) 'WinPrintDiag.ps1'
$fnFile = Join-Path $here '_functions_extracted.ps1'
$resultFile = Join-Path $here 'token_rules_test_result.txt'

$out = New-Object System.Collections.ArrayList
$script:fail = 0
[void]$out.Add('打印机识别规则回归测试  ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
[void]$out.Add('被测脚本: ' + $toolPs1)

try {
    if (-not (Test-Path -LiteralPath $toolPs1)) { throw '找不到 WinPrintDiag.ps1' }
    $src = [System.IO.File]::ReadAllText($toolPs1, [System.Text.Encoding]::UTF8)

    $m1 = [regex]::Match($src, '(?s)function Get-PrinterTokens \{.*?\r?\n\}')
    $m2 = [regex]::Match($src, '(?s)function Get-ConnKind \{.*?\r?\n\}')
    if (-not $m1.Success) { throw '抽不出 Get-PrinterTokens（函数名或格式被改过？）' }
    if (-not $m2.Success) { throw '抽不出 Get-ConnKind（函数名或格式被改过？）' }
    [void]$out.Add('函数抽取: Get-PrinterTokens/' + $m1.Success + '  Get-ConnKind/' + $m2.Success)

    $nl = [char]13 + [char]10
    [System.IO.File]::WriteAllText($fnFile, $m1.Value + $nl + $m2.Value, (New-Object System.Text.UTF8Encoding($true)))
    . $fnFile

    function Check-Conn([string]$port, [string]$want) {
        $got = Get-ConnKind -Port $port
        $tag = 'OK  '
        if ($got -ne $want) { $tag = 'FAIL'; $script:fail = $script:fail + 1 }
        [void]$out.Add(('[{0}] 端口 "{1}" -> "{2}"  期望 "{3}"' -f $tag, $port, $got, $want))
    }

    function Check-Pair([bool]$wantWarn, [string]$n1, [string]$p1, [string]$n2, [string]$p2, [string]$desc) {
        $a = @(Get-PrinterTokens -Name $n1 -Port $p1)
        $b = @(Get-PrinterTokens -Name $n2 -Port $p2)
        $shared = @($a | Where-Object { $b -contains $_ })
        $hit = ($shared.Count -gt 0)
        $tag = 'OK  '
        if ($hit -ne $wantWarn) { $tag = 'FAIL'; $script:fail = $script:fail + 1 }
        [void]$out.Add(('[{0}] {1}' -f $tag, $desc))
        [void]$out.Add(('        A = {0}  |  {1}' -f $n1, $p1))
        [void]$out.Add(('        B = {0}  |  {1}' -f $n2, $p2))
        [void]$out.Add(('        期望告警={0}  实际={1}  共有标识=[{2}]' -f $wantWarn, $hit, ($shared -join ',')))
    }

    [void]$out.Add('')
    [void]$out.Add('======== 连接方式识别 ========')
    Check-Conn 'USB001' 'USB 直连'
    Check-Conn 'USB002' 'USB 直连'
    Check-Conn 'WSD-c0c4001a-e93a-4cd5-8f4f-1234' '网络 WSD'
    Check-Conn 'IPP_Pantum-13CEDD_1' '网络 IPP'
    Check-Conn 'IP_192.168.31.50' '网络 IP'
    Check-Conn 'PORTPROMPT:' '虚拟/本地'
    Check-Conn '\\.\pipe\PDFPrint' '虚拟/本地'
    Check-Conn 'Documents\*.pdf' '虚拟/本地'
    Check-Conn 'PDF-XChange5-ABBYY-FR15' '虚拟/本地'
    Check-Conn 'Kingsoft Virtual Printer Port' '虚拟/本地'
    Check-Conn 'HP1A2B3C' '其他'

    [void]$out.Add('')
    [void]$out.Add('======== 同一设备判定：应告警 ========')
    Check-Pair $true 'Pantum M6200NW Series 0001' 'IPP_Pantum-13CEDD_1' 'Pantum-13CEDD (M6200NW series)' 'WSD-c0c4001a-e93a-4cd5-8f4f-1234' '同机 IPP + WSD 双条目（实测机型）'
    Check-Pair $true 'Brother HL-L2350DW' 'USB001' 'Brother HL-L2350DW' 'WSD-1122334455667788' '典型：USB + 网络 WSD 双条目'
    Check-Pair $true 'HP LaserJet M1136 MFP' 'USB001' 'HP LaserJet M1136 MFP' 'IP_192.168.31.77' '典型：USB + 网络 IP 双条目'

    [void]$out.Add('')
    [void]$out.Add('======== 同一设备判定：不应误报 ========')
    Check-Pair $false 'Adobe PDF' 'Documents\*.pdf' 'PDF-XChange 5.0 for ABBYY FineReader 15' 'PDF-XChange5-ABBYY-FR15' '两个虚拟 PDF 打印机（名字都含 PDF）'
    Check-Pair $false 'Microsoft Print to PDF' 'PORTPROMPT:' 'PDF24' '\\.\pipe\PDFPrint' '两个虚拟打印机'
    Check-Pair $false 'Pantum M6200NW Series 0001' 'IPP_Pantum-13CEDD_1' 'Pantum M6500 Series 0001' 'IPP_Pantum-AA0001_1' '不同机型，仅数字后缀相同'
    Check-Pair $false '导出为WPS PDF' 'Kingsoft Virtual Printer Port' 'Adobe PDF' 'Documents\*.pdf' '两个虚拟打印机（含中文名）'
    Check-Pair $false 'HP LaserJet 1020 (Copy 1)' 'USB001' 'HP LaserJet 1020 (Copy 2)' 'USB002' '同型号重复副本（保守取舍：不报）'
    Check-Pair $false 'OneNote (Desktop)' 'Microsoft.Office.OneNote_2015' 'Microsoft XPS Document Writer' 'PORTPROMPT:' '系统自带虚拟打印机'

    [void]$out.Add('')
    [void]$out.Add('==== 失败用例数: ' + $script:fail + ' ====')
}
catch {
    $script:fail = $script:fail + 1
    [void]$out.Add('EXCEPTION: ' + $_.Exception.GetType().FullName)
    [void]$out.Add($_.Exception.Message)
}
finally {
    if (Test-Path -LiteralPath $fnFile) { Remove-Item -LiteralPath $fnFile -Force -ErrorAction SilentlyContinue }
    [System.IO.File]::WriteAllLines($resultFile, $out, (New-Object System.Text.UTF8Encoding($false)))
}

Write-Output ('失败用例数: ' + $script:fail)
Write-Output ('明细: ' + $resultFile)
if ($script:fail -gt 0) { exit 1 } else { exit 0 }
