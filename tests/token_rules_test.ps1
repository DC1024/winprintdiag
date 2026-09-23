# 打印机识别规则回归测试
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File tests\token_rules_test.ps1
# 做法：从..\WinPrintDiag.ps1 里把被测函数原文抽出来落盘再 dot-source，
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

    # 抽取方式：从 "function 名字 {" 一直取到第一个「顶格右括号」的行。
    # 因此工具里新增函数时，函数内部的所有右括号都必须保持缩进，否则会被提前截断。
    $fnNames = @(
        'Get-DispWidth'
        'Pad-R'
        'Get-ConnKind'
        'Get-PrinterTokens'
        'Get-PrinterUsage'
        'Get-UsageOf'
        'Get-UsageText'
        'Get-DupVerdict'
        'Get-DupAdvice'
        'Get-Trunc'
    )
    $nl = [char]13 + [char]10
    $sb = New-Object System.Text.StringBuilder
    foreach ($fn in $fnNames) {
        $m = [regex]::Match($src, '(?s)function ' + [regex]::Escape($fn) + ' \{.*?\r?\n\}')
        if (-not $m.Success) { throw ('抽不出 ' + $fn + '（函数名或格式被改过？）') }
        [void]$sb.Append($m.Value + $nl)
    }
    [System.IO.File]::WriteAllText($fnFile, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
    [void]$out.Add('函数抽取: ' + $fnNames.Count + '/' + $fnNames.Count + ' 全部成功')
    . $fnFile

    function Add-Result([bool]$ok, [string]$line, [string]$detail) {
        $tag = 'OK  '
        if (-not $ok) { $tag = 'FAIL'; $script:fail = $script:fail + 1 }
        [void]$out.Add(('[{0}] {1}' -f $tag, $line))
        if (-not [string]::IsNullOrEmpty($detail)) { [void]$out.Add('        ' + $detail) }
    }

    # ---------- 连接方式识别 ----------
    function Check-Conn([string]$port, [string]$want) {
        $got = Get-ConnKind -Port $port
        Add-Result ($got -eq $want) ('端口 "{0}" -> "{1}"  期望 "{2}"' -f $port, $got, $want) ''
    }

    # ---------- 同一设备判定 ----------
    function Check-Pair([bool]$wantWarn, [string]$n1, [string]$p1, [string]$n2, [string]$p2, [string]$desc) {
        $a = @(Get-PrinterTokens -Name $n1 -Port $p1)
        $b = @(Get-PrinterTokens -Name $n2 -Port $p2)
        $shared = @($a | Where-Object { $b -contains $_ })
        $hit = ($shared.Count -gt 0)
        Add-Result ($hit -eq $wantWarn) $desc ''
        [void]$out.Add(('        A = {0}  |  {1}' -f $n1, $p1))
        [void]$out.Add(('        B = {0}  |  {1}' -f $n2, $p2))
        [void]$out.Add(('        期望告警={0}  实际={1}  共有标识=[{2}]' -f $wantWarn, $hit, ($shared -join ',')))
    }

    # ---------- 使用记录查询 ----------
    function Check-Usage([string]$desc, $usage, [string]$name, [string]$port, [int]$wantTimes, [string]$wantText) {
        $u = Get-UsageOf -Usage $usage -Name $name -Port $port
        $txt = Get-UsageText -U $u
        $ok = (([int]$u.Times) -eq $wantTimes) -and ($txt -eq $wantText)
        Add-Result $ok $desc ('        name="{0}" port="{1}" -> Times={2} 文本="{3}"  期望 Times={4} 文本="{5}"' -f $name, $port, $u.Times, $txt, $wantTimes, $wantText)
    }

    # ---------- 「该留哪条」判定 ----------
    function Check-Verdict([string]$desc, [hashtable]$p, [string]$wantKeep, [bool]$wantBoth) {
        $v = Get-DupVerdict @p
        $ok = ($v.Keep -eq $wantKeep) -and ($v.BothUsed -eq $wantBoth)
        Add-Result $ok $desc ('        Keep="{0}" (期望"{1}")  BothUsed={2} (期望{3})' -f $v.Keep, $wantKeep, $v.BothUsed, $wantBoth)
        [void]$out.Add('        依据: ' + $v.Reason)
        if ($v.Keep -ne '') {
            [void]$out.Add(('        保留="{0}"  删除="{1}"' -f $v.KeepName, $v.DropName))
        }
    }

    # ---------- 多行建议块的形状 ----------
    function Check-Advice([string]$desc, $A, $B, $usage, [string]$mustContain, [bool]$wantKeepBlock) {
        $adv = ''
        try { $adv = [string](Get-DupAdvice -A $A -B $B -Usage $usage) }
        catch {
            Add-Result $false $desc ('        抛异常: ' + $_.Exception.Message)
            return
        }
        $ok = $true
        $why = New-Object System.Collections.ArrayList
        if ($adv -notlike ("`n" + '*')) { $ok = $false; [void]$why.Add('不以换行开头（会导致报告里和告警正文粘在一起）') }
        if ($mustContain -ne '' -and $adv -notlike ('*' + $mustContain + '*')) { $ok = $false; [void]$why.Add(('缺少关键字 "' + $mustContain + '"')) }
        $hasKeep = ($adv -like '*建议保留:*')
        if ($hasKeep -ne $wantKeepBlock) { $ok = $false; [void]$why.Add(('建议保留块 期望=' + $wantKeepBlock + ' 实际=' + $hasKeep)) }
        # 报告里告警行的写法是 Add-Line ('  ' + $f)，即前缀 "  [警告] "（显示宽度 9）。
        # 建议块除首行空行外，每行至少要缩进 7 格，加上这 2 格才不会跑到 [警告] 之外；
        # 表格明细行是 7+2=9 格的子表，属于允许的更深缩进。
        $wantPad = (Get-DispWidth -Text ('  ' + '[' + '警告' + '] ')) - 2
        $lines = $adv.Split([char]10)
        for ($k = 1; $k -lt $lines.Count; $k++) {
            $ln = $lines[$k]
            if ($ln.Length -eq 0) { continue }
            $pad = $ln.Length - $ln.TrimStart(' ').Length
            if ($pad -lt $wantPad) {
                $ok = $false
                [void]$why.Add(('第 ' + $k + ' 行只缩进 ' + $pad + ' 格，少于 ' + $wantPad + ' 格，会跑到 [警告] 外面: ' + $ln))
                break
            }
        }
        Add-Result $ok $desc ('        ' + ($why -join ' / '))
        foreach ($ln in $lines) { [void]$out.Add('        |' + $ln) }
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

    # ---- 构造一份使用记录，模拟「本机实测」：IPP 用过 6 次，WSD 从未用过 ----
    $byPort = @{}
    $byPort['ipp_pantum-13cedd_1'] = [ordered]@{ Times = 6; Last = '2026-09-23 13:18:05'; Name = 'Pantum M6200NW Series 0001' }
    $byName = @{}
    $byName['pantum m6200nw series 0001'] = [ordered]@{ Times = 6; Last = '2026-09-23 13:18:05'; Port = 'IPP_Pantum-13CEDD_1' }
    $mock = [ordered]@{
        Total = 6; Matched = 6; Error = $false; Disabled = $false
        ByPort = $byPort; ByName = $byName
    }
    $mockDisabled = [ordered]@{ Total = 0; Matched = 0; Error = $false; Disabled = $true; ByPort = @{}; ByName = @{} }
    $mockError = [ordered]@{ Total = 0; Matched = 0; Error = $true; Disabled = $false; ByPort = @{}; ByName = @{} }
    $mockEmpty = [ordered]@{ Total = 0; Matched = 0; Error = $false; Disabled = $false; ByPort = @{}; ByName = @{} }

    [void]$out.Add('')
    [void]$out.Add('======== 使用记录查询 ========')
    Check-Usage '按端口命中' $mock 'Pantum M6200NW Series 0001' 'IPP_Pantum-13CEDD_1' 6 '最近 2026-09-23 13:18:05，共 6 次'
    Check-Usage '端口没命中，退回按名字命中' $mock 'Pantum M6200NW Series 0001' 'SOMETHING-ELSE' 6 '最近 2026-09-23 13:18:05，共 6 次'
    Check-Usage '从未用过的条目' $mock 'Pantum-13CEDD (M6200NW series)' 'WSD-c0c4001a-e93a-4cd5-8f4f-1234' 0 '无使用记录'
    Check-Usage '使用记录整体为空' $mockEmpty 'Pantum M6200NW Series 0001' 'IPP_Pantum-13CEDD_1' 0 '无使用记录'
    Check-Usage 'Usage 为 null 时不能抛异常' $null 'Pantum M6200NW Series 0001' 'IPP_Pantum-13CEDD_1' 0 '无使用记录'

    [void]$out.Add('')
    [void]$out.Add('======== 「该留哪条」判定 ========')
    # ① 只有一条有记录 -> 留它
    Check-Verdict '仅 A 有打印记录' @{ NameA = 'Pantum M6200NW Series 0001'; UseA = 6; LastA = '2026-09-23 13:18:05'; IppA = $false
                                         NameB = 'Pantum-13CEDD (M6200NW series)'; UseB = 0; LastB = ''; IppB = $false } 'A' $false
    Check-Verdict '仅 B 有打印记录' @{ NameA = 'Pantum-13CEDD (M6200NW series)'; UseA = 0; LastA = ''; IppA = $false
                                         NameB = 'Pantum M6200NW Series 0001'; UseB = 6; LastB = '2026-09-23 13:18:05'; IppB = $false } 'B' $false
    # ② 两条都在用 -> 留最近用过的，并标记 BothUsed（提醒人工确认）
    Check-Verdict '两条都在用，A 更近' @{ NameA = 'A 条目'; UseA = 2; LastA = '2026-09-23 13:18:05'; IppA = $false
                                           NameB = 'B 条目'; UseB = 9; LastB = '2026-09-20 09:00:00'; IppB = $false } 'A' $true
    Check-Verdict '两条都在用，B 更近' @{ NameA = 'A 条目'; UseA = 9; LastA = '2026-09-20 09:00:00'; IppA = $false
                                           NameB = 'B 条目'; UseB = 2; LastB = '2026-09-23 13:18:05'; IppB = $false } 'B' $true
    # ③ 都没记录 -> 退化为按驱动判断，留厂商驱动那条（IppA=$true 表示 A 是通用 IPP 类驱动）
    Check-Verdict '都没记录，A 是通用 IPP 类驱动' @{ NameA = 'A 通用驱动'; UseA = 0; LastA = ''; IppA = $true
                                                   NameB = 'B 厂商驱动'; UseB = 0; LastB = ''; IppB = $false } 'B' $false
    Check-Verdict '都没记录，B 是通用 IPP 类驱动' @{ NameA = 'A 厂商驱动'; UseA = 0; LastA = ''; IppA = $false
                                                   NameB = 'B 通用驱动'; UseB = 0; LastB = ''; IppB = $true } 'A' $false
    # ④ 都没记录、驱动同类 -> 不给结论（空返回，由 Get-DupAdvice 走「无法判断」分支）
    Check-Verdict '都没记录，驱动同类 -> 不下结论' @{ NameA = 'A 条目'; UseA = 0; LastA = ''; IppA = $false
                                                     NameB = 'B 条目'; UseB = 0; LastB = ''; IppB = $false } '' $false
    Check-Verdict '都没记录，都是通用 IPP -> 不下结论' @{ NameA = 'A 条目'; UseA = 0; LastA = ''; IppA = $true
                                                         NameB = 'B 条目'; UseB = 0; LastB = ''; IppB = $true } '' $false
    # 边界：次数相同、Last 为空时不能崩，且必须落在某一侧
    Check-Verdict '两条都用过且次数相同（Last 同为空）' @{ NameA = 'A 条目'; UseA = 1; LastA = ''; IppA = $false
                                                          NameB = 'B 条目'; UseB = 1; LastB = ''; IppB = $false } 'B' $true

    [void]$out.Add('')
    [void]$out.Add('======== 多行建议块 ========')
    $A = [ordered]@{ Name = 'Pantum M6200NW Series 0001'; Port = 'IPP_Pantum-13CEDD_1'; Conn = '网络 IPP'; IsIpp = $false }
    $B = [ordered]@{ Name = 'Pantum-13CEDD (M6200NW series)'; Port = 'WSD-c0c4001a-e93a-4cd5-8f4f-1234'; Conn = '网络 WSD'; IsIpp = $false }
    Check-Advice '实测场景：IPP 在用 / WSD 无记录' $A $B $mock '建议删除: 「Pantum-13CEDD (M6200NW series)」' $true
    Check-Advice '日志未启用时的兜底文案' $A $B $mockDisabled '打印操作日志未启用' $false
    Check-Advice '日志不可读时的兜底文案' $A $B $mockError '打印操作日志不可读' $false
    Check-Advice '日志一条记录都没有时的兜底文案' $A $B $mockEmpty '没有任何历史打印记录' $false

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
