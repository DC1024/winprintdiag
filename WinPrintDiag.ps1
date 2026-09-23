# WinPrintDiag - Windows print subsystem portable diagnostic tool
# Distributed as-is, single file, no dependencies, PowerShell 5.1 compatible.
# License: use freely.
param(
    [switch]$Repair,
    [switch]$ClearQueue,
    [switch]$OpenReport,
    [string]$OutDir = ''
)

$ErrorActionPreference = 'Continue'
try {
    if ([Console]::OutputEncoding.WebName -ne 'gb2312' -and [Console]::OutputEncoding.CodePage -ne 936) {
        [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(936)
    }
}
catch {
}
$ToolVersion = '1.0'
$StartTime = Get-Date

$Report = New-Object System.Collections.ArrayList
$Flags  = New-Object System.Collections.ArrayList

function Add-Line {
    param([AllowEmptyString()][string]$Text)
    [void]$Report.Add($Text)
}

function Add-Blank {
    [void]$Report.Add('')
}

function Add-Flag {
    param([string]$Level, [string]$Text)
    $tag = $Level
    switch ($Level) {
        'CRITICAL' { $tag = '严重' }
        'WARN' { $tag = '警告' }
        'INFO' { $tag = '提示' }
    }
    [void]$Flags.Add(('[' + $tag + '] ' + $Text))
}

function Get-DispWidth {
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return 0 }
    $w = 0
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int]$ch
        if (($code -ge 0x1100 -and $code -le 0x115F) -or
            ($code -ge 0x2E80 -and $code -le 0xA4CF) -or
            ($code -ge 0xAC00 -and $code -le 0xD7A3) -or
            ($code -ge 0xF900 -and $code -le 0xFAFF) -or
            ($code -ge 0xFE30 -and $code -le 0xFE6F) -or
            ($code -ge 0xFF00 -and $code -le 0xFF60) -or
            ($code -ge 0xFFE0 -and $code -le 0xFFE6)) { $w = $w + 2 }
        else { $w = $w + 1 }
    }
    return $w
}

function Pad-R {
    param([AllowEmptyString()][string]$Text, [int]$Width)
    $w = Get-DispWidth -Text $Text
    if ($w -ge $Width) { return $Text }
    return ($Text + (' ' * ($Width - $w)))
}

function Get-SigText {
    param([AllowEmptyString()][string]$Text)
    switch ($Text) {
        'Valid' { return '有效' }
        'NotSigned' { return '未签名' }
        'HashMismatch' { return '哈希不匹配' }
        'NotTrusted' { return '不受信任' }
        'UnknownError' { return '未知错误' }
        'NotSupportedFileFormat' { return '格式不支持' }
        'Incompatible' { return '不兼容' }
        'MISSING' { return '文件缺失' }
        'UNKNOWN' { return '无法判定' }
    }
    if ([string]::IsNullOrEmpty($Text)) { return '（空）' }
    return $Text
}

function Get-SvcText {
    param([AllowEmptyString()][string]$Text)
    switch ($Text) {
        'Running' { return '运行中' }
        'Stopped' { return '已停止' }
        'StartPending' { return '正在启动' }
        'StopPending' { return '正在停止' }
        'Paused' { return '已暂停' }
        'Automatic' { return '自动' }
        'Manual' { return '手动' }
        'Disabled' { return '已禁用' }
        'AutomaticDelayedStart' { return '自动（延迟启动）' }
    }
    return $Text
}

function Get-DateText {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd') }
    $s = [string]$Value
    if ($s -match '^\d{8}$') { return ($s.Substring(0, 4) + '-' + $s.Substring(4, 2) + '-' + $s.Substring(6, 2)) }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse($s, [ref]$dt)) { return $dt.ToString('yyyy-MM-dd') }
    return $s
}

function Get-ConnKind {
    param([AllowEmptyString()][string]$Port)
    if ([string]::IsNullOrEmpty($Port)) { return '未知' }
    if ($Port -match '^USB\d+$') { return 'USB 直连' }
    if ($Port -match '^WSD') { return '网络 WSD' }
    if ($Port -match '^IPP') { return '网络 IPP' }
    if ($Port -match '^IP_|^IP-') { return '网络 IP' }
    if ($Port -match '^TCP') { return '网络 TCP' }
    if ($Port -match '^PORTPROMPT|^\\\\|pipe|\.pdf$|^FILE:|^SHRFAX|^nul|virtual|pdf') { return '虚拟/本地' }
    return '其他'
}

# 从「打印机名 + 端口名」提取设备标识词，用来判断两条记录是不是同一台物理打印机。
# 只保留「长度 >= 5 且同时含字母和数字」的词：这类词通常是机型或主机名后缀（如 13cedd / m6200nw），
# 而 pdf / series / 0001 这种通用词或纯数字会被排除，避免把不同机型误判成同一台。
function Get-PrinterTokens {
    param([AllowEmptyString()][string]$Name, [AllowEmptyString()][string]$Port)
    $raw = ($Name + ' ' + $Port).ToLower()
    $parts = [regex]::Split($raw, '[^a-z0-9]+')
    $res = New-Object System.Collections.ArrayList
    foreach ($t in $parts) {
        if ($t.Length -lt 5) { continue }
        if ($t -notmatch '[a-z]') { continue }
        if ($t -notmatch '[0-9]') { continue }
        if (-not $res.Contains($t)) { [void]$res.Add($t) }
    }
    return $res
}

function Get-Trunc {
    param([AllowEmptyString()][string]$Text, [int]$Width)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    if ($Width -le 1) { return '' }
    if ((Get-DispWidth -Text $Text) -le $Width) { return $Text }
    $sb = New-Object System.Text.StringBuilder
    $w = 0
    foreach ($ch in $Text.ToCharArray()) {
        $cw = Get-DispWidth -Text ([string]$ch)
        if (($w + $cw) -gt ($Width - 1)) { break }
        [void]$sb.Append($ch)
        $w = $w + $cw
    }
    return ($sb.ToString() + '~')
}

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object Security.Principal.WindowsPrincipal($id)
    return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-UniStrings {
    param([string]$Path, [int]$MinLen = 3, [int]$MaxItems = 12)
    $res = New-Object System.Collections.ArrayList
    if (-not (Test-Path -LiteralPath $Path)) { return $res }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $run = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt ($bytes.Length - 1)) {
        $lo = [int]$bytes[$i]
        $hi = [int]$bytes[$i + 1]
        if ($hi -eq 0 -and $lo -ge 32 -and $lo -lt 127) {
            [void]$run.Append([char]$lo)
        }
        else {
            if ($run.Length -ge $MinLen) {
                [void]$res.Add($run.ToString())
                if ($res.Count -ge $MaxItems) { break }
            }
            [void]$run.Clear()
        }
        $i = $i + 2
    }
    return $res
}

function Get-PrintBinaryRow {
    param([string]$Path)
    $row = [ordered]@{
        Name    = ''
        Size    = 0
        Version = ''
        MTime   = ''
        Sig     = 'MISSING'
        Note    = ''
    }
    $row.Name = Split-Path -Path $Path -Leaf
    if (-not (Test-Path -LiteralPath $Path)) { return $row }
    $fi = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($null -eq $fi) { return $row }
    $row.Size = $fi.Length
    $row.Version = $fi.VersionInfo.ProductVersion
    $row.MTime = $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
    try {
        $sig = Get-AuthenticodeSignature -FilePath $Path -ErrorAction Stop
        $row.Sig = [string]$sig.Status
    }
    catch {
        $row.Sig = 'UNKNOWN'
    }
    return $row
}

# ---------- resolve output dir ----------
if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $defPath = [string]$MyInvocation.MyCommand.Definition
    $cmdPath = [string]$MyInvocation.MyCommand.Path
    $scriptRoot = ''
    if (-not [string]::IsNullOrWhiteSpace($cmdPath)) {
        $scriptRoot = Split-Path -Parent $cmdPath
    }
    if ([string]::IsNullOrWhiteSpace($scriptRoot) -and $PSScriptRoot) {
        $scriptRoot = $PSScriptRoot
    }
    if ([string]::IsNullOrWhiteSpace($scriptRoot) -and $defPath.Length -lt 1024 -and $defPath -notmatch "`n") {
        try {
            $scriptRoot = Split-Path -Parent $defPath
        }
        catch {
            $scriptRoot = ''
        }
    }
    if ([string]::IsNullOrWhiteSpace($scriptRoot)) {
        $scriptRoot = [System.IO.Directory]::GetCurrentDirectory()
    }
    $OutDir = $scriptRoot
}
if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}

$IsAdmin = Test-IsAdmin

$modeText = '仅诊断（只读）'
if ($Repair) { $modeText = '诊断 + 修复' }
Add-Line ('WinPrintDiag ' + $ToolVersion + '  生成时间 ' + $StartTime.ToString('yyyy-MM-dd HH:mm:ss'))
$adminText = '否'
if ($IsAdmin) { $adminText = '是' }
Add-Line ('管理员权限 : ' + $adminText)
Add-Line ('运行模式   : ' + $modeText)
Add-Blank

# ---------- 1. environment ----------
Add-Line '===== [1] 环境 ====='
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
if ($os) {
    Add-Line ('操作系统   : ' + $os.Caption)
    Add-Line ('内部版本   : ' + $os.BuildNumber + ' / ' + $os.Version)
}
$cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
if ($cv) {
    Add-Line ('功能版本   : ' + $cv.DisplayVersion + '  UBR: ' + $cv.UBR)
    if ($cv.InstallDate) {
        $ins = [datetime]'1970-01-01'
        Add-Line ('系统安装日 : ' + $ins.AddSeconds($cv.InstallDate).ToLocalTime().ToString('yyyy-MM-dd'))
    }
}
Add-Line ('PowerShell : ' + $PSVersionTable.PSVersion.ToString())
Add-Line ('当前用户   : ' + [Security.Principal.WindowsIdentity]::GetCurrent().Name)
Add-Blank

# ---------- 2. print spooler service ----------
Add-Line '===== [2] 打印服务状态 ====='
$svc = Get-Service Spooler -ErrorAction SilentlyContinue
if ($svc) {
    Add-Line ('服务状态   : ' + (Get-SvcText -Text ([string]$svc.Status)) + '     启动类型: ' + (Get-SvcText -Text ([string]$svc.StartType)))
}
else {
    Add-Line '服务状态   : 未找到该服务'
    Add-Flag 'CRITICAL' '未找到 Print Spooler 服务（系统打印子系统可能已损坏）'
}
$procList = @(Get-Process spoolsv -ErrorAction SilentlyContinue)
if ($procList.Count -eq 0) {
    Add-Line '进程状态   : 未运行'
}
else {
    Add-Line ('进程状态   : 运行中，实例数 ' + $procList.Count)
    $procStartTime = $null
    try {
        $procStartTime = $procList[0].StartTime
    }
    catch {
        $procStartTime = $null
    }
    if ($null -ne $procStartTime) {
        $uptimeSec = [int]((Get-Date) - $procStartTime).TotalSeconds
        Add-Line ('启动时间   : ' + $procStartTime.ToString('yyyy-MM-dd HH:mm:ss') + '  (已运行 ' + $uptimeSec + ' 秒)')
        if ($uptimeSec -lt 300) {
            Add-Flag 'WARN' ('spoolsv 仅运行 ' + $uptimeSec + ' 秒，疑似刚重启或处于崩溃循环')
        }
    }
    else {
        Add-Line '启动时间   : （受保护进程，改查事件日志）'
        $lastStartTime = $null
        try {
            $stateEv = @(Get-WinEvent -FilterHashtable @{LogName = 'System'; Id = 7036; StartTime = (Get-Date).AddDays(-2)} -ErrorAction SilentlyContinue |
                Where-Object { $_.Message -match 'Print Spooler|Spooler' })
            if ($stateEv.Count -gt 0) {
                $lastStartTime = ($stateEv | Sort-Object TimeCreated | Select-Object -Last 1).TimeCreated
            }
        }
        catch {
            $lastStartTime = $null
        }
        if ($null -ne $lastStartTime) {
            $uptimeSec2 = [int]((Get-Date) - $lastStartTime).TotalSeconds
            Add-Line ('最近状态   : ' + $lastStartTime.ToString('yyyy-MM-dd HH:mm:ss') + '  (约已运行 ' + $uptimeSec2 + ' 秒)')
            if ($uptimeSec2 -lt 300) {
                Add-Flag 'WARN' ('spoolsv 仅运行约 ' + $uptimeSec2 + ' 秒，疑似刚重启或处于崩溃循环')
            }
        }
        else {
            Add-Line '最近状态   : 不可用（本机不记录服务启停事件）'
            Add-Line '            崩溃循环判定改由“近期崩溃次数”承担'
        }
    }
}
Add-Blank

# ---------- 3. crash history ----------
Add-Line '===== [3] 崩溃历史 ====='
$crashEvents = $null
try {
    $crashEvents = @(Get-WinEvent -FilterHashtable @{LogName = 'Application'; Id = 1000} -ErrorAction SilentlyContinue |
        Where-Object { $_.Message -match 'spoolsv' })
}
catch {
    $crashEvents = @()
}
Add-Line ('spoolsv 崩溃事件（应用程序日志 1000）: ' + $crashEvents.Count)
if ($crashEvents.Count -gt 0) {
    $sorted = $crashEvents | Sort-Object TimeCreated
    Add-Line ('最早       : ' + $sorted[0].TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))
    Add-Line ('最近       : ' + $sorted[$sorted.Count - 1].TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))
    Add-Line '-- 按日统计（最近 14 天）--'
    $grouped = $crashEvents | ForEach-Object { $_.TimeCreated.ToString('yyyy-MM-dd') } | Group-Object | Sort-Object Name -Descending | Select-Object -First 14
    foreach ($g in $grouped) { Add-Line ('  ' + $g.Name + '  x' + $g.Count) }
    Add-Line '-- 出错模块统计 --'
    $mods = @{}
    foreach ($e in $crashEvents) {
        $match = [regex]::Match($e.Message, '(?m)^出错模块名称[:：]\s*([^,，]+)')
        if ($match.Success) {
            $mod = $match.Groups[1].Value.Trim()
            if ($mods.ContainsKey($mod)) { $mods[$mod] = $mods[$mod] + 1 }
            else { $mods[$mod] = 1 }
        }
    }
    foreach ($k in $mods.Keys) { Add-Line ('  ' + $k + '  x' + $mods[$k]) }
    $recentCut = (Get-Date).AddMinutes(-30)
    $recent = @($crashEvents | Where-Object { $_.TimeCreated -gt $recentCut })
    Add-Line ('近 30 分钟崩溃次数: ' + $recent.Count)
    if ($recent.Count -ge 3) { Add-Flag 'CRITICAL' '近 30 分钟内崩溃 3 次以上，处于崩溃循环' }
}
Add-Blank

# ---------- 4. service control manager events ----------
Add-Line '===== [4] 服务控制管理器事件 ====='
try {
    $scm = @(Get-WinEvent -FilterHashtable @{LogName = 'System'; ProviderName = 'Service Control Manager'} -ErrorAction SilentlyContinue |
        Where-Object { $_.Message -match 'Spooler' -and ($_.Id -eq 7031 -or $_.Id -eq 7034) })
}
catch {
    $scm = @()
}
Add-Line ('Spooler 服务异常终止事件（7031/7034）: ' + $scm.Count)
if ($scm.Count -gt 0) {
    $scmSorted = $scm | Sort-Object TimeCreated
    Add-Line ('最早       : ' + $scmSorted[0].TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))
    Add-Line ('最近       : ' + $scmSorted[$scmSorted.Count - 1].TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))
}
Add-Line '-- 日志可视窗口（决定证据能回溯多远）--'
foreach ($lg in @('Application', 'System')) {
    try {
        $info = Get-WinEvent -ListLog $lg -ErrorAction Stop
        $old = Get-WinEvent -LogName $lg -MaxEvents 1 -Oldest -ErrorAction SilentlyContinue
        $lgName = '应用程序'
        if ($lg -eq 'System') { $lgName = '系统' }
        Add-Line ('  ' + $lgName + ': 最早=' + $old.TimeCreated.ToString('yyyy-MM-dd') + '  容量=' + [int]($info.FileSize / 1MB) + 'MB/' + [int]($info.MaximumSizeInBytes / 1MB) + 'MB')
    }
    catch {
        Add-Line ('  ' + $lg + ': 不可用')
    }
}
Add-Blank

# ---------- 5. print stack file integrity ----------
Add-Line '===== [5] 打印组件文件完好性 ====='
$coreFiles = @(
    'C:\Windows\System32\spoolsv.exe',
    'C:\Windows\System32\usbmon.dll',
    'C:\Windows\System32\localspl.dll',
    'C:\Windows\System32\win32spl.dll',
    'C:\Windows\System32\winspool.drv',
    'C:\Windows\System32\spoolss.dll',
    'C:\Windows\System32\tcpmon.dll',
    'C:\Windows\System32\spool\prtprocs\x64\winprint.dll'
)
Add-Line ((Pad-R '文件' 18) + (Pad-R '大小' 10) + (Pad-R '版本' 19) + (Pad-R '修改时间' 18) + '签名')
$suspectFiles = New-Object System.Collections.ArrayList
foreach ($f in $coreFiles) {
    $row = Get-PrintBinaryRow -Path $f
    Add-Line ((Pad-R ([string]$row.Name) 18) + (Pad-R ([string]$row.Size) 10) + (Pad-R ([string]$row.Version) 19) + (Pad-R ([string]$row.MTime) 18) + (Get-SigText -Text ([string]$row.Sig)))
    if ($row.Sig -ne 'Valid') {
        [void]$suspectFiles.Add($f)
        Add-Flag 'CRITICAL' ('系统打印组件签名异常: ' + $row.Name + ' -> ' + $row.Sig)
    }
}
Add-Blank

Add-Line '-- 组件存储里的 spoolsv.exe 候选 --'
$sxscandidates = @()
try {
    $sxscandidates = @(Get-ChildItem 'C:\Windows\WinSxS' -Directory -Filter '*printing-spooler-core*' -ErrorAction SilentlyContinue |
        Where-Object { Test-Path (Join-Path $_.FullName 'spoolsv.exe') })
}
catch {
    $sxscandidates = @()
}
$targetSpool = 'C:\Windows\System32\spoolsv.exe'
$suspectSpoolsv = (Test-Path -LiteralPath $targetSpool) -and ((Get-AuthenticodeSignature -FilePath $targetSpool -ErrorAction SilentlyContinue).Status -ne 'Valid')
if ($suspectSpoolsv) {
    foreach ($d in $sxscandidates) {
        $sf = Join-Path $d.FullName 'spoolsv.exe'
        $fi = Get-Item -LiteralPath $sf -ErrorAction SilentlyContinue
        $sig = Get-AuthenticodeSignature -FilePath $sf -ErrorAction SilentlyContinue
        Add-Line ('  ' + $d.Name)
        Add-Line ('      体积=' + $fi.Length + '  版本=' + $fi.VersionInfo.ProductVersion + '  签名=' + (Get-SigText -Text ([string]$sig.Status)))
    }
    $cur = Get-Item -LiteralPath $targetSpool -ErrorAction SilentlyContinue
    Add-Line ('  当前磁盘文件: 体积=' + $cur.Length + '  版本=' + $cur.VersionInfo.ProductVersion + '  修改时间=' + $cur.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))
    Add-Flag 'CRITICAL' 'System32\spoolsv.exe 与组件存储不一致（可能被第三方替换）'
}
else {
    Add-Line '  （spoolsv.exe 签名有效，无需比对）'
}
Add-Blank

# ---------- 6. printers and ports (registry, works while spooler is down) ----------
Add-Line '===== [6] 打印机 / 端口 / 驱动配对 ====='
$printerKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers'
$printers = @()
try {
    $printers = @(Get-ChildItem $printerKey -ErrorAction Stop)
}
catch {
    Add-Line '无法读取打印机注册表项'
}
if ($printers.Count -eq 0) {
    Add-Line '未配置任何打印机'
}
else {
    $prnList = New-Object System.Collections.ArrayList
    foreach ($p in $printers) {
        $port = [string]$p.GetValue('Port')
        $drv = [string]$p.GetValue('Printer Driver')
        $o = [ordered]@{
            Name  = [string]$p.PSChildName
            Port  = $port
            Drv   = $drv
            Conn  = (Get-ConnKind -Port $port)
            Tok   = @(Get-PrinterTokens -Name ([string]$p.PSChildName) -Port $port)
            IsIpp = ($drv -match 'IPP Class Driver|Mopria')
        }
        [void]$prnList.Add($o)
    }

    Add-Line ((Pad-R '打印机' 40) + (Pad-R '端口' 28) + ' ' + (Pad-R '连接方式' 12) + '驱动')
    foreach ($o in $prnList) {
        Add-Line ((Pad-R (Get-Trunc -Text $o.Name -Width 40) 40) + (Pad-R (Get-Trunc -Text $o.Port -Width 28) 28) + ' ' + (Pad-R ([string]$o.Conn) 12) + ([string]$o.Drv))
    }

    # -- 同一台物理打印机是否被注册成了多个条目 --
    $dupNames = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $prnList.Count; $i++) {
        for ($j = $i + 1; $j -lt $prnList.Count; $j++) {
            $a = $prnList[$i]
            $b = $prnList[$j]
            $shared = @($a.Tok | Where-Object { $b.Tok -contains $_ })
            if ($shared.Count -eq 0) { continue }
            $usbCombo = (($a.Conn -match '^USB') -xor ($b.Conn -match '^USB'))
            $msg = '同一台打印机注册了多个条目: ' + $a.Name + ' [' + $a.Conn + '] 与 ' + $b.Name + ' [' + $b.Conn + ']  (共有标识: ' + ($shared -join '/') + ')'
            if ($usbCombo) {
                Add-Flag 'CRITICAL' ($msg + ' —— 同一台机器同时保留 USB 直连与网络两条通道，是最容易触发 spoolsv 崩溃的组合')
            }
            else {
                Add-Flag 'WARN' $msg
            }
            if (-not $dupNames.Contains([string]$a.Name)) { [void]$dupNames.Add([string]$a.Name) }
            if (-not $dupNames.Contains([string]$b.Name)) { [void]$dupNames.Add([string]$b.Name) }
        }
    }

    # -- USB 端口却挂在 IPP/Mopria 类驱动上 --
    foreach ($o in $prnList) {
        if (($o.Port -match '^USB\d+$') -and ($o.Drv -match 'IPP|Mopria|Class Driver')) {
            Add-Flag 'CRITICAL' ('端口/驱动错配: ' + $o.Name + '  (' + $o.Port + ' + ' + $o.Drv + ')')
        }
    }

    # -- 单独使用通用 IPP 类驱动（无厂商适配）；已被"重复条目"覆盖的不再重复报 --
    foreach ($o in $prnList) {
        if (-not $o.IsIpp) { continue }
        if ($dupNames.Contains([string]$o.Name)) { continue }
        Add-Flag 'INFO' ('使用通用 IPP 类驱动（无厂商适配）: ' + $o.Name + '  (' + $o.Drv + '，' + $o.Conn + ')')
    }
}
Add-Line '-- USB 监视器端口 --'
try {
    $usbPorts = @(Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\USB Monitor\Ports' -ErrorAction Stop)
    foreach ($u in $usbPorts) { Add-Line ('  ' + $u.PSChildName) }
    if ($usbPorts.Count -eq 0) { Add-Line '  （无）' }
}
catch {
    Add-Line '  （无法读取）'
}
Add-Blank

# ---------- 7. print queue ----------
Add-Line '===== [7] 打印队列 ====='
$queueDir = Join-Path $env:SystemRoot 'System32\spool\PRINTERS'
$qFiles = @()
try {
    $qFiles = @(Get-ChildItem -LiteralPath $queueDir -File -ErrorAction Stop)
}
catch {
    $qFiles = @()
}
Add-Line ('队列文件数 : ' + $qFiles.Count)
if ($qFiles.Count -gt 0) {
    $total = ($qFiles | Measure-Object Length -Sum).Sum
    Add-Line ('合计体积   : ' + [int]($total / 1MB) + ' MB')
    $newest = $qFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $oldest = $qFiles | Sort-Object LastWriteTime | Select-Object -First 1
    Add-Line ('最新任务   : ' + $newest.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))
    Add-Line ('最早任务   : ' + $oldest.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))
    Add-Flag 'WARN' ('打印队列里有 ' + $qFiles.Count + ' 个未完成任务')
    Add-Line '-- 解码队列任务描述（前 6 个）--'
    $shdFiles = @($qFiles | Where-Object { $_.Extension -eq '.SHD' } | Sort-Object Name | Select-Object -First 6)
    foreach ($s in $shdFiles) {
        $strings = Get-UniStrings -Path $s.FullName
        Add-Line ('  ' + $s.Name + ' : ' + (($strings | Select-Object -First 8) -join ' ~ '))
    }
}
else {
    Add-Line '队列为空（没有打印任务时即为正常）'
}
Add-Blank

# ---------- 8. change correlation ----------
Add-Line '===== [8] 变更相关性 ====='
Add-Line '-- 近期安装的补丁（45 天内）--'
try {
    $hot = @(Get-HotFix -ErrorAction SilentlyContinue | Sort-Object InstalledOn -Descending | Select-Object -First 10)
    foreach ($h in $hot) {
        Add-Line ('  ' + (Get-DateText -Value $h.InstalledOn) + '  ' + $h.HotFixID)
    }
}
catch {
    Add-Line '  （无法读取更新历史）'
}
Add-Line '-- 近期安装的软件（45 天内）--'
$thr = (Get-Date).AddDays(-45).ToString('yyyyMMdd')
$uninPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
$apps = @()
foreach ($up in $uninPaths) {
    $apps += @(Get-ItemProperty $up -ErrorAction SilentlyContinue | Where-Object { $_.InstallDate -and ($_.InstallDate -ge $thr) -and $_.DisplayName })
}
foreach ($a in ($apps | Sort-Object InstallDate -Descending | Select-Object -First 10)) {
    Add-Line ('  ' + (Get-DateText -Value $a.InstallDate) + '  ' + $a.DisplayName)
}
if ($apps.Count -eq 0) { Add-Line '  （无）' }
Add-Line '-- 意外关机 / 重启（30 天内）--'
try {
    $bad = @(Get-WinEvent -FilterHashtable @{LogName = 'System'; Id = 6008; StartTime = (Get-Date).AddDays(-30)} -ErrorAction SilentlyContinue)
    Add-Line ('  事件 6008 意外关机次数: ' + $bad.Count)
    foreach ($b in ($bad | Select-Object -First 5)) { Add-Line ('    ' + $b.TimeCreated.ToString('yyyy-MM-dd HH:mm')) }
    $boot = @(Get-WinEvent -FilterHashtable @{LogName = 'System'; Id = 6005; StartTime = (Get-Date).AddDays(-7)} -ErrorAction SilentlyContinue)
    Add-Line ('  近 7 天开机次数: ' + $boot.Count)
    if ($boot.Count -ge 15) { Add-Flag 'WARN' '近 7 天开机次数过多，系统可能反复重启' }
}
catch {
    Add-Line '  （无法读取）'
}
$mdCount = 0
if (Test-Path 'C:\Windows\Minidump') {
    $mdCount = @(Get-ChildItem 'C:\Windows\Minidump' -File -ErrorAction SilentlyContinue).Count
}
Add-Line ('  蓝屏转储文件: ' + $mdCount)
Add-Blank

# ---------- 9. audit log status ----------
Add-Line '===== [9] 打印审计日志 ====='
try {
    $opLog = Get-WinEvent -ListLog 'Microsoft-Windows-PrintService/Operational' -ErrorAction Stop
    $opText = '否'
    if ($opLog.IsEnabled) { $opText = '是' }
    Add-Line ('打印操作日志（PrintService/Operational）已启用: ' + $opText)
    if (-not $opLog.IsEnabled) {
        Add-Flag 'INFO' '打印操作日志未开启，发生问题时无法回溯“谁在何时打印了什么”'
    }
}
catch {
    Add-Line '无法查询打印操作日志状态'
}
Add-Blank

# ---------- 10. verdict ----------
Add-Line '===== [10] 结论与建议 ====='
if ($Flags.Count -eq 0) {
    Add-Line '未发现明显异常。'
}
else {
    Add-Line '发现以下问题:'
    foreach ($f in $Flags) { Add-Line ('  ' + $f) }
}
Add-Blank
Add-Line '建议动作:'
Add-Line '  1. 若存在“端口/驱动错配”，删除该打印机并改用官方驱动重建（不要让它落在 IPP 类驱动上）'
Add-Line '  2. 若存在“同一台打印机注册了多个条目”，打开 设置 → 蓝牙和其他设备 → 打印机和扫描仪，'
Add-Line '     把多余的那条删掉，同一台机器只留一条通道（USB 直连 或 网络，二选一）'
Add-Line '  3. 若存在“组件签名异常”，以管理员身份加 -Repair 运行本工具，会从组件存储还原并隔离外来文件'
Add-Line '  4. 若队列堆积，加 -ClearQueue 运行（会先备份到带时间戳的目录，不做删除）'
Add-Line '  5. 开启打印操作日志以便日后审计'
Add-Blank

# ---------- 11. repair (admin + explicit switch) ----------
if ($Repair -or $ClearQueue) {
    Add-Line '===== [11] 修复动作 ====='
    if (-not $IsAdmin) {
        Add-Line '已跳过：需要以管理员身份运行才会执行修复。'
    }
    else {
        if ($ClearQueue -and $qFiles.Count -gt 0) {
            $bkRoot = Join-Path (Split-Path -Parent $queueDir) ('PRINTERS_backup_' + (Get-Date).ToString('yyyyMMdd_HHmmss'))
            New-Item -ItemType Directory -Path $bkRoot -Force | Out-Null
            $moved = 0
            foreach ($qf in $qFiles) {
                try {
                    Move-Item -LiteralPath $qf.FullName -Destination $bkRoot -Force -ErrorAction Stop
                    $moved = $moved + 1
                }
                catch {
                    Add-Line ('  移动失败: ' + $qf.Name)
                }
            }
            Add-Line ('  队列文件已备份到: ' + $moved + ' -> ' + $bkRoot)
        }

        if ($Repair) {
            try {
                $lgBefore = Get-WinEvent -ListLog 'Microsoft-Windows-PrintService/Operational' -ErrorAction Stop
                if (-not $lgBefore.IsEnabled) {
                    & wevtutil sl 'Microsoft-Windows-PrintService/Operational' /e:true 2>&1 | Out-Null
                    Add-Line '  打印操作日志已开启'
                }
            }
            catch {
                Add-Line '  开启打印操作日志失败'
            }

            foreach ($sf in $suspectFiles) {
                $leaf = Split-Path -Path $sf -Leaf
                Add-Line ('  正在从组件存储还原: ' + $leaf)
                $candidates = @(Get-ChildItem 'C:\Windows\WinSxS' -Directory -ErrorAction SilentlyContinue |
                    Where-Object { (Test-Path (Join-Path $_.FullName $leaf)) -and ($_.Name -match 'printing') })
                if ($candidates.Count -eq 0) {
                    Add-Line '    组件存储中未找到可用源文件，已跳过'
                    continue
                }
                $src = Join-Path ($candidates | Sort-Object Name | Select-Object -Last 1).FullName $leaf
                & net stop spooler 2>&1 | Out-Null
                Start-Sleep -Seconds 3
                & takeown /F $sf /A 2>&1 | Out-Null
                & icacls $sf /grant 'Administrators:F' /C 2>&1 | Out-Null
                $qDir = Join-Path $OutDir ('quarantine_' + (Get-Date).ToString('yyyyMMdd_HHmmss'))
                New-Item -ItemType Directory -Path $qDir -Force | Out-Null
                $aside = Join-Path $qDir ($leaf + '.foreign')
                $ok = $false
                try {
                    Move-Item -LiteralPath $sf -Destination $aside -Force -ErrorAction Stop
                    $ok = $true
                }
                catch {
                    Add-Line ('    移动失败: ' + $_.Exception.Message)
                }
                if ($ok) {
                    Copy-Item -LiteralPath $src -Destination $sf -Force -ErrorAction SilentlyContinue
                    $s2 = Get-AuthenticodeSignature -FilePath $sf -ErrorAction SilentlyContinue
                    Add-Line ('    已还原，当前签名: ' + (Get-SigText -Text ([string]$s2.Status)))
                    Add-Line ('    原文件已保留在: ' + $aside)
                }
                & net start spooler 2>&1 | Out-Null
                Start-Sleep -Seconds 4
                $after = Get-Service Spooler -ErrorAction SilentlyContinue
                Add-Line ('    服务状态: ' + (Get-SvcText -Text ([string]$after.Status)))
            }
        }
    }
    Add-Blank
}

# ---------- write report ----------
$stamp = $StartTime.ToString('yyyyMMdd_HHmmss')
$reportPath = Join-Path $OutDir ('WinPrintDiag_' + $stamp + '.txt')
[System.IO.File]::WriteAllLines($reportPath, $Report, (New-Object System.Text.UTF8Encoding($true)))

Write-Output ''
Write-Output ('WinPrintDiag ' + $ToolVersion + ' 体检完成。')
Write-Output ('报告文件: ' + $reportPath)
Write-Output ('发现项   : ' + $Flags.Count)
foreach ($f in $Flags) { Write-Output ('   ' + $f) }
Write-Output ''
if ($OpenReport) {
    & notepad $reportPath
}
