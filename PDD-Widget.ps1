# =====================================================================
#  PDD Stock Float Widget
#  Interaction reference: Moyu-Reader style floating mini-window.
#  Windows PowerShell 5.1 + WinForms, no compile required.
#
#  Features:
#   - Frameless rounded always-on-top translucent card
#   - Auto-fade to hidden when mouse leaves; reappears on hover
#   - Stepless opacity: mouse wheel on hover, or slider in right-click
#     menu -> Settings
#   - Quote fetched in a separate child process (UI never blocks);
#     multi-source fallback eastmoney -> tencent -> sina(optional)
#   - US trading session state (ET) incl. NYSE holidays (approx)
#   - Double-click opens Eastmoney quote page
#   - Left-drag moves; right-click opens context menu
#
#  Usage:
#    powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File PDD-Widget.ps1 [-Demo]
#    -Demo : built-in demo quote (offline, for UI preview)
# =====================================================================

param([switch]$Demo, [switch]$NoRun)

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---- DPI --------------------------------------------------------------
# EXE 由 csc 生成, 没有 dpiAware 清单, 进程默认是 DPI 非感知: 125%/150% 缩放下
# 整块卡片被位图拉伸, 文字发虚。这里在创建任何窗口前声明 System DPI 感知,
# 再按实际 DPI 缩放几何尺寸(字号用磅, 本身随 DPI 放大, 比例自然一致)。
$script:Scale = 1.0
function Enable-DpiAwareness {
    $ok = $false
    try {
        if (-not ('DpiUtil' -as [type])) {
            $sig = 'using System;using System.Runtime.InteropServices;' +
                   'public static class DpiUtil{' +
                   '[DllImport("user32.dll")]public static extern bool SetProcessDpiAwarenessContext(IntPtr v);' +
                   '[DllImport("user32.dll")]public static extern uint GetDpiForSystem();' +
                   '[DllImport("user32.dll")]public static extern bool SetProcessDPIAware();}'
            Add-Type -TypeDefinition $sig -ErrorAction Stop | Out-Null
        }
    } catch {}
    try { $ok = [DpiUtil]::SetProcessDpiAwarenessContext([IntPtr](-2)) } catch {}   # -2 = SYSTEM_AWARE
    if (-not $ok) { try { $ok = [DpiUtil]::SetProcessDPIAware() } catch {} }
    # 没能接管 DPI 感知时保持 Scale=1: 让系统继续按位图拉伸, 避免"系统缩放 + 自己缩放"叠加
    if (-not $ok) { $script:Scale = 1.0; return }
    $dpi = 96
    try { $dpi = [int][DpiUtil]::GetDpiForSystem() } catch {}
    if ($dpi -le 0) { $dpi = 96 }
    $script:Scale = [math]::Round($dpi / 96.0, 3)
}
Enable-DpiAwareness

# ---- constants -------------------------------------------------------
$script:Ticker    = 'PDD'
$script:StockName   = '拼多多'
$script:FetchScript = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) 'PDD-Fetch.ps1'
$script:FetchTimeoutMs = 15000   # 常驻子进程下单轮最坏 ~12s(3 源 x 4s)
$script:DefaultSources = @('tencent', 'eastmoney', 'sina')

function Get-MarketKind([string]$code) {
    $c = $code.Trim().ToLower()
    if ($c -match '^(sh|sz|bj)?[0-9]{6}$') { return 'ASH' }
    if ($c -match '^hk[0-9]{4,5}$') { return 'HK' }
    if ($c -match '^[0-9]{4,5}$') { return 'HK' }
    return 'US'
}
function Get-AShareExchange([string]$num) {
    # 与取数脚本同一套规则: 6=沪A, 900=沪B, 92=北交所新代码段, 43/83/87/88=北交所, 其余=深市
    if ($num -match '^6') { return 'sh' }
    if ($num -match '^92') { return 'bj' }
    if ($num -match '^9') { return 'sh' }
    if ($num -match '^(43|83|87|88)') { return 'bj' }
    return 'sz'
}
function Get-QuotePageUrl {
    $raw = ([string]$script:Cfg.Ticker).Trim().ToUpperInvariant()
    if (-not $raw) { $raw = 'PDD' }
    $kind = Get-MarketKind $raw
    $c = $raw -replace '^(SH|SZ|BJ|HK|US)', ''
    if ($kind -eq 'ASH') {
        $ex = Get-AShareExchange $c
        return "https://quote.eastmoney.com/$ex$c.html"
    }
    if ($kind -eq 'HK') { return "https://quote.eastmoney.com/hk/$c.html" }
    return "https://quote.eastmoney.com/us/$c.html"
}
function Sync-TickerFromCfg {
    $code = ([string]$script:Cfg.Ticker).Trim().ToUpperInvariant()
    if (-not $code) { $code = 'PDD'; $script:Cfg.Ticker = 'PDD' }
    $script:Ticker = $code
    $name = [string]$script:Cfg.StockName
    if (-not $name) { $name = $code }
    $script:StockName = $name
}
function Get-StockPresets {
    # 设置对话框与右键菜单共用同一份预设, 避免两处不一致
    return [ordered]@{
        'PDD' = '拼多多(美股)'; 'BABA' = '阿里巴巴(美股)'; 'JD' = '京东(美股)'; 'NVDA' = '英伟达(美股)'
        'TSLA' = '特斯拉(美股)'; 'AAPL' = '苹果(美股)'; 'MSFT' = '微软(美股)'
        '600519' = '贵州茅台(A股)'; '000001' = '平安银行(A股)'; '300750' = '宁德时代(A股)'
        '002594' = '比亚迪(A股)'; '600036' = '招商银行(A股)'; '688981' = '中芯国际(A股)'
        '00700' = '腾讯控股(港股)'; '09988' = '阿里巴巴-W(港股)'; '01810' = '小米集团(港股)'
        '03690' = '美团(港股)'; '09618' = '京东集团(港股)'
    }
}
function Set-Stock([string]$code, [string]$name) {
    $code = ([string]$code).Trim().ToUpperInvariant()
    if ($code -notmatch '^[A-Za-z0-9.\-]{1,12}$') { return }
    $script:Cfg.Ticker = $code
    if (-not $name) { $name = $code }
    $script:Cfg.StockName = $name
    Sync-TickerFromCfg
    $script:SesCache = $null
    $script:Quote = $null
    $script:SourceUsed = ''
    $script:FetchErr = ''
    Save-Cfg
    if ($script:TmrRefresh) { $script:TmrRefresh.Interval = (Get-EffectiveRefreshSec) * 1000 }
    $script:PaintRev++
    if ($script:DemoOn) { $script:Quote = Get-DemoQuote } else { Start-Fetch }
    Show-Toast ('已切换: ' + $code + ' ' + $name)
}
function Get-AutostartPath { return (Join-Path ([Environment]::GetFolderPath('Startup')) '股价悬浮窗.lnk') }
function Test-Autostart { try { return (Test-Path (Get-AutostartPath)) } catch { return $false } }
function Set-Autostart([bool]$enable) {
    try {
        $lnk = Get-AutostartPath
        if (-not $enable) {
            if (Test-Path $lnk) { Remove-Item $lnk -Force }
            return $true
        }
        $wsh = New-Object -ComObject WScript.Shell
        $sc = $wsh.CreateShortcut($lnk)
        $exe = [string]$env:PDDWIDGET_EXE
        if ($exe -and (Test-Path $exe)) {
            $sc.TargetPath = $exe
            $sc.Arguments = ''
            $sc.WorkingDirectory = (Split-Path -Parent $exe)
        } else {
            # 脚本模式: 用 wscript 调 vbs 静默启动, 不弹黑窗
            $vbs = Join-Path $PSScriptRoot 'PDD-Widget.vbs'
            $sc.TargetPath = 'wscript.exe'
            $sc.Arguments = '"' + $vbs + '"'
            $sc.WorkingDirectory = $PSScriptRoot
        }
        $sc.Description = '股价悬浮窗 开机自启动'
        $sc.Save()
        return $true
    } catch { Trace-Msg ('autostart err: ' + $_.Exception.Message); return $false }
}

$script:ColorUp   = [System.Drawing.Color]::FromArgb(255, 236, 76, 71)   # up
$script:ColorDown = [System.Drawing.Color]::FromArgb(255, 20, 170, 90)   # down
$script:ColorFlat = [System.Drawing.Color]::FromArgb(255, 178, 184, 192)
$script:ThemeRiseGreen = $false

# ---- state -----------------------------------------------------------
$script:Cfg = @{
    X = -1; Y = -1
    Opacity    = 0.90
    HiddenOp   = 0.05
    AutoHide   = $true
    TopMost    = $true
    RefreshSec = 6
    BodyStyle  = 'glass'
    RiseGreen  = $false
    Ticker     = 'PDD'
    StockName  = '拼多多'
    Sources    = $script:DefaultSources
    SrcMigrated = $false      # 数据源顺序迁移只做一次, 之后尊重用户配置
}
$script:Quote = $null
$script:SourceUsed = ''
$script:FetchErr = ''
$script:Fetching = $false
$script:FetchProc = $null
$script:FetchTmp = ''
$script:FetchStart = [datetime]::MinValue
$script:FetchHost = $null        # 常驻取数子进程
$script:FetchHostIn = $null      # 其 stdin 写入器
$script:QuoteAt = [datetime]::MinValue
$script:PaintRev = 0             # 内容版本号: 只有变了才重画
$script:SeenRev = -1
$script:FadeRun = $false
$script:DemoOn = [bool]$Demo
$script:IsDragging = $false
$script:DragOffset = $null
$script:CtxOpen = $false
$script:FadeTarget = 0.9
$script:ToastText = ''
$script:ToastUntil = [datetime]::MinValue
$script:Form = $null
$script:TmrRefresh = $null
$script:TmrFade = $null
$script:TmrHover = $null
$script:TmrPoll = $null
$script:SesCache = $null; $script:SesKind = ''; $script:SesCacheAt = [datetime]::MinValue
$script:LastSavedQuote = $null; $script:LastSavedAt = [datetime]::MinValue
$script:CfgDirty = $false; $script:CfgDirtyAt = [datetime]::MinValue
$script:HolidayTable = $null

$script:FPrice = $null; $script:FPct = $null; $script:FName = $null
$script:FMeta = $null; $script:FTiny = $null
$script:BrText = $null; $script:BrDim = $null
$script:BrBadge = $null; $script:BrBadgeG = $null; $script:PenCard = $null

# ---- paths & config --------------------------------------------------
function Get-ConfigDir {
    $sd = $PSScriptRoot
    $candidates = @((Join-Path $sd 'config'), (Join-Path $env:APPDATA 'PDDStockWidget'))
    foreach ($p in $candidates) {
        try {
            if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
            $t = Join-Path $p ('.w' + [guid]::NewGuid().ToString('N'))
            [System.IO.File]::WriteAllText($t, 'x')
            Remove-Item $t -Force -ErrorAction SilentlyContinue
            return $p
        } catch { }
    }
    return $candidates[0]
}
$script:ConfigDir = Get-ConfigDir
$script:CfgPath = Join-Path $script:ConfigDir 'config.json'
$script:LastPath = Join-Path $script:ConfigDir 'lastQuote.json'

function Save-Cfg {
    try { $script:Cfg | ConvertTo-Json -Depth 5 | Set-Content -Path $script:CfgPath -Encoding UTF8 } catch {}
}
function Load-Cfg {
    try {
        if (Test-Path $script:CfgPath) {
            $j = Get-Content $script:CfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
            # 只把 JSON 中且存在于默认配置的键写入, 避免枚举时修改哈希表
            foreach ($prop in $j.PSObject.Properties) {
                if ($script:Cfg.Contains($prop.Name)) { $script:Cfg[$prop.Name] = $prop.Value }
            }
        }
    } catch {}
}
function Load-LastQuote {
    try {
        if (Test-Path $script:LastPath) {
            $j = Get-Content $script:LastPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $cmpA = ([string]$j.Code).ToUpperInvariant() -replace '^(SH|SZ|BJ|HK|US|\.)', ''
            $cmpB = $script:Ticker -replace '^(SH|SZ|BJ|HK|US)', ''
            if ($j -and $j.Price -gt 0 -and ($cmpA.StartsWith($cmpB) -or $cmpB.StartsWith($cmpA))) {
                $j | Add-Member -NotePropertyName Stale -NotePropertyValue $true -Force
                return $j
            }
        }
    } catch {}
    return $null
}
function Save-LastQuote($q) {
    # 节流: 代码变化 / 涨跌幅变化>=0.1 个百分点 / 超过 120 秒才落盘 (原先每 6 秒写一次, 约 1.4 万次/天)
    try {
        $prev = $script:LastSavedQuote
        $codeChanged = (-not $prev) -or ([string]$q.Code -ne [string]$prev.Code)
        $moved = (-not $prev) -or ([math]::Abs([double]$q.Pct - [double]$prev.Pct) -ge 0.1)
        $due = (-not $script:LastSavedAt) -or $script:LastSavedAt -eq [datetime]::MinValue -or ((Get-Date) - $script:LastSavedAt).TotalSeconds -ge 120
        if ($codeChanged -or $moved -or $due) {
            $q | ConvertTo-Json -Depth 4 | Set-Content -Path $script:LastPath -Encoding UTF8
            $script:LastSavedQuote = $q
            $script:LastSavedAt = Get-Date
        }
    } catch {}
}
$script:TracePath = Join-Path $script:ConfigDir 'trace.log'
function Trace-Msg($m) {
    try { Add-Content -Encoding UTF8 -Path $script:TracePath -Value ("[{0:HH:mm:ss.fff}] {1}" -f (Get-Date), $m) } catch {}
}

# ---- US session (ET, incl NYSE holidays approx) ----------------------
function Get-EasterSunday([int]$y) {
    $a = $y % 19; $b = [math]::Floor($y / 100); $c = $y % 100
    $d = [math]::Floor($b / 4); $e = $b % 4
    $f = [math]::Floor(($b + 8) / 25); $g = [math]::Floor(($b - $f + 1) / 3)
    $h = (19 * $a + $b - $d - $g + 15) % 30
    $i = [math]::Floor($c / 4); $k = $c % 4
    $l = (32 + 2 * $e + 2 * $i - $h - $k) % 7
    $m = [math]::Floor(($a + 11 * $h + 22 * $l) / 451)
    $month = [math]::Floor(($h + $l - 7 * $m + 114) / 31)
    $day = (($h + $l - 7 * $m + 114) % 31) + 1
    return (Get-Date -Year $y -Month $month -Day $day).Date
}
function Get-NthWeekday([datetime]$d, [int]$week, [System.DayOfWeek]$dow) {
    $first = Get-Date -Year $d.Year -Month $d.Month -Day 1
    $diff = ([int]$dow - [int]$first.DayOfWeek + 7) % 7
    return $first.AddDays($diff + 7 * ($week - 1))
}
function Get-LastWeekday([datetime]$d, [System.DayOfWeek]$dow) {
    $last = (Get-Date -Year $d.Year -Month $d.Month -Day 1).AddMonths(1).AddDays(-1)
    $diff = ([int]$last.DayOfWeek - [int]$dow + 7) % 7
    return $last.AddDays(-$diff)
}
function Get-ObservedDate([datetime]$d) {
    if ($d.DayOfWeek -eq [DayOfWeek]::Saturday) { return $d.AddDays(-1) }
    if ($d.DayOfWeek -eq [DayOfWeek]::Sunday)   { return $d.AddDays(1) }
    return $d
}
function Test-NyseHoliday([datetime]$d) {
    $d = $d.Date; $y = $d.Year
    $list = New-Object System.Collections.Generic.List[datetime]
    $list.Add((Get-Date -Year $y -Month 1 -Day 1).Date)
    $list.Add((Get-NthWeekday (Get-Date -Year $y -Month 1 -Day 1) 3 ([DayOfWeek]::Monday)))
    $list.Add((Get-NthWeekday (Get-Date -Year $y -Month 2 -Day 1) 3 ([DayOfWeek]::Monday)))
    $list.Add((Get-EasterSunday $y).AddDays(-2))
    $list.Add((Get-LastWeekday (Get-Date -Year $y -Month 5 -Day 1) ([DayOfWeek]::Monday)))
    $list.Add((Get-ObservedDate (Get-Date -Year $y -Month 6 -Day 19)))
    $list.Add((Get-ObservedDate (Get-Date -Year $y -Month 7 -Day 4)))
    $list.Add((Get-NthWeekday (Get-Date -Year $y -Month 9 -Day 1) 1 ([DayOfWeek]::Monday)))
    $list.Add((Get-NthWeekday (Get-Date -Year $y -Month 11 -Day 1) 4 ([DayOfWeek]::Thursday)))
    $list.Add((Get-ObservedDate (Get-Date -Year $y -Month 12 -Day 25)))
    return $list.Contains($d)
}
function Get-MarketHolidayTable {
    if ($script:HolidayTable) { return $script:HolidayTable }
    # 2026 数据来源:
    #   A股  国务院办公厅 国办发明电〔2025〕7号 (2026年部分节假日安排)
    #   港股 GovHK "General holidays for 2026"
    # 每年需更新; 也可在 config 目录放 holidays.json 覆盖, 格式: {"ASH":["2026-01-01"],"HK":["2026-01-01"]}
    $t = @{
        'ASH' = @(
            '2026-01-01','2026-01-02',
            '2026-02-16','2026-02-17','2026-02-18','2026-02-19','2026-02-20','2026-02-23',
            '2026-04-06',
            '2026-05-01','2026-05-04','2026-05-05',
            '2026-06-19',
            '2026-09-25',
            '2026-10-01','2026-10-02','2026-10-05','2026-10-06','2026-10-07'
        )
        'HK' = @(
            '2026-01-01','2026-02-17','2026-02-18','2026-02-19',
            '2026-04-03','2026-04-06','2026-04-07',
            '2026-05-01','2026-05-25','2026-06-19','2026-07-01',
            '2026-10-01','2026-10-19','2026-12-25'
        )
    }
    try {
        $json = Join-Path $script:ConfigDir 'holidays.json'
        if (Test-Path $json) {
            $j = Get-Content $json -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($k in @('ASH', 'HK')) {
                $p = $j.PSObject.Properties[$k]
                if ($p -and @($p.Value).Count -gt 0) { $t[$k] = @($p.Value) }
            }
        }
    } catch {}
    $script:HolidayTable = $t
    return $t
}
function Test-MarketHoliday([string]$kind, [datetime]$d) {
    if ($kind -eq 'US') { return (Test-NyseHoliday $d) }
    try {
        $tbl = Get-MarketHolidayTable
        if (-not $tbl.ContainsKey($kind)) { return $false }
        return ($tbl[$kind] -contains $d.ToString('yyyy-MM-dd'))
    } catch { return $false }
}
function Test-HolidayTableFresh([string]$kind, [int]$year) {
    # 表里若没有今年任何一条, 说明跨年了: 不能把"查不到"当成"不是假期"
    try {
        $tbl = Get-MarketHolidayTable
        if (-not $tbl.ContainsKey($kind)) { return $false }
        foreach ($d in @($tbl[$kind])) { if ([string]$d -like "$year-*") { return $true } }
    } catch {}
    return $false
}
function Get-NextMarketHoliday([string]$kind) {
    try {
        $tbl = Get-MarketHolidayTable
        if (-not $tbl.ContainsKey($kind)) { return '' }
        $today = (Get-Date).ToString('yyyy-MM-dd')
        $next = @($tbl[$kind] | Where-Object { $_ -ge $today } | Sort-Object)
        if ($next.Count -gt 0) { return [string]$next[0] }
    } catch {}
    return ''
}
function Get-TradingSession([string]$kind) {
    # 结果缓存 20 秒: 绘制时每帧要调用两次, 原先每次都重建假日表
    $now = Get-Date
    if ($script:SesCache -and $script:SesKind -eq $kind -and ($now - $script:SesCacheAt).TotalSeconds -lt 20) { return $script:SesCache }
    $labelMap = @{ PRE = '盘前'; REGULAR = '交易中'; BREAK = '午间休市'; AFTER = '盘后'; CLOSED = '休市' }
    $mktLabel = @{ US = '美股'; ASH = 'A股'; HK = '港股' }[$kind]
    if ($kind -eq 'US') {
        $tz = [TimeZoneInfo]::FindSystemTimeZoneById('Eastern Standard Time')
        $lt = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
    } else {
        # A股/港股均为 UTC+8 且无夏令时
        $lt = [DateTime]::UtcNow.AddHours(8)
    }
    $res = @{ Key = 'CLOSED'; Label = '休市'; Local = $lt; Mkt = $mktLabel; Minutes = ($lt.Hour * 60 + $lt.Minute); Note = '' }
    # 内置假期表只覆盖到 2026: 跨年后不能把"查不到"当成"不是假期", 而是标注未覆盖
    if ($kind -ne 'US' -and -not (Test-HolidayTableFresh $kind $lt.Year)) {
        $res.Note = '假期表未覆盖 ' + $lt.Year + ', 仅按时间判断'
    }
    $weekend = ($lt.DayOfWeek -eq [DayOfWeek]::Saturday -or $lt.DayOfWeek -eq [DayOfWeek]::Sunday)
    if ($weekend -or (Test-MarketHoliday $kind $lt)) {
        $script:SesCache = $res; $script:SesKind = $kind; $script:SesCacheAt = $now
        return $res
    }
    if ($kind -eq 'US') {
        if ($res.Minutes -ge 240 -and $res.Minutes -lt 570)  { $res.Key = 'PRE' }
        elseif ($res.Minutes -ge 570 -and $res.Minutes -lt 960)  { $res.Key = 'REGULAR' }
        elseif ($res.Minutes -ge 960 -and $res.Minutes -lt 1200) { $res.Key = 'AFTER' }
    } elseif ($kind -eq 'ASH') {
        # 9:15-9:25 集合竞价 / 9:30-11:30 / 午休 / 13:00-15:00 (北京时间, 含 2026 法定节假日)
        if ($res.Minutes -ge 555 -and $res.Minutes -lt 570) { $res.Key = 'PRE' }
        elseif (($res.Minutes -ge 570 -and $res.Minutes -lt 690) -or ($res.Minutes -ge 780 -and $res.Minutes -lt 900)) { $res.Key = 'REGULAR' }
        elseif ($res.Minutes -ge 690 -and $res.Minutes -lt 780) { $res.Key = 'BREAK' }
    } else {
        # 港股 9:00-9:30 竞价 / 9:30-12:00 / 午休 / 13:00-16:00
        if ($res.Minutes -ge 540 -and $res.Minutes -lt 570) { $res.Key = 'PRE' }
        elseif (($res.Minutes -ge 570 -and $res.Minutes -lt 720) -or ($res.Minutes -ge 780 -and $res.Minutes -lt 960)) { $res.Key = 'REGULAR' }
        elseif ($res.Minutes -ge 720 -and $res.Minutes -lt 780) { $res.Key = 'BREAK' }
    }
    $res.Label = $labelMap[$res.Key]
    $script:SesCache = $res; $script:SesKind = $kind; $script:SesCacheAt = $now
    return $res
}
function Get-EffectiveRefreshSec {
    # 休市/盘前盘后自动降频, 避免夜间与周末无意义请求
    $base = [int]$script:Cfg.RefreshSec
    if ($base -lt 3) { $base = 3 }
    if ($script:DemoOn) { return $base }
    try {
        $ses = Get-TradingSession (Get-MarketKind $script:Ticker)
        switch ($ses.Key) {
            'REGULAR' { return $base }
            'PRE'     { return [math]::Max($base, 20) }
            'AFTER'   { return [math]::Max($base, 30) }
            default   { return [math]::Max($base, 60) }
        }
    } catch { return $base }
}

# ---- formatting ------------------------------------------------------
# 一律按不变文化格式化: 否则德/法区域会把小数点变成逗号, 显示成 "1435,26"
$script:InvCI = [System.Globalization.CultureInfo]::InvariantCulture
function Format-Vol([double]$v) {
    if ($v -ge 1e8) { return [string]::Format($script:InvCI, '{0:N2}亿', ($v / 1e8)) }
    if ($v -ge 1e4) { return [string]::Format($script:InvCI, '{0:N2}万', ($v / 1e4)) }
    return [string]::Format($script:InvCI, '{0:N0}', $v)
}
function Format-Amt([double]$v, [string]$cur) {
    $unit = @{ USD = '美元'; CNY = '元'; HKD = '港元' }[$cur]
    if (-not $unit) { $unit = '美元' }
    if ($v -ge 1e8) { return ([string]::Format($script:InvCI, '{0:N2}亿', ($v / 1e8)) + $unit) }
    if ($v -ge 1e4) { return ([string]::Format($script:InvCI, '{0:N1}万', ($v / 1e4)) + $unit) }
    return ([string]::Format($script:InvCI, '{0:N0}', $v) + $unit)
}
function Format-Price([double]$v) { return [string]::Format($script:InvCI, '{0:N2}', $v) }
# 明细行空间紧, 不带千分位, 四个价位才放得下
function Format-PricePlain([double]$v) { return [string]::Format($script:InvCI, '{0:0.00}', $v) }

# ---- async fetch (resident child process) ----------------------------
# 原先每轮刷新都 Start-Process 一个新 powershell: 实测单次 0.9~1.3s, 其中 PS5.1
# 冷启动就占 0.4s, 6s 刷新 = 每天 1.4 万次进程创建。改为常驻子进程 + stdin 命令,
# 单轮只剩网络时间; 起不来时才退回一次性进程。
function Get-PowerShellExe {
    $p = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $p) { return $p }
    return 'powershell.exe'
}

function Stop-FetchHost {
    try {
        if ($script:FetchHostIn) { $script:FetchHostIn.WriteLine('QUIT'); $script:FetchHostIn.Flush() }
    } catch {}
    try { if ($script:FetchHostIn) { $script:FetchHostIn.Dispose() } } catch {}
    $script:FetchHostIn = $null
    try {
        if ($script:FetchHost -and -not $script:FetchHost.HasExited) {
            if (-not $script:FetchHost.WaitForExit(2000)) { $script:FetchHost.Kill() }
        }
    } catch {}
    try { if ($script:FetchHost) { $script:FetchHost.Dispose() } } catch {}
    $script:FetchHost = $null
}

function Start-FetchHost {
    if ($script:FetchHost -and $script:FetchHostIn -and -not $script:FetchHost.HasExited) { return $true }
    Stop-FetchHost
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = Get-PowerShellExe
        $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $script:FetchScript + '" -Serve'
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
        $psi.RedirectStandardInput = $true
        $proc = [System.Diagnostics.Process]::Start($psi)
        if (-not $proc) { return $false }
        $script:FetchHost = $proc
        $script:FetchHostIn = $proc.StandardInput
        return $true
    } catch {
        Trace-Msg ('fetch-host err: ' + $_.Exception.Message)
        return $false
    }
}

function Start-Fetch {
    if ($script:DemoOn -or $script:Fetching) { return }
    if ($script:FetchProc -and $script:FetchProc.HasExited) {
        try { $script:FetchProc.Dispose() } catch {}
        $script:FetchProc = $null
    }
    if ($script:FetchProc) { return }
    if ($script:FetchTmp -and (Test-Path -LiteralPath $script:FetchTmp)) {
        Remove-Item -LiteralPath $script:FetchTmp -Force -ErrorAction SilentlyContinue
    }
    $script:FetchTmp = Join-Path $script:ConfigDir ('fetch-' + [guid]::NewGuid().ToString('N') + '.json')
    # 白名单校验: 防止被手工改坏的 config.json 把参数注入到子进程
    $tk = ([string]$script:Ticker).Trim()
    if ($tk -notmatch '^[A-Za-z0-9.\-]{1,12}$') { $tk = 'PDD' }
    $srcs = @($script:Cfg.Sources | Where-Object { $_ -in @('tencent', 'eastmoney', 'sina') })
    if ($srcs.Count -eq 0) { $srcs = @('tencent') }
    # 上次成功的源排到最前: 掉线的源不会每轮都先吃一次超时
    if ($script:LastGoodSrc -and ($srcs -contains $script:LastGoodSrc)) {
        $srcs = @($script:LastGoodSrc) + @($srcs | Where-Object { $_ -ne $script:LastGoodSrc })
    }
    $sources = ($srcs -join ',')

    $script:FetchStart = Get-Date
    $script:Fetching = $true
    if ($script:TmrPoll) { $script:TmrPoll.Start() }

    if (Start-FetchHost) {
        try {
            $script:FetchHostIn.WriteLine($tk + '|' + $sources + '|' + $script:FetchTmp)
            $script:FetchHostIn.Flush()
            return
        } catch {
            Trace-Msg ('fetch-host write err: ' + $_.Exception.Message)
            Stop-FetchHost
        }
    }

    # 兜底: 常驻进程起不来时退回一次性进程
    $fs = $script:FetchScript
    $a = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $fs + '"'),
            '-Ticker', $tk, '-Sources', $sources, '-Out', ('"' + $script:FetchTmp + '"'))
    try {
        $script:FetchProc = Start-Process -FilePath (Get-PowerShellExe) -ArgumentList $a -WindowStyle Hidden -PassThru
    } catch {
        $script:FetchErr = '启动取数失败: ' + $_.Exception.Message
        $script:FetchProc = $null
        $script:Fetching = $false
        if ($script:TmrPoll) { try { $script:TmrPoll.Stop() } catch {} }
    }
}

function Try-Complete-Fetch {
    # 只有拿到完整 JSON 才算结束本轮: 子进程还在写就留给下一轮轮询,
    # 顺带把"读到半截文件 -> 解析失败"这条误报消掉
    if (-not $script:FetchTmp) { return $false }
    if (-not (Test-Path -LiteralPath $script:FetchTmp)) { return $false }
    $txt = ''
    try {
        $txt = [System.IO.File]::ReadAllText($script:FetchTmp, (New-Object System.Text.UTF8Encoding($false)))
    } catch { return $false }
    if (-not $txt -or -not $txt.Trim()) { return $false }
    $j = $null
    try { $j = $txt | ConvertFrom-Json } catch { return $false }
    if ($null -eq $j) { return $false }
    Remove-Item -LiteralPath $script:FetchTmp -Force -ErrorAction SilentlyContinue
    try {
        if ($j.ok) {
            $d = $j.q
            $q = [pscustomobject]@{
                Name = [string]$d.Name; Code = [string]$d.Code
                Price = [double]$d.Price; PrevClose = [double]$d.PrevClose
                Open = [double]$d.Open; High = [double]$d.High; Low = [double]$d.Low
                Volume = [double]$d.Volume; Amount = [double]$d.Amount
                QuoteTime = [string]$d.QuoteTime
                Chg = [double]$d.Chg; Pct = [double]$d.Pct; Up = [int]$d.Up; Stale = $false
                Currency = [string]$d.Currency; Market = [string]$d.Market
                Warn = [string]$d.Warn
            }
            $script:Quote = $q
            $script:QuoteAt = Get-Date
            $script:LastGoodSrc = [string]$j.src
            $sn = [string]$j.src
            $script:SourceUsed = switch ($sn) { 'eastmoney' { '东财' } 'tencent' { '腾讯' } 'sina' { '新浪' } default { $sn } }
            $script:FetchErr = ''
            Save-LastQuote $q
        } else {
            $script:FetchErr = if ($j.err) { [string]$j.err } else { '无数据' }
        }
    } catch {
        $script:FetchErr = '解析失败: ' + $_.Exception.Message
    }
    $script:PaintRev++
    return $true
}

# ---- drawing helpers -------------------------------------------------
function New-RoundedRegion([int]$w, [int]$h, [int]$r) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $r * 2
    $path.AddArc(0, 0, $d, $d, 180, 90)
    $path.AddArc($w - $d, 0, $d, $d, 270, 90)
    $path.AddArc($w - $d, $h - $d, $d, $d, 0, 90)
    $path.AddArc(0, $h - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    $reg = New-Object System.Drawing.Region($path)
    $path.Dispose()
    return $reg
}
function Set-DoubleBuffered($ctl) {
    try {
        $prop = $ctl.GetType().GetProperty('DoubleBuffered',
            [System.Reflection.BindingFlags]'Instance,NonPublic,Public')
        if ($prop) { $prop.SetValue($ctl, $true, $null) }
    } catch {}
}
function Update-ThemeColors { $script:ThemeRiseGreen = [bool]$script:Cfg.RiseGreen }
function Get-DirBrush([int]$up) {
    if ($script:ThemeRiseGreen) {
        if ($up -gt 0) { return $script:BrDown }
        if ($up -lt 0) { return $script:BrUp }
        return $script:BrFlat
    }
    if ($up -gt 0) { return $script:BrUp }
    if ($up -lt 0) { return $script:BrDown }
    return $script:BrFlat
}
function Initialize-Assets {
    # 字体用磅值(随 DPI 自动放大), 几何尺寸按 $script:Scale 缩放, 两者比例一致。
    # 单独成函数是为了让 tests\Render-Smoke 能离屏复用同一套资源。
    $script:FPrice = New-Object System.Drawing.Font('Segoe UI', 25, [System.Drawing.FontStyle]::Bold)
    $script:FPct   = New-Object System.Drawing.Font('Segoe UI', 12.5, [System.Drawing.FontStyle]::Bold)
    $script:FName  = New-Object System.Drawing.Font('Segoe UI', 10.5, [System.Drawing.FontStyle]::Bold)
    $script:FMeta  = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $script:FTiny  = New-Object System.Drawing.Font('Microsoft YaHei UI', 8)
    $script:BrText   = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 234, 238, 244))
    $script:BrDim    = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 146, 154, 164))
    $script:BrBadge  = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 47, 174, 133))
    $script:BrBadgeG = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 118, 128, 140))
    $script:BrUp     = New-Object System.Drawing.SolidBrush($script:ColorUp)
    $script:BrDown   = New-Object System.Drawing.SolidBrush($script:ColorDown)
    $script:BrFlat   = New-Object System.Drawing.SolidBrush($script:ColorFlat)
    $script:BrWarn   = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 232, 172, 74))
    $script:PenCard  = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(60, 255, 255, 255), 1)
}

function Show-Toast([string]$msg, [int]$ms = 1600) {
    $script:ToastText = $msg
    $script:ToastUntil = (Get-Date).AddMilliseconds($ms)
    $script:PaintRev++   # 由悬停定时器负责重画, 避免每一处都直接 Invalidate
}

# ---- paint -----------------------------------------------------------
function Draw-Card($g, [int]$w, [int]$h) {
    # w/h 由调用方传入(便于离屏渲染自检), 不再依赖全局 $Form
    if ($w -le 0 -or $h -le 0) { return }
    $s = $script:Scale
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    if ($script:Cfg.BodyStyle -ne 'text') {
        $rect = New-Object System.Drawing.Rectangle(1, 1, ($w - 2), ($h - 2))
        $grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect,
            [System.Drawing.Color]::FromArgb(255, 26, 29, 37),
            [System.Drawing.Color]::FromArgb(255, 12, 14, 19), 90.0)
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $r = [int](16 * $s)
        $path.AddArc(1, 1, ($r * 2), ($r * 2), 180, 90)
        $path.AddArc(($w - 1 - $r * 2), 1, ($r * 2), ($r * 2), 270, 90)
        $path.AddArc(($w - 1 - $r * 2), ($h - 1 - $r * 2), ($r * 2), ($r * 2), 0, 90)
        $path.AddArc(1, ($h - 1 - $r * 2), ($r * 2), ($r * 2), 90, 90)
        $path.CloseFigure()
        $g.FillPath($grad, $path)
        $g.DrawPath($script:PenCard, $path)
        $path.Dispose(); $grad.Dispose()
    }

    $x0 = [int](16 * $s)
    $q = $script:Quote
    $dirBrush = $script:BrFlat
    if ($q) { $dirBrush = Get-DirBrush $q.Up }

    # badge 几何先算: 标题以它为右边界, 长名字用省略号截断, 不会压到徽章上
    $bw = [int](66 * $s); $bh = [int](20 * $s); $bx = $w - [int](16 * $s) - $bw; $by = [int](9 * $s)

    # header
    $name = "$($script:Ticker)  $($script:StockName)"
    if ($q -and $q.Name) { $name = "$($q.Code)  $($q.Name)" }
    $sfH = New-Object System.Drawing.StringFormat
    $sfH.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
    $sfH.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
    $g.DrawString($name, $script:FName, $script:BrText,
        (New-Object System.Drawing.RectangleF($x0, [int](10 * $s), [float]($bx - $x0 - [int](6 * $s)), [float](22 * $s))), $sfH)
    $sfH.Dispose()

    $st = Get-TradingSession (Get-MarketKind $script:Ticker)
    $badge = if ($script:DemoOn) { '演示' } else { $st.Label }
    $badgeCol = $script:BrBadge
    if ($st.Key -eq 'PRE' -or $st.Key -eq 'AFTER') {
        $badgeCol = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 226, 160, 58))
    } elseif ($st.Key -eq 'CLOSED' -or $st.Key -eq 'BREAK') { $badgeCol = $script:BrBadgeG }
    $bpath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $arc = [int](10 * $s)
    $bpath.AddArc($bx, $by, $arc, $arc, 180, 90); $bpath.AddArc(($bx + $bw - $arc), $by, $arc, $arc, 270, 90)
    $bpath.AddArc(($bx + $bw - $arc), ($by + $bh - $arc), $arc, $arc, 0, 90); $bpath.AddArc($bx, ($by + $bh - $arc), $arc, $arc, 90, 90)
    $bpath.CloseFigure()
    $g.FillPath($badgeCol, $bpath)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString($badge, $script:FTiny, [System.Drawing.Brushes]::White,
        (New-Object System.Drawing.RectangleF($bx, ($by - 1), $bw, $bh)), $sf)
    $sf.Dispose(); $bpath.Dispose()
    if ($badgeCol -ne $script:BrBadge -and $badgeCol -ne $script:BrBadgeG) { $badgeCol.Dispose() }

    # price
    if ($q) {
        # 启动时先显示的缓存(Stale)用灰字, 拿到实时数据再恢复涨跌色
        $priceBrush = if ($q.Stale) { $script:BrDim } else { $dirBrush }
        $priceTxt = Format-Price $q.Price
        $g.DrawString($priceTxt, $script:FPrice, $priceBrush, $x0, [int](40 * $s))
        $chgTxt = if ($q.Chg -ge 0) { '+' + (Format-Price $q.Chg) } else { (Format-Price $q.Chg) }
        $pctTxt = if ($q.Pct -ge 0) { '+' + [string]::Format($script:InvCI, '{0:N2}%', $q.Pct) } else { [string]::Format($script:InvCI, '{0:N2}%', $q.Pct) }
        $arrow = if ($q.Up -gt 0) { '▲ ' } elseif ($q.Up -lt 0) { '▼ ' } else { '— ' }
        # 涨跌行按实测价格宽度定位: 价格位数变化/高 DPI 下都不会压到价格上
        $pw = $g.MeasureString($priceTxt, $script:FPrice).Width
        $px = $x0 + [int]$pw + [int](4 * $s)
        $g.DrawString(("$arrow$chgTxt  $pctTxt"), $script:FPct, $priceBrush, $px, [int](47 * $s))
    } else {
        # 不再固定写"网络不可达": 代码不存在/网络超时是不同原因, 具体原因见下面的 err 行
        $msg = if ($script:Fetching) { '获取行情中…' } elseif ($script:FetchErr) { '获取失败' } else { '等待数据…' }
        $g.DrawString($msg, $script:FPrice, $script:BrDim, $x0, [int](42 * $s))
    }

    # detail lines
    $meta1 = ''; $meta2 = ''
    $metaW = [float]($w - ($x0 * 2))
    if ($q) {
        $o = if ($q.Open -gt 0) { Format-PricePlain $q.Open } else { '-' }
        $p = if ($q.PrevClose -gt 0) { Format-PricePlain $q.PrevClose } else { '-' }
        $hi = if ($q.High -gt 0) { Format-PricePlain $q.High } else { '-' }
        $lo = if ($q.Low -gt 0) { Format-PricePlain $q.Low } else { '-' }
        $meta1 = "昨收 $p 今开 $o 最高 $hi 最低 $lo"
        $volTxt = if ($q.Volume -gt 0) { Format-Vol $q.Volume } else { '-' }
        $amtTxt = if ($q.Amount -gt 0) { Format-Amt $q.Amount ([string]$q.Currency) } else { '-' }
        $meta2 = "量 $volTxt   额 $amtTxt"
        if ($q.QuoteTime) { $meta2 += "   " + $q.QuoteTime }
    }
    # 任意一行太长时按省略号截断, 不再溢出卡片
    $sfM = New-Object System.Drawing.StringFormat
    $sfM.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
    $sfM.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
    $g.DrawString($meta1, $script:FTiny, $script:BrDim,
        (New-Object System.Drawing.RectangleF($x0, [int](97 * $s), $metaW, [float](16 * $s))), $sfM)
    if (-not $q -and $script:FetchErr) {
        $e = $script:FetchErr
        if ($e.Length -gt 40) { $e = $e.Substring(0, 40) + '...' }
        $g.DrawString($e, $script:FTiny, $script:BrWarn,
            (New-Object System.Drawing.RectangleF($x0, [int](118 * $s), $metaW, [float](16 * $s))), $sfM)
    } elseif ($q -and $q.Warn) {
        # 数据源给的异常提示(如北交所老代码)放在量额行, 不再和 footer 的重叠
        $g.DrawString([string]$q.Warn, $script:FTiny, $script:BrWarn,
            (New-Object System.Drawing.RectangleF($x0, [int](118 * $s), $metaW, [float](16 * $s))), $sfM)
    } else {
        $g.DrawString($meta2, $script:FTiny, $script:BrDim,
            (New-Object System.Drawing.RectangleF($x0, [int](118 * $s), $metaW, [float](16 * $s))), $sfM)
    }
    $sfM.Dispose()

    # footer: 先照顾异常/提醒, 其次才是数据源与时段
    $foot = ''; $footBrush = $script:BrDim
    if ($script:DemoOn) {
        $foot = '演示模式'
    } else {
        $ses = Get-TradingSession (Get-MarketKind $script:Ticker)
        $stTxt = $ses.Mkt + ' ' + $ses.Local.ToString('HH:mm') + '  ' + $ses.Label
        if ($q -and -not $script:Fetching -and $script:FetchErr) {
            $foot = $script:FetchErr; $footBrush = $script:BrWarn
        } elseif (-not $q -and $ses.Note) {
            $foot = $ses.Note
        } elseif ($q) {
            $foot = "数据源 $($script:SourceUsed) · $stTxt"
            if ($script:QuoteAt -ne [datetime]::MinValue) {
                $age = [int]((Get-Date) - $script:QuoteAt).TotalSeconds
                if ($age -ge 60) { $foot += ' · ' + [string]([int]($age / 60)) + ' 分钟前' }
            }
        } elseif ($script:Fetching) { $foot = "刷新中… · $stTxt" }
        else { $foot = $stTxt }
    }
    $sf2 = New-Object System.Drawing.StringFormat
    $sf2.Alignment = [System.Drawing.StringAlignment]::Far
    $sf2.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
    $sf2.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
    $g.DrawString($foot, $script:FTiny, $footBrush,
        (New-Object System.Drawing.RectangleF($x0, [int](132 * $s), $metaW, [float](16 * $s))), $sf2)
    $sf2.Dispose()

    # toast: 不透明底, 否则会透出下面那行字
    if ($script:ToastText -and (Get-Date) -lt $script:ToastUntil) {
        $sz = $g.MeasureString($script:ToastText, $script:FTiny)
        $tw = [int]([math]::Ceiling($sz.Width) + (24 * $s)); $th = [int]([math]::Ceiling($sz.Height) + (8 * $s))
        $tx = [int](($w - $tw) / 2); if ($tx -lt 2) { $tx = 2 }
        $ty = $h - $th - [int](4 * $s)
        $tb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(250, 8, 10, 14))
        $g.FillRectangle($tb, $tx, $ty, $tw, $th)
        $g.DrawRectangle($script:PenCard, $tx, $ty, $tw, $th)
        $g.DrawString($script:ToastText, $script:FTiny, $script:BrText, ($tx + [int](12 * $s)), ($ty + [int](4 * $s)))
        $tb.Dispose()
    }
}

# ---- input events ----------------------------------------------------
function On-MouseDown($e) {
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        $script:IsDragging = $true
        $script:DragOffset = New-Object System.Drawing.Point($e.X, $e.Y)
    }
}
function On-MouseMove($e) {
    if ($script:IsDragging -and $script:DragOffset) {
        $pt = New-Object System.Drawing.Point($e.X, $e.Y)
        $np = $script:Form.PointToScreen($pt)
        $script:Form.Location = New-Object System.Drawing.Point(
            ($np.X - $script:DragOffset.X), ($np.Y - $script:DragOffset.Y))
    }
}
function On-MouseUp($e) {
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        if ($script:IsDragging) {
            # 拖动结束即记录位置(防抖保存), 避免被强杀时丢失
            try {
                $script:Cfg.X = $script:Form.Location.X
                $script:Cfg.Y = $script:Form.Location.Y
                $script:CfgDirty = $true; $script:CfgDirtyAt = Get-Date
            } catch {}
        }
        $script:IsDragging = $false; $script:DragOffset = $null
    }
}
function On-DoubleClick($e) {
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) { Start-Process (Get-QuotePageUrl) }
}
function On-Wheel($e) {
    $delta = if ($e.Delta -gt 0) { 0.05 } else { -0.05 }
    $script:Cfg.Opacity = [math]::Round([math]::Min(1.0, [math]::Max(0.2, [double]$script:Cfg.Opacity + $delta)), 2)
    $script:FadeTarget = [double]$script:Cfg.Opacity
    if (-not $script:FadeRun -and $script:TmrFade) { $script:TmrFade.Start(); $script:FadeRun = $true }
    # 防抖: 滚动过程中不写盘, 停手约 0.7 秒后统一保存
    $script:CfgDirty = $true; $script:CfgDirtyAt = Get-Date
    Show-Toast ('透明度 {0}%' -f [int]($script:Cfg.Opacity * 100))
    $e.Handled = $true
}

# ---- build form ------------------------------------------------------
function Get-DemoQuote {
    # 演示数据按市场给中性数值与币种, 避免选了A股/港股却显示美元与美股价格
    $kind = Get-MarketKind $script:Ticker
    $cur = @{ US = 'USD'; ASH = 'CNY'; HK = 'HKD' }[$kind]
    if (-not $cur) { $cur = 'USD' }
    $prev = @{ US = 82.21; ASH = 20.00; HK = 100.00 }[$kind]
    if (-not $prev) { $prev = 50.00 }
    $p = [math]::Round($prev * 0.9701, 2)
    return [pscustomobject]@{
        Name = ([string]$script:Cfg.StockName); Code = $script:Ticker; Price = $p; PrevClose = $prev
        Open = [math]::Round($prev * 0.988, 2); High = [math]::Round($prev * 1.004, 2); Low = [math]::Round($prev * 0.965, 2)
        Volume = 12345678; Amount = 987654321; QuoteTime = '演示数据'
        Chg = ($p - $prev); Pct = (($p / $prev - 1) * 100); Up = -1; Stale = $false
        Currency = $cur; Market = $kind
    }
}

function Build-Form {
    $s = $script:Scale
    $W = [int](350 * $s); $H = [int](152 * $s); $R = [int](16 * $s)
    $cfg = $script:Cfg

    $form = New-Object System.Windows.Forms.Form
    $form.Text = '股价悬浮窗'
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $form.ShowInTaskbar = $false
    $form.TopMost = [bool]$cfg.TopMost
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
    $form.BackColor = [System.Drawing.Color]::FromArgb(255, 10, 12, 16)
    $form.ClientSize = New-Object System.Drawing.Size($W, $H)
    Set-DoubleBuffered $form
    $form.Region = New-RoundedRegion $W $H $R

    # 位置校验: 顶部中间锚点必须落在某个屏幕工作区内, 否则回退到主屏右下角
    # (原先只比 PrimaryScreen 的 Width/Height, 副屏上的合法位置会被误判为越界)
    $x = [int]$cfg.X; $y = [int]$cfg.Y
    $needDefault = ($x -lt 0 -or $y -lt 0)
    if (-not $needDefault) {
        $anchor = New-Object System.Drawing.Point(($x + [int]($W / 2)), ($y + 12))
        $onScreen = $false
        foreach ($sc in [System.Windows.Forms.Screen]::AllScreens) {
            if ($sc.WorkingArea.Contains($anchor)) { $onScreen = $true; break }
        }
        if (-not $onScreen) { $needDefault = $true }
    }
    if ($needDefault) {
        $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $x = $wa.Right - $W - 24; $y = $wa.Bottom - $H - 48
    }
    $form.Location = New-Object System.Drawing.Point($x, $y)

    Initialize-Assets
    $script:PaintRev++

    $form.add_Paint({ param($s2, $e)
        try { Draw-Card $e.Graphics $form.ClientSize.Width $form.ClientSize.Height }
        catch { Trace-Msg ('PAINT ERR: ' + $_.Exception.Message) } })
    $form.add_MouseDown({ param($s, $e) On-MouseDown $e })
    $form.add_MouseMove({ param($s, $e) On-MouseMove $e })
    $form.add_MouseUp({ param($s, $e) On-MouseUp $e })
    $form.add_MouseDoubleClick({ param($s, $e) On-DoubleClick $e })
    $form.add_MouseWheel({ param($s, $e) On-Wheel $e })

    # context menu
    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    $menu.ShowImageMargin = $false
    $itRefresh = New-Object System.Windows.Forms.ToolStripMenuItem('立即刷新')
    $itOpen    = New-Object System.Windows.Forms.ToolStripMenuItem('打开东财行情页')
    $itStock   = New-Object System.Windows.Forms.ToolStripMenuItem('切换股票')
    $presetsMenu = Get-StockPresets
    foreach ($pc in $presetsMenu.Keys) {
        $mi = New-Object System.Windows.Forms.ToolStripMenuItem([string]($pc + '  ' + $presetsMenu[$pc]))
        $mCode = [string]$pc
        $mName = [string]$presetsMenu[$pc]
        $mi.add_Click({ try { Set-Stock $mCode $mName } catch { Trace-Msg ('menu-stock err: ' + $_.Exception.Message) } }.GetNewClosure())
        $null = $itStock.DropDownItems.Add($mi)
    }
    $itAutoHide = New-Object System.Windows.Forms.ToolStripMenuItem('鼠标移开自动隐藏')
    $itTop     = New-Object System.Windows.Forms.ToolStripMenuItem('窗口置顶')
    $itOpacity = New-Object System.Windows.Forms.ToolStripMenuItem('透明度设置…')
    $itBody    = New-Object System.Windows.Forms.ToolStripMenuItem('纯透明模式(仅文字)')
    $itDemo    = New-Object System.Windows.Forms.ToolStripMenuItem('演示数据(不联网)')
    $itAuto    = New-Object System.Windows.Forms.ToolStripMenuItem('开机自启动')
    $itExit    = New-Object System.Windows.Forms.ToolStripMenuItem('退出')
    $null = $menu.Items.Add($itRefresh); $null = $menu.Items.Add($itOpen)
    $null = $menu.Items.Add($itStock)
    $null = $menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $null = $menu.Items.Add($itAutoHide); $null = $menu.Items.Add($itTop)
    $null = $menu.Items.Add($itOpacity); $null = $menu.Items.Add($itBody)
    $null = $menu.Items.Add($itDemo)
    $null = $menu.Items.Add($itAuto)
    $null = $menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $null = $menu.Items.Add($itExit)

    $itRefresh.add_Click({ try { Start-Fetch } catch { Trace-Msg ('menu-refresh err: ' + $_.Exception.Message) } })
    $itOpen.add_Click({ try { Start-Process (Get-QuotePageUrl) } catch { Trace-Msg ('menu-open err: ' + $_.Exception.Message) } })
    $itAutoHide.add_Click({ try { $itAutoHide.Checked = -not $itAutoHide.Checked; $script:Cfg.AutoHide = $itAutoHide.Checked; Save-Cfg } catch { Trace-Msg ('menu-autohide err: ' + $_.Exception.Message) } })
    $itTop.add_Click({ try { $itTop.Checked = -not $itTop.Checked; $script:Cfg.TopMost = $itTop.Checked; $form.TopMost = $itTop.Checked; Save-Cfg } catch { Trace-Msg ('menu-top err: ' + $_.Exception.Message) } })
    $itBody.add_Click({ try { $itBody.Checked = -not $itBody.Checked; $script:Cfg.BodyStyle = if ($itBody.Checked) { 'text' } else { 'glass' }; Save-Cfg; $form.Invalidate() } catch { Trace-Msg ('menu-body err: ' + $_.Exception.Message) } })
    $itDemo.add_Click({ try {
        $itDemo.Checked = -not $itDemo.Checked
        $script:DemoOn = $itDemo.Checked
        if ($script:DemoOn) { $script:Quote = Get-DemoQuote; $script:SourceUsed = 'DEMO'; $script:FetchErr = ''; $form.Invalidate() } else { Start-Fetch }
    } catch { Trace-Msg ('menu-demo err: ' + $_.Exception.Message) } })
    $itAuto.add_Click({ try {
        $itAuto.Checked = -not $itAuto.Checked
        if (-not (Set-Autostart $itAuto.Checked)) { $itAuto.Checked = -not $itAuto.Checked }
    } catch { Trace-Msg ('menu-autostart err: ' + $_.Exception.Message) } })
    $itExit.add_Click({ try { $form.Close() } catch { Trace-Msg ('menu-exit err: ' + $_.Exception.Message) } })
    $itOpacity.add_Click({ try { Show-SettingsDialog } catch { Trace-Msg ('menu-opacity err: ' + $_.Exception.Message) } })
    $itAutoHide.Checked = [bool]$cfg.AutoHide
    $itTop.Checked = [bool]$cfg.TopMost
    $itBody.Checked = ($script:Cfg.BodyStyle -eq 'text')
    $itDemo.Checked = $script:DemoOn
    $itAuto.Checked = Test-Autostart

    $menu.add_Opening({ $script:CtxOpen = $true })
    $menu.add_Closed({ $script:CtxOpen = $false })
    $form.ContextMenuStrip = $menu

    # timers
    $tmrRefresh = New-Object System.Windows.Forms.Timer
    $tmrRefresh.Interval = (Get-EffectiveRefreshSec) * 1000
    $tmrRefresh.add_Tick({ try {
        # 每次刷新前按当前交易时段重算间隔 (休市自动降到 60 秒)
        if ($script:TmrRefresh) { $script:TmrRefresh.Interval = (Get-EffectiveRefreshSec) * 1000 }
        Start-Fetch
    } catch { Trace-Msg ('REFRESH ERR: ' + $_.Exception.Message) } })
    $tmrRefresh.Start()

    $script:HoverRect = $null
    $tmrHover = New-Object System.Windows.Forms.Timer
    $tmrHover.Interval = 400
    $tmrHover.add_Tick({
        try {
            if ($script:IsDragging -or $script:CtxOpen) { return }
            $p = [System.Windows.Forms.Cursor]::Position
            $rect = $script:HoverRect
            if ($null -eq $rect -or $rect.X -ne $form.Location.X -or $rect.Y -ne $form.Location.Y -or
                $rect.Width -ne $form.Width -or $rect.Height -ne $form.Height) {
                $rect = New-Object System.Drawing.Rectangle($form.Location, $form.Size)
                $script:HoverRect = $rect
            }
            $inside = $rect.Contains($p)
            $tgt = [double]$script:Cfg.Opacity
            if (-not $inside -and $script:Cfg.AutoHide) { $tgt = [double]$script:Cfg.HiddenOp }
            if ([math]::Abs($tgt - $script:FadeTarget) -gt 0.004) {
                $script:FadeTarget = $tgt
                if (-not $script:FadeRun) { $script:TmrFade.Start(); $script:FadeRun = $true }
            }
            # 配置防抖落盘 (滚轮调透明度/拖动位置后约 0.7 秒统一写入)
            if ($script:CfgDirty -and ((Get-Date) - $script:CfgDirtyAt).TotalMilliseconds -gt 700) {
                $script:CfgDirty = $false
                Save-Cfg
            }
            # 只在内容版本变化(行情/错误/切换标的)时重画: 原先每 250ms 无条件全量重绘
            if ($script:PaintRev -ne $script:SeenRev) {
                $script:SeenRev = $script:PaintRev
                $form.Invalidate()
            }
        } catch { Trace-Msg ('HOVER ERR: ' + $_.Exception.Message) }
    })
    $tmrHover.Start()

    $tmrFade = New-Object System.Windows.Forms.Timer
    $tmrFade.Interval = 16
    $tmrFade.add_Tick({
        try {
            $cur = $form.Opacity
            $diff = $script:FadeTarget - $cur
            if ([math]::Abs($diff) -gt 0.004) {
                $step = 0.05 * [math]::Sign($diff)
                $form.Opacity = [math]::Round([math]::Min(1.0, [math]::Max(0.02, ($cur + $step))), 3)
            } else {
                if ([math]::Abs($cur - $script:FadeTarget) -gt 0.001) { $form.Opacity = $script:FadeTarget }
                # 动画收敛即停表: 空闲时不再每秒 60 次唤醒
                $script:TmrFade.Stop()
                $script:FadeRun = $false
            }
            if ((Get-Date) -gt $script:ToastUntil -and $script:ToastText) {
                $script:ToastText = ''
                $script:PaintRev++
            }
        } catch { Trace-Msg ('FADE ERR: ' + $_.Exception.Message) }
    })

    $tmrPoll = New-Object System.Windows.Forms.Timer
    $tmrPoll.Interval = 300
    $tmrPoll.add_Tick({
        try {
            if (-not $script:Fetching) { $script:TmrPoll.Stop(); return }
            # 一次性子进程兜底模式: 等进程退出
            $proc = $script:FetchProc
            if ($proc -and -not $proc.HasExited -and
                ((Get-Date) - $script:FetchStart).TotalMilliseconds -lt $script:FetchTimeoutMs) { return }
            if ($proc -and -not $proc.HasExited) { try { $proc.Kill() } catch {} }
            $done = Try-Complete-Fetch
            if (-not $done -and ((Get-Date) - $script:FetchStart).TotalMilliseconds -gt $script:FetchTimeoutMs) {
                # 常驻进程卡死: 杀掉重建, 免得后续每轮都卡
                $script:FetchErr = '取数超时'
                $script:PaintRev++
                Stop-FetchHost
                $done = $true
            }
            if ($done) {
                $script:Fetching = $false
                $script:TmrPoll.Stop()
                if ($script:FetchProc) { try { $script:FetchProc.Dispose() } catch {}; $script:FetchProc = $null }
            }
        } catch { Trace-Msg ('POLL ERR: ' + $_.Exception.Message) }
    })

    $form.add_Shown({ Trace-Msg 'shown'; Start-Fetch })
    $form.add_FormClosed({
        try {
            Trace-Msg 'form closed'
            Stop-FetchHost
            if ($script:FetchProc) {
                try { if (-not $script:FetchProc.HasExited) { $script:FetchProc.Kill() } } catch {}
                try { $script:FetchProc.Dispose() } catch {}
                $script:FetchProc = $null
            }
            $script:Cfg.X = $form.Location.X
            $script:Cfg.Y = $form.Location.Y
            Save-Cfg
        } catch { Trace-Msg ('FORMCLOSED ERR: ' + $_.Exception.Message) }
    })

    $script:Form = $form
    $script:TmrRefresh = $tmrRefresh
    $script:TmrHover = $tmrHover
    $script:TmrFade = $tmrFade
    $script:TmrPoll = $tmrPoll
}

# ---- settings dialog ------------------------------------------------
function Show-SettingsDialog {
    $form = $script:Form
    $cfg = $script:Cfg

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = '悬浮窗设置'
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $dlg.ClientSize = New-Object System.Drawing.Size(398, 392)
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.TopMost = $true
    $dlg.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $dlg.BackColor = [System.Drawing.Color]::White

    $presets = Get-StockPresets
    $curCode = ([string]$cfg.Ticker).Trim().ToUpperInvariant()

    $y = 18
    $lblS0 = New-Object System.Windows.Forms.Label
    $lblS0.Text = '股票代码:'
    $lblS0.Location = New-Object System.Drawing.Point(20, $y); $lblS0.AutoSize = $true
    $cmbStock = New-Object System.Windows.Forms.ComboBox
    $cmbStock.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    $cmbStock.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
    $cmbStock.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
    foreach ($k in $presets.Keys) { $null = $cmbStock.Items.Add($k + ' · ' + $presets[$k]) }
    $cmbStock.Location = New-Object System.Drawing.Point(136, ($y - 4)); $cmbStock.Width = 190
    if ($presets.Contains($curCode)) { $cmbStock.Text = $curCode + ' · ' + $presets[$curCode] } else { $cmbStock.Text = $curCode }
    $dlg.Controls.AddRange(@($lblS0, $cmbStock))
    $tipS = New-Object System.Windows.Forms.Label
    $tipS.Text = '支持 A股 600519 / 港股 00700 / 美股 PDD, 可带 sh/sz/hk/us 前缀'
    $tipS.ForeColor = [System.Drawing.Color]::Gray
    $tipS.Location = New-Object System.Drawing.Point(136, ($y + 20)); $tipS.AutoSize = $true
    $dlg.Controls.Add($tipS)

    $y += 52
    $lblO = New-Object System.Windows.Forms.Label
    $lblO.Text = '悬停显示透明度:'
    $lblO.Location = New-Object System.Drawing.Point(20, $y); $lblO.AutoSize = $true
    $trkO = New-Object System.Windows.Forms.TrackBar
    $trkO.Minimum = 20; $trkO.Maximum = 100; $trkO.TickFrequency = 10
    $trkO.Value = [int]([double]$cfg.Opacity * 100)
    $trkO.Location = New-Object System.Drawing.Point(160, ($y - 8)); $trkO.Width = 150
    $valO = New-Object System.Windows.Forms.Label
    $valO.Text = ('{0}%' -f $trkO.Value)
    $valO.Location = New-Object System.Drawing.Point(322, $y); $valO.AutoSize = $true
    $dlg.Controls.AddRange(@($lblO, $trkO, $valO))
    $trkO.add_Scroll({ $valO.Text = ('{0}%' -f $trkO.Value) })

    $y += 50
    $lblH = New-Object System.Windows.Forms.Label
    $lblH.Text = '移开后的隐藏透明度:'
    $lblH.Location = New-Object System.Drawing.Point(20, $y); $lblH.AutoSize = $true
    $trkH = New-Object System.Windows.Forms.TrackBar
    $trkH.Minimum = 4; $trkH.Maximum = 30; $trkH.TickFrequency = 2
    $trkH.Value = [int]([double]$cfg.HiddenOp * 100)
    $trkH.Location = New-Object System.Drawing.Point(160, ($y - 8)); $trkH.Width = 150
    $valH = New-Object System.Windows.Forms.Label
    $valH.Text = ('{0}%' -f $trkH.Value)
    $valH.Location = New-Object System.Drawing.Point(322, $y); $valH.AutoSize = $true
    $dlg.Controls.AddRange(@($lblH, $trkH, $valH))
    $trkH.add_Scroll({ $valH.Text = ('{0}%' -f $trkH.Value) })
    $tipH = New-Object System.Windows.Forms.Label
    $tipH.Text = '注意: 须大于0%才能悬停唤回(建议4~8%)。'
    $tipH.ForeColor = [System.Drawing.Color]::Gray
    $tipH.Location = New-Object System.Drawing.Point(20, ($y + 24)); $tipH.AutoSize = $true
    $dlg.Controls.Add($tipH)

    $y += 60
    $chkHide = New-Object System.Windows.Forms.CheckBox
    $chkHide.Text = '鼠标移开自动隐藏'
    $chkHide.Checked = [bool]$cfg.AutoHide
    $chkHide.Location = New-Object System.Drawing.Point(20, $y); $chkHide.AutoSize = $true
    $chkTop = New-Object System.Windows.Forms.CheckBox
    $chkTop.Text = '窗口置顶'
    $chkTop.Checked = [bool]$cfg.TopMost
    $chkTop.Location = New-Object System.Drawing.Point(210, $y); $chkTop.AutoSize = $true
    $dlg.Controls.AddRange(@($chkHide, $chkTop))

    $y += 32
    $lblR = New-Object System.Windows.Forms.Label
    $lblR.Text = '刷新间隔:'
    $lblR.Location = New-Object System.Drawing.Point(20, $y); $lblR.AutoSize = $true
    $cmbR = New-Object System.Windows.Forms.ComboBox
    $cmbR.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $null = $cmbR.Items.AddRange(@('3 s', '5 s', '6 s', '10 s', '15 s', '30 s'))
    $idx = [Array]::IndexOf(@(3, 5, 6, 10, 15, 30), [int]$cfg.RefreshSec)
    $cmbR.SelectedIndex = if ($idx -ge 0) { $idx } else { 2 }
    $cmbR.Location = New-Object System.Drawing.Point(100, ($y - 4)); $cmbR.Width = 90
    $dlg.Controls.AddRange(@($lblR, $cmbR))

    $y += 32
    $lblS = New-Object System.Windows.Forms.Label
    $lblS.Text = '涨跌配色:'
    $lblS.Location = New-Object System.Drawing.Point(20, $y); $lblS.AutoSize = $true
    $radCn = New-Object System.Windows.Forms.RadioButton
    $radCn.Text = '红涨绿跌'; $radCn.Location = New-Object System.Drawing.Point(90, ($y - 2)); $radCn.AutoSize = $true
    $radUs = New-Object System.Windows.Forms.RadioButton
    $radUs.Text = '绿涨红跌(国际)'; $radUs.Location = New-Object System.Drawing.Point(240, ($y - 2)); $radUs.AutoSize = $true
    if ([bool]$cfg.RiseGreen) { $radUs.Checked = $true } else { $radCn.Checked = $true }
    $dlg.Controls.AddRange(@($lblS, $radCn, $radUs))

    $y += 30
    $lblB = New-Object System.Windows.Forms.Label
    $lblB.Text = '外观:'
    $lblB.Location = New-Object System.Drawing.Point(20, $y); $lblB.AutoSize = $true
    $radGlass = New-Object System.Windows.Forms.RadioButton
    $radGlass.Text = '玻璃卡片'; $radGlass.Location = New-Object System.Drawing.Point(90, ($y - 2)); $radGlass.AutoSize = $true
    $radText = New-Object System.Windows.Forms.RadioButton
    $radText.Text = '纯透明(仅文字)'; $radText.Location = New-Object System.Drawing.Point(205, ($y - 2)); $radText.AutoSize = $true
    if ($cfg.BodyStyle -eq 'text') { $radText.Checked = $true } else { $radGlass.Checked = $true }
    $dlg.Controls.AddRange(@($lblB, $radGlass, $radText))

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = '应用'
    $btnOk.Location = New-Object System.Drawing.Point(230, ($y + 34)); $btnOk.Size = New-Object System.Drawing.Size(100, 30)
    $btnOk.add_Click({ try {
        $raw = ([string]$cmbStock.Text).Trim()
        $code = $raw.ToUpperInvariant()
        $nm = ''
        if ($raw -match '·') { $parts = $raw -split '·'; $code = $parts[0].Trim().ToUpperInvariant(); $nm = $parts[1].Trim() }
        $m2 = [regex]::Match($code, '^(SH|SZ|BJ|HK|US)?[A-Z0-9][A-Z0-9.\-]{0,9}$')
        if ($m2.Success) { $code = $m2.Value } else { $code = 'PDD' }
        if (-not $nm -or $nm -eq $code) { if ($presets.Contains($code)) { $nm = $presets[$code] } else { $nm = $code } }
        $cfg.Ticker = $code
        $cfg.StockName = $nm
        Sync-TickerFromCfg
        $script:SesCache = $null
        $script:Quote = $null
        $script:SourceUsed = ''
        $script:FetchErr = ''
        $cfg.Opacity = [math]::Round($trkO.Value / 100.0, 2)
        $cfg.HiddenOp = [math]::Round($trkH.Value / 100.0, 2)
        $cfg.AutoHide = $chkHide.Checked
        $cfg.TopMost = $chkTop.Checked
        $cfg.RiseGreen = $radUs.Checked
        $cfg.BodyStyle = if ($radText.Checked) { 'text' } else { 'glass' }
        $secs = @(3, 5, 6, 10, 15, 30)[$cmbR.SelectedIndex]
        $cfg.RefreshSec = $secs
        $form.TopMost = [bool]$cfg.TopMost
        if ($script:TmrRefresh) { $script:TmrRefresh.Interval = (Get-EffectiveRefreshSec) * 1000 }
        $script:FadeTarget = [double]$cfg.Opacity
        if (-not $script:FadeRun -and $script:TmrFade) { $script:TmrFade.Start(); $script:FadeRun = $true }
        Update-ThemeColors
        Save-Cfg
        $script:PaintRev++
        if ($script:DemoOn) { $script:Quote = Get-DemoQuote } else { Start-Fetch }
        $form.Invalidate()
        $dlg.Close()
        } catch { Trace-Msg ('SETTINGS APPLY ERR: ' + $_.Exception.Message) }
    })
    $dlg.Controls.Add($btnOk)

    $dlg.ShowDialog($form) | Out-Null
}

# ---- main ------------------------------------------------------------
function Main {
    Trace-Msg 'main start'
    try {
        # 捕获 UI 线程与域级未处理异常, 防止未捕获异常导致进程退出
        try {
            [System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException) | Out-Null
        } catch {}
        try {
            [System.Windows.Forms.Application]::add_ThreadException({
                param($s, $e)
                Trace-Msg ('THREAD EXCEPTION: ' + $e.Exception.ToString())
            })
        } catch {}
        try {
            $null = [System.AppDomain]::CurrentDomain.add_UnhandledException({
                param($s, $e)
                try { Trace-Msg ('APPDOMAIN EXCEPTION: ' + $e.ExceptionObject.ToString()) } catch {}
            })
        } catch {}

        Load-Cfg
        Sync-TickerFromCfg
        Update-ThemeColors
        # 清理上次异常退出遗留的取数临时文件
        try { Get-ChildItem (Join-Path $script:ConfigDir 'fetch-*.json') -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue } catch {}
        # 迁移: 早期版本默认把东财排在首位(本机常不可达), 改为腾讯优先。
        # 只迁移一次, 之后尊重用户在 config.json 里的源顺序。
        try {
            if (-not $script:Cfg.SrcMigrated) {
                $sc = @($script:Cfg.Sources)
                if ($sc.Count -gt 0 -and $sc[0] -eq 'eastmoney') {
                    $script:Cfg.Sources = @('tencent') + @($sc | Where-Object { $_ -ne 'tencent' })
                }
                $script:Cfg.SrcMigrated = $true
                Save-Cfg
            }
        } catch {}
        # 启动诊断: 记录当前市场/时段/下一个节假日/DPI
        try {
            $mk = Get-MarketKind $script:Ticker
            $ses = Get-TradingSession $mk
            $nh = Get-NextMarketHoliday $mk
            $exeInfo = if ($env:PDDWIDGET_EXE) { [string]$env:PDDWIDGET_EXE } else { 'script' }
            Trace-Msg ('startup: ticker=' + $script:Ticker + ' market=' + $mk + ' session=' + $ses.Key + ' nextHoliday=' + $nh + ' scale=' + $script:Scale + ' launchedBy=' + $exeInfo)
        } catch {}

        $owns = $false
        $mutex = $null
        try { $mutex = New-Object System.Threading.Mutex($true, 'PDDStockWidget_v1', [ref]$owns) } catch {}
        if (-not $owns) {
            Trace-Msg 'single-instance: already running'
            try {
                [System.Windows.Forms.MessageBox]::Show('股价悬浮窗已在运行。', '提示',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            } catch {}
            return
        }

        if ($script:DemoOn) {
            $script:Quote = Get-DemoQuote
            $script:SourceUsed = 'DEMO'
        } else {
            $cached = Load-LastQuote
            if ($cached) {
                $script:Quote = $cached
                $script:SourceUsed = '缓存'
            }
        }

        Trace-Msg 'building form'
        Build-Form
        Trace-Msg 'form built, entering message loop'

        [System.Windows.Forms.Application]::Run($script:Form)

        try { $mutex.Dispose() } catch {}
    } catch {
        Trace-Msg ('MAIN EXCEPTION: ' + $_.Exception.ToString())
    }
}

if (-not $NoRun) { Main }
