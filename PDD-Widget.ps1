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

param([switch]$Demo)

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---- constants -------------------------------------------------------
$script:Ticker    = 'PDD'
$script:StockName   = '拼多多'
$script:FetchScript = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) 'PDD-Fetch.ps1'
$script:FetchTimeoutMs = 30000
$script:DefaultSources = @('tencent', 'eastmoney', 'sina')

function Get-MarketKind([string]$code) {
    $c = $code.Trim().ToLower()
    if ($c -match '^(sh|sz|bj)?[0-9]{6}$') { return 'ASH' }
    if ($c -match '^hk[0-9]{4,5}$') { return 'HK' }
    if ($c -match '^[0-9]{4,5}$') { return 'HK' }
    return 'US'
}
function Get-QuotePageUrl {
    $raw = ([string]$script:Cfg.Ticker).Trim().ToUpperInvariant()
    if (-not $raw) { $raw = 'PDD' }
    $kind = Get-MarketKind $raw
    $c = $raw -replace '^(SH|SZ|BJ|HK|US)', ''
    if ($kind -eq 'ASH') {
        $ex = if ($c[0] -eq '6' -or $c[0] -eq '9') { 'sh' } elseif ($c[0] -eq '4' -or $c[0] -eq '8') { 'bj' } else { 'sz' }
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
    if ($script:DemoOn) { $script:Quote = Get-DemoQuote } else { Start-Fetch }
    if ($script:Form) { $script:Form.Invalidate() }
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
}
$script:Quote = $null
$script:SourceUsed = ''
$script:FetchErr = ''
$script:Fetching = $false
$script:FetchProc = $null
$script:FetchTmp = ''
$script:FetchStart = [datetime]::MinValue
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
    $labelMap = @{ PRE = '盘前'; REGULAR = '交易中'; AFTER = '盘后'; CLOSED = '休市' }
    $mktLabel = @{ US = '美股'; ASH = 'A股'; HK = '港股' }[$kind]
    if ($kind -eq 'US') {
        $tz = [TimeZoneInfo]::FindSystemTimeZoneById('Eastern Standard Time')
        $lt = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
    } else {
        # A股/港股均为 UTC+8 且无夏令时
        $lt = [DateTime]::UtcNow.AddHours(8)
    }
    $res = @{ Key = 'CLOSED'; Label = '休市'; Local = $lt; Mkt = $mktLabel; Minutes = ($lt.Hour * 60 + $lt.Minute) }
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
        # 9:30-11:30 / 13:00-15:00 (北京时间, 已含 2026 法定节假日)
        if (($res.Minutes -ge 570 -and $res.Minutes -lt 690) -or ($res.Minutes -ge 780 -and $res.Minutes -lt 900)) { $res.Key = 'REGULAR' }
    } else {
        # 港股 9:30-12:00 / 13:00-16:00
        if (($res.Minutes -ge 570 -and $res.Minutes -lt 720) -or ($res.Minutes -ge 780 -and $res.Minutes -lt 960)) { $res.Key = 'REGULAR' }
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
function Format-Vol([double]$v) {
    if ($v -ge 1e8) { return ('{0:N2}亿' -f ($v / 1e8)) }
    if ($v -ge 1e4) { return ('{0:N2}万' -f ($v / 1e4)) }
    return ('{0:N0}' -f $v)
}
function Format-Amt([double]$v, [string]$cur) {
    $unit = @{ USD = '美元'; CNY = '元'; HKD = '港元' }[$cur]
    if (-not $unit) { $unit = '美元' }
    if ($v -ge 1e8) { return ('{0:N2}亿{1}' -f ($v / 1e8), $unit) }
    if ($v -ge 1e4) { return ('{0:N1}万{1}' -f ($v / 1e4), $unit) }
    return ('{0:N0}{1}' -f $v, $unit)
}

# ---- async fetch (child process) -------------------------------------
function Start-Fetch {
    if ($script:DemoOn -or $script:Fetching) { return }
    if ($script:FetchProc -and $script:FetchProc.HasExited) {
        try { $script:FetchProc.Dispose() } catch {}
        $script:FetchProc = $null
    }
    if ($script:FetchProc) { return }
    $script:FetchTmp = Join-Path $script:ConfigDir ('fetch-' + [guid]::NewGuid().ToString('N') + '.json')
    # 白名单校验: 防止被手工改坏的 config.json 把参数注入到子进程
    $tk = ([string]$script:Ticker).Trim()
    if ($tk -notmatch '^[A-Za-z0-9.\-]{1,12}$') { $tk = 'PDD' }
    $srcs = @($script:Cfg.Sources | Where-Object { $_ -in @('tencent', 'eastmoney', 'sina') })
    if ($srcs.Count -eq 0) { $srcs = @('tencent') }
    $sources = ($srcs -join ',')
    $fs = $script:FetchScript
    $args = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Sta',
              '-File', ('"' + $fs + '"'),
              '-Ticker', $tk,
              '-Sources', $sources,
              '-Out', ('"' + $script:FetchTmp + '"'))
    try {
        $script:FetchProc = Start-Process -FilePath 'powershell.exe' -ArgumentList $args -WindowStyle Hidden -PassThru
        $script:Fetching = $true
        $script:FetchStart = Get-Date
        if ($script:TmrPoll) { $script:TmrPoll.Start() }
    } catch {
        $script:FetchErr = '启动取数失败: ' + $_.Exception.Message
        $script:FetchProc = $null
    }
}

function Complete-Fetch {
    $script:Fetching = $false
    try { if ($script:TmrPoll) { $script:TmrPoll.Stop() } } catch {}
    try {
        if (Test-Path $script:FetchTmp) {
            $txt = Get-Content $script:FetchTmp -Raw -Encoding UTF8
            Remove-Item $script:FetchTmp -Force -ErrorAction SilentlyContinue
            $j = $txt | ConvertFrom-Json
            if ($j -and $j.ok) {
                $d = $j.q
                $q = [pscustomobject]@{
                    Name = [string]$d.Name; Code = [string]$d.Code
                    Price = [double]$d.Price; PrevClose = [double]$d.PrevClose
                    Open = [double]$d.Open; High = [double]$d.High; Low = [double]$d.Low
                    Volume = [double]$d.Volume; Amount = [double]$d.Amount
                    QuoteTime = [string]$d.QuoteTime
                    Chg = [double]$d.Chg; Pct = [double]$d.Pct; Up = [int]$d.Up; Stale = $false
                    Currency = [string]$d.Currency; Market = [string]$d.Market
                }
                $script:Quote = $q
                $sn = [string]$j.src
                $script:SourceUsed = switch ($sn) { 'eastmoney' { '东财' } 'tencent' { '腾讯' } 'sina' { '新浪' } default { $sn } }
                $script:FetchErr = ''
                Save-LastQuote $q
            } else {
                $script:FetchErr = if ($j -and $j.err) { [string]$j.err } else { '无数据' }
            }
        } else {
            $script:FetchErr = '取数超时'
        }
    } catch {
        $script:FetchErr = '解析失败: ' + $_.Exception.Message
    }
    if ($script:Form) { $script:Form.Invalidate() }
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
function Show-Toast([string]$msg, [int]$ms = 1600) {
    $script:ToastText = $msg
    $script:ToastUntil = (Get-Date).AddMilliseconds($ms)
    if ($script:Form) { $script:Form.Invalidate() }
}

# ---- paint -----------------------------------------------------------
function Draw-Card($g) {
    $form = $script:Form
    if (-not $form) { return }
    $w = $form.ClientSize.Width; $h = $form.ClientSize.Height
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    if ($script:Cfg.BodyStyle -ne 'text') {
        $rect = New-Object System.Drawing.Rectangle(1, 1, ($w - 2), ($h - 2))
        $grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect,
            [System.Drawing.Color]::FromArgb(255, 26, 29, 37),
            [System.Drawing.Color]::FromArgb(255, 12, 14, 19), 90.0)
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $r = 16
        $path.AddArc(1, 1, ($r * 2), ($r * 2), 180, 90)
        $path.AddArc(($w - 1 - $r * 2), 1, ($r * 2), ($r * 2), 270, 90)
        $path.AddArc(($w - 1 - $r * 2), ($h - 1 - $r * 2), ($r * 2), ($r * 2), 0, 90)
        $path.AddArc(1, ($h - 1 - $r * 2), ($r * 2), ($r * 2), 90, 90)
        $path.CloseFigure()
        $g.FillPath($grad, $path)
        $g.DrawPath($script:PenCard, $path)
        $path.Dispose(); $grad.Dispose()
    }

    $x0 = 16
    $q = $script:Quote
    $dirBrush = $script:BrFlat
    if ($q) { $dirBrush = Get-DirBrush $q.Up }

    # header
    $name = "$($script:Ticker)  $($script:StockName)"
    if ($q -and $q.Name) { $name = "$($q.Code)  $($q.Name)" }
    $g.DrawString($name, $script:FName, $script:BrText, $x0, 12)

    $st = Get-TradingSession (Get-MarketKind $script:Ticker)
    $badge = if ($script:DemoOn) { '演示' } else { $st.Label }
    $badgeCol = $script:BrBadge
    if ($st.Key -eq 'PRE' -or $st.Key -eq 'AFTER') {
        $badgeCol = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 226, 160, 58))
    } elseif ($st.Key -eq 'CLOSED') { $badgeCol = $script:BrBadgeG }
    $bw = 66; $bh = 20; $bx = $w - 16 - $bw; $by = 9
    $bpath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $bpath.AddArc($bx, $by, 10, 10, 180, 90); $bpath.AddArc(($bx + $bw - 10), $by, 10, 10, 270, 90)
    $bpath.AddArc(($bx + $bw - 10), ($by + $bh - 10), 10, 10, 0, 90); $bpath.AddArc($bx, ($by + $bh - 10), 10, 10, 90, 90)
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
        $g.DrawString(('{0:N2}' -f $q.Price), $script:FPrice, $dirBrush, $x0, 40)
        $chgTxt = if ($q.Chg -ge 0) { ('+{0:N2}' -f $q.Chg) } else { ('{0:N2}' -f $q.Chg) }
        $pctTxt = if ($q.Pct -ge 0) { ('+{0:N2}%' -f $q.Pct) } else { ('{0:N2}%' -f $q.Pct) }
        $arrow = if ($q.Up -gt 0) { '▲ ' } elseif ($q.Up -lt 0) { '▼ ' } else { '— ' }
        $g.DrawString(("$arrow$chgTxt  $pctTxt"), $script:FPct, $dirBrush, 158, 47)
    } else {
        $msg = if ($script:Fetching) { '获取行情中…' } elseif ($script:FetchErr) { '获取失败·网络不可达' } else { '等待数据…' }
        $g.DrawString($msg, $script:FPrice, $script:BrDim, $x0, 42)
    }

    # detail lines
    $meta1 = ''; $meta2 = ''
    if ($q) {
        $o = if ($q.Open -gt 0) { ('{0:N2}' -f $q.Open) } else { '-' }
        $p = if ($q.PrevClose -gt 0) { ('{0:N2}' -f $q.PrevClose) } else { '-' }
        $hi = if ($q.High -gt 0) { ('{0:N2}' -f $q.High) } else { '-' }
        $lo = if ($q.Low -gt 0) { ('{0:N2}' -f $q.Low) } else { '-' }
        $meta1 = "昨收 $p  今开 $o  最高 $hi  最低 $lo"
        $volTxt = if ($q.Volume -gt 0) { Format-Vol $q.Volume } else { '-' }
        $amtTxt = if ($q.Amount -gt 0) { Format-Amt $q.Amount ([string]$q.Currency) } else { '-' }
        $meta2 = "量 $volTxt   额 $amtTxt"
        if ($q.QuoteTime) { $meta2 += "   " + $q.QuoteTime }
    }
    $g.DrawString($meta1, $script:FMeta, $script:BrDim, $x0, 96)
    if (-not $q -and $script:FetchErr) {
        $e = $script:FetchErr
        if ($e.Length -gt 46) { $e = $e.Substring(0, 46) + '...' }
        $g.DrawString($e, $script:FTiny, $script:BrDim, $x0, 118)
    } else {
        $g.DrawString($meta2, $script:FTiny, $script:BrDim, $x0, 118)
    }

    # footer
    $foot = ''
    if (-not $script:DemoOn) {
        $ses = Get-TradingSession (Get-MarketKind $script:Ticker)
        $stTxt = $ses.Mkt + ' ' + $ses.Local.ToString('HH:mm') + '  ' + $ses.Label
        if ($q) { $foot = "数据源 $($script:SourceUsed) · $stTxt" }
        elseif ($script:Fetching) { $foot = "刷新中… · $stTxt" }
        else { $foot = $stTxt }
    } else { $foot = '演示模式' }
    $sf2 = New-Object System.Drawing.StringFormat
    $sf2.Alignment = [System.Drawing.StringAlignment]::Far
    $g.DrawString($foot, $script:FTiny, $script:BrDim,
        (New-Object System.Drawing.RectangleF(16, 132, ($w - 32), 16)), $sf2)
    $sf2.Dispose()

    # toast
    if ($script:ToastText -and (Get-Date) -lt $script:ToastUntil) {
        $sz = $g.MeasureString($script:ToastText, $script:FTiny)
        $tw = [math]::Ceiling($sz.Width) + 24; $th = [math]::Ceiling($sz.Height) + 10
        $tx = [int](($w - $tw) / 2); $ty = $h - $th - 6
        $tb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(220, 8, 10, 14))
        $g.FillRectangle($tb, $tx, $ty, $tw, $th)
        $g.DrawString($script:ToastText, $script:FTiny, $script:BrText, ($tx + 12), ($ty + 4))
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
    $W = 350; $H = 152; $R = 16
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
    $script:PenCard  = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(60, 255, 255, 255), 1)

    $form.add_Paint({ param($s, $e) try { Draw-Card $e.Graphics } catch { Trace-Msg ('PAINT ERR: ' + $_.Exception.Message) } })
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

    $tmrHover = New-Object System.Windows.Forms.Timer
    $tmrHover.Interval = 250
    $tmrHover.add_Tick({
        try {
            if ($script:IsDragging -or $script:CtxOpen) { return }
            $p = [System.Windows.Forms.Cursor]::Position
            $r = New-Object System.Drawing.Rectangle($form.Location, $form.Size)
            $inside = $r.Contains($p)
            if ($inside) { $script:FadeTarget = [double]$script:Cfg.Opacity }
            elseif ($script:Cfg.AutoHide) { $script:FadeTarget = [double]$script:Cfg.HiddenOp }
            else { $script:FadeTarget = [double]$script:Cfg.Opacity }
            # 配置防抖落盘 (滚轮调透明度/拖动位置后约 0.7 秒统一写入)
            if ($script:CfgDirty -and ((Get-Date) - $script:CfgDirtyAt).TotalMilliseconds -gt 700) {
                $script:CfgDirty = $false
                Save-Cfg
            }
            $form.Invalidate()
        } catch { Trace-Msg ('HOVER ERR: ' + $_.Exception.Message) }
    })
    $tmrHover.Start()

    $tmrFade = New-Object System.Windows.Forms.Timer
    $tmrFade.Interval = 18
    $tmrFade.add_Tick({
        try {
            $cur = $form.Opacity
            $diff = $script:FadeTarget - $cur
            if ([math]::Abs($diff) -gt 0.004) {
                $step = 0.05 * [math]::Sign($diff)
                $form.Opacity = [math]::Round([math]::Min(1.0, [math]::Max(0.02, ($cur + $step))), 3)
            } elseif ([math]::Abs($cur - $script:FadeTarget) -gt 0.002) {
                $form.Opacity = $script:FadeTarget
            }
            if ((Get-Date) -gt $script:ToastUntil) {
                if ($script:ToastText) { $script:ToastText = ''; $form.Invalidate() }
            }
        } catch { Trace-Msg ('FADE ERR: ' + $_.Exception.Message) }
    })
    $tmrFade.Start()

    $tmrPoll = New-Object System.Windows.Forms.Timer
    $tmrPoll.Interval = 400
    $tmrPoll.add_Tick({
        try {
            if (-not $script:Fetching) { return }
            $proc = $script:FetchProc
            if ($proc -and -not $proc.HasExited) {
                if (((Get-Date) - $script:FetchStart).TotalMilliseconds -gt $script:FetchTimeoutMs) {
                    try { $proc.Kill() } catch {}
                } else { return }
            }
            Complete-Fetch
            if ($script:FetchProc) { try { $script:FetchProc.Dispose() } catch {}; $script:FetchProc = $null }
        } catch { Trace-Msg ('POLL ERR: ' + $_.Exception.Message) }
    })

    $form.add_Shown({ Trace-Msg 'shown'; Start-Fetch })
    $form.add_FormClosed({
        try {
            Trace-Msg 'form closed'
            if ($script:FetchProc) {
                try { if (-not $script:FetchProc.HasExited) { $script:FetchProc.Kill() } } catch {}
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
        Update-ThemeColors
        Save-Cfg
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
        # 迁移: 早期版本默认把东财排在首位(本机常不可达), 改为腾讯优先
        try {
            $sc = @($script:Cfg.Sources)
            if ($sc.Count -gt 0 -and $sc[0] -eq 'eastmoney') {
                $script:Cfg.Sources = @('tencent') + @($sc | Where-Object { $_ -ne 'tencent' })
                Save-Cfg
            }
        } catch {}
        # 启动诊断: 记录当前市场/时段/下一个节假日
        try {
            $mk = Get-MarketKind $script:Ticker
            $ses = Get-TradingSession $mk
            $nh = Get-NextMarketHoliday $mk
            $exeInfo = if ($env:PDDWIDGET_EXE) { [string]$env:PDDWIDGET_EXE } else { 'script' }
            Trace-Msg ('startup: ticker=' + $script:Ticker + ' market=' + $mk + ' session=' + $ses.Key + ' nextHoliday=' + $nh + ' launchedBy=' + $exeInfo)
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

Main
