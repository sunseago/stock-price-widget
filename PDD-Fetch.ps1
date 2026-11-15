# =====================================================================
#  PDD-Fetch.ps1  -  Quote fetcher
#    -Serve : 常驻模式, 从 stdin 读 "ticker|src1,src2|outPath", 结果写文件
#    一次性 : -Ticker xxx -Sources a,b -Out path
#  Writes JSON to -Out file: {"ok":true,"q":{...},"src":"..."}
#  Supports A-shares (600519/sh600519/920047/bj920047), HK (00700/hk00700),
#  US (PDD/usPDD)
# =====================================================================
param(
    [string[]]$Sources = @('tencent', 'eastmoney', 'sina'),
    [string]$Ticker = 'PDD',
    [string]$Out = '',
    [switch]$Serve
)

# Windows PowerShell 5.1: accept -Sources 'a,b,c' as a single comma string
if ($Sources.Count -eq 1 -and $Sources[0] -like '*,*') {
    $Sources = [string[]]($Sources[0] -split ',')
}
$script:DefaultSources = @('tencent', 'eastmoney', 'sina')
$ProgressPreference = 'SilentlyContinue'

# Force TLS 1.2 so quote servers do not reject the connection
try {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
} catch {}

$ErrorActionPreference = 'Continue'

# System.Net.Http is NOT loaded by default in Windows PowerShell 5.1
try { Add-Type -AssemblyName System.Net.Http -ErrorAction Stop } catch {}

# ---- http (single pooled client: 复用连接, 避免每次请求重新握手) ----------
$script:Http = $null
function Get-HttpClient {
    if ($script:Http) { return $script:Http }
    $handler = New-Object System.Net.Http.HttpClientHandler
    $handler.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
    $client = New-Object System.Net.Http.HttpClient($handler)
    # 兜底超时; 单请求超时由 CancellationTokenSource 控制
    $client.Timeout = [TimeSpan]::FromSeconds(60)
    $client.DefaultRequestHeaders.UserAgent.ParseAdd('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36')
    $script:Http = $client
    return $client
}

function Invoke-HttpBytes {
    param([string]$Url, [int]$TimeoutMs = 6000, [hashtable]$Headers = @{})
    $client = Get-HttpClient
    $cts = New-Object System.Threading.CancellationTokenSource
    $resp = $null
    try {
        $cts.CancelAfter($TimeoutMs)
        $req = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Get, $Url)
        try {
            foreach ($k in $Headers.Keys) { [void]$req.Headers.TryAddWithoutValidation($k, [string]$Headers[$k]) }
            # 注意: .NET Framework 的 HttpClient 在 PS5.1 下解析不到 3 参重载, 用 2 参即可
            # (SendAsync(HttpRequestMessage, CancellationToken) 默认就是 ResponseContent 语义)
            $resp = $client.SendAsync($req, $cts.Token).GetAwaiter().GetResult()
        } finally {
            $req.Dispose()
        }
        if (-not $resp.IsSuccessStatusCode) { throw "HTTP $([int]$resp.StatusCode)" }
        return $resp.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'cancel') { throw "timeout after ${TimeoutMs}ms" }
        throw $msg
    } finally {
        $cts.Dispose()
        if ($resp) { $resp.Dispose() }
    }
}
function Invoke-HttpText {
    param([string]$Url, [string]$EncodingName = 'utf-8', [int]$TimeoutMs = 6000, [hashtable]$Headers = @{})
    $bytes = Invoke-HttpBytes -Url $Url -TimeoutMs $TimeoutMs -Headers $Headers
    return [System.Text.Encoding]::GetEncoding($EncodingName).GetString($bytes)
}

# 行情报文一律按不变文化解析: 避免德/法区域(小数逗号)把 "78.97" 解析成 0
function To-DoubleOrZero([string]$s) {
    $d = 0.0
    $ci = [System.Globalization.CultureInfo]::InvariantCulture
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, $ci, [ref]$d)) { return $d }
    return 0.0
}

# ---- market detection ------------------------------------------------
# A股交易所推断: 不能只看首位数字
#   6xxxxx  -> 沪A        900xxx -> 沪B
#   92xxxx  -> 北交所(2024 起新代码段)
#   43/83/87/88xxxx -> 北交所(原精选层/新三板平移)
#   其余(0/1/2/3) -> 深市
function Get-AShareExchange([string]$num) {
    if ($num -match '^6') { return 'sh' }
    if ($num -match '^92') { return 'bj' }
    if ($num -match '^9') { return 'sh' }
    if ($num -match '^(43|83|87|88)') { return 'bj' }
    return 'sz'
}

# Returns @{ Kind='US'|'ASH'|'HK'; Code; Tc; EmMarkets=int[]; Cur; Sina; Ex }
# 沪B(900xxx)以美元计价、深B(200xxx)以港元计价, 成交额单位不能一律按"元"
function Get-BShareCurrency([string]$num, [string]$default = 'CNY') {
    if ($num -match '^900[0-9]{3}$') { return 'USD' }
    if ($num -match '^200[0-9]{3}$') { return 'HKD' }
    return $default
}

function Resolve-Market([string]$raw) {
    $t = $raw.Trim().ToLower()
    if ($t -match '^(sh|sz|bj)([0-9]{6})$') {
        $ex = $Matches[1]; $num = $Matches[2]
        # 北交所: 腾讯 bj 前缀已验证可用 (qt.gtimg.cn/q=bj920047); 东财市场号 0 有效
        $mk = if ($ex -eq 'sh') { @(1) } elseif ($ex -eq 'sz') { @(0) } else { @(0, 1) }
        return @{ Kind='ASH'; Code=$num; Tc="$ex$num"; EmMarkets=$mk; Cur=(Get-BShareCurrency $num 'CNY'); Sina="$ex$num"; Ex=$ex }
    }
    if ($t -match '^hk([0-9]{4,5})$') {
        $num = $Matches[1].PadLeft(5, '0')
        return @{ Kind='HK'; Code=$num; Tc="hk$num"; EmMarkets=@(116); Cur='HKD'; Sina="rt_hk$num"; Ex='hk' }
    }
    if ($t -match '^us([a-z][a-z0-9.\-]*)$') {
        $sym = $Matches[1].ToUpper()
        return @{ Kind='US'; Code=$sym; Tc="us$sym"; EmMarkets=@(105, 106, 107); Cur='USD'; Sina="gb_$($sym.ToLower())"; Ex='us' }
    }
    if ($t -match '^([0-9]{6})$') {
        $num = $Matches[1]
        $ex = Get-AShareExchange $num
        $mk = if ($ex -eq 'sh') { @(1) } elseif ($ex -eq 'sz') { @(0) } else { @(0, 1) }
        return @{ Kind='ASH'; Code=$num; Tc="$ex$num"; EmMarkets=$mk; Cur=(Get-BShareCurrency $num 'CNY'); Sina="$ex$num"; Ex=$ex }
    }
    if ($t -match '^([0-9]{4,5})$') {
        $num = $Matches[1].PadLeft(5, '0')
        return @{ Kind='HK'; Code=$num; Tc="hk$num"; EmMarkets=@(116); Cur='HKD'; Sina="rt_hk$num"; Ex='hk' }
    }
    if ($t -match '^([a-z][a-z0-9.\-]*)$') {
        $sym = $t.ToUpper()
        return @{ Kind='US'; Code=$sym; Tc="us$sym"; EmMarkets=@(105, 106, 107); Cur='USD'; Sina="gb_$($sym.ToLower())"; Ex='us' }
    }
    throw "无法识别的股票代码: $raw (支持 A股 600519 / 北交所 920047 / 港股 00700 / 美股 PDD, 可加 sh/sz/bj/hk/us 前缀)"
}

function New-Quote {
    param([string]$Name,[string]$Code,[double]$Price,[double]$Prev,[double]$Open,[double]$High,[double]$Low,
          [double]$Vol,[double]$Amt,[string]$QTime,[string]$Source,[string]$Warn = '')
    if ($Prev -le 0) { $Prev = 0 }
    $chg = 0.0; $pct = 0.0
    if ($Prev -gt 0) { $chg = $Price - $Prev; $pct = ($Price / $Prev - 1) * 100 }
    $up = 0
    if ($chg -gt 0.004) { $up = 1 } elseif ($chg -lt -0.004) { $up = -1 }
    return [pscustomobject]@{
        Name=$Name; Code=$Code; Price=$Price; PrevClose=$Prev; Open=$Open; High=$High; Low=$Low
        Volume=$Vol; Amount=$Amt; QuoteTime=$QTime; Source=$Source; Chg=$chg; Pct=$pct; Up=$up
        Currency=$script:Market.Cur; Market=$script:Market.Kind; Warn=$Warn
    }
}

# unix 秒 -> 该市场的本地时间字符串
function Format-QuoteStamp([double]$unixSec) {
    try {
        $utc = [datetime]::SpecifyKind([datetime]'1970-01-01', [System.DateTimeKind]::Utc).AddSeconds($unixSec)
        if ($script:Market.Kind -eq 'US') {
            try {
                $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById('Eastern Standard Time')
                $lt = [System.TimeZoneInfo]::ConvertTimeFromUtc($utc, $tz)
            } catch { $lt = $utc.AddHours(-5) }
        } else {
            $lt = $utc.AddHours(8)   # A股/港股均为 UTC+8 且无夏令时
        }
        return $lt.ToString('yyyy-MM-dd HH:mm:ss')
    } catch { return '' }
}

# ---- providers --------------------------------------------------------
function Get-Quote-Eastmoney {
    $errText = ''
    foreach ($mk in $script:Market.EmMarkets) {
        try {
            $url = "https://push2.eastmoney.com/api/qt/stock/get?secid=$mk.$($script:Market.Code)&fields=f43,f44,f45,f46,f47,f48,f57,f58,f59,f60,f86,f169,f170"
            $txt = Invoke-HttpText -Url $url -EncodingName 'utf-8' -TimeoutMs 5000
            $j = $txt | ConvertFrom-Json
            if ($j.rc -ne 0 -or $null -eq $j.data) { continue }
            $d = $j.data
            # 价格小数位由 f59 给出(A股 2 / 港股 3 / 美股 3), 不再按市场写死除数
            $dec = 2
            if ($null -ne $d.f59 -and [string]$d.f59 -ne '') { $dec = [int]$d.f59 }
            $div = [math]::Pow(10, $dec)
            $price = [double]$d.f43 / $div
            if ($price -le 0) { continue }
            $vol = [double]$d.f47
            if ($script:Market.Kind -eq 'ASH') { $vol = $vol * 100 }   # 东财 A股成交量为"手", 统一归一为股
            $qtime = ''
            if ($null -ne $d.f86 -and [double]$d.f86 -gt 0) { $qtime = Format-QuoteStamp ([double]$d.f86) }
            return (New-Quote -Name ([string]$d.f58) -Code ([string]$d.f57) -Price $price `
                -Prev ([double]$d.f60/$div) -Open ([double]$d.f46/$div) -High ([double]$d.f44/$div) -Low ([double]$d.f45/$div) `
                -Vol $vol -Amt ([double]$d.f48) -QTime $qtime -Source 'eastmoney')
        } catch { $errText = $_.Exception.Message }
    }
    throw 'eastmoney no quote for ' + $Ticker + ' (' + $errText + ')'
}

function Get-Quote-Tencent {
    $tc = $script:Market.Tc
    $url = "https://qt.gtimg.cn/q=$tc"
    $txt = Invoke-HttpText -Url $url -EncodingName 'gb18030' -TimeoutMs 6000 -Headers @{ Referer = 'https://gu.qq.com/' }
    $m = [regex]::Match($txt, 'v_' + $tc + '="([^"]*)"')
    if (-not $m.Success) { $m = [regex]::Match($txt, 'v_\w+="([^"]*)"') }
    if (-not $m.Success) { throw 'tencent bad response' }
    $f = $m.Groups[1].Value -split '~'
    if ($f.Count -lt 40) { throw 'tencent missing fields' }
    $price = To-DoubleOrZero $f[3]
    if ($price -le 0) { throw 'invalid price' }
    $kind = $script:Market.Kind
    if ($kind -eq 'US') {
        # 以 "USD" 标记为锚点: ...时间~涨跌~涨跌幅~最高~最低~USD~成交量~成交额...
        $idxUsd = [Array]::IndexOf($f, 'USD')
        if ($idxUsd -lt 8 -or $idxUsd -gt ($f.Count - 5)) { throw 'tencent layout unexpected' }
        $name   = if ($f[1]) { [string]$f[1] } else { $script:Market.Code }
        $high   = To-DoubleOrZero $f[$idxUsd - 2]
        $low    = To-DoubleOrZero $f[$idxUsd - 1]
        $vol    = To-DoubleOrZero $f[$idxUsd + 1]
        $amt    = To-DoubleOrZero $f[$idxUsd + 2]
        $qtime  = [string]$f[$idxUsd - 5]
        # 腾讯美股代码带交易所后缀 (PDD.OQ), 去掉后缀与用户输入保持一致
        $code   = ([string]$f[2]) -replace '\.(OQ|NQ|NY|US|AM|PK|OB|N)\.?$', ''
        if (-not $code) { $code = $script:Market.Code }
        return (New-Quote -Name $name -Code $code -Price $price `
            -Prev (To-DoubleOrZero $f[4]) -Open (To-DoubleOrZero $f[5]) -High $high -Low $low `
            -Vol $vol -Amt $amt -QTime $qtime -Source 'tencent')
    }
    # A股/港股统一布局: 3现价 4昨收 5今开 6量 30时间 31涨跌 32涨幅 33最高 34最低 37成交额
    $name  = if ($f[1]) { [string]$f[1] } else { $script:Market.Code }
    $prev  = To-DoubleOrZero $f[4]
    $open  = To-DoubleOrZero $f[5]
    $vol   = To-DoubleOrZero $f[6]
    if ($kind -eq 'ASH') { $vol = $vol * 100 }   # 腾讯 A股成交量为"手", 统一归一为股
    $high  = To-DoubleOrZero $f[33]
    $low   = To-DoubleOrZero $f[34]
    $amt   = To-DoubleOrZero $f[37]
    $rawT  = [string]$f[30]
    $qtime = $rawT
    if ($kind -eq 'ASH') {
        # A股时间 20260909161436, 成交额单位 万 -> 元
        $amt = $amt * 10000.0
        if ($rawT -match '^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})$') {
            $qtime = '{0}-{1}-{2} {3}:{4}:{5}' -f $Matches[1], $Matches[2], $Matches[3], $Matches[4], $Matches[5], $Matches[6]
        }
    } elseif ($kind -eq 'HK' -and $rawT) {
        # 港股时间 2026/11/13 16:08:24 -> 统一成分隔符为 "-" 的格式
        $qtime = $rawT -replace '/', '-'
    }
    # 北交所已迁移/停用的老代码: 腾讯返回早盘 09:00 占位价 + 量额全 0, 不能当成实时行情
    $warn = ''
    if ($script:Market.Ex -eq 'bj' -and $script:Market.Code -match '^(43|83|87|88)' -and $vol -le 0 -and $amt -le 0 -and $price -gt 0) {
        $warn = '量额为 0, 可能已停用/迁移至 92xxxx'
    }
    return (New-Quote -Name $name -Code ([string]$f[2]) -Price $price `
        -Prev $prev -Open $open -High $high -Low $low `
        -Vol $vol -Amt $amt -QTime $qtime -Source 'tencent' -Warn $warn)
}

function Get-Quote-Sina {
    $sym = $script:Market.Sina
    $url = "https://hq.sinajs.cn/list=$sym"
    $txt = Invoke-HttpText -Url $url -EncodingName 'gb18030' -TimeoutMs 5000 -Headers @{ Referer = 'https://finance.sina.com.cn' }
    $m = [regex]::Match($txt, '"(.*)"')
    if (-not $m.Success) { throw 'sina bad response' }
    $f = $m.Groups[1].Value -split ','
    if ($f.Count -lt 6) { throw 'sina missing fields' }
    $kind = $script:Market.Kind
    if ($kind -eq 'ASH') {
        # A股: name,open,prev,current,high,low,buy,sell,vol(股),amount(元),...
        $price = To-DoubleOrZero $f[3]
        if ($price -le 0) { throw 'invalid price' }
        $prev = To-DoubleOrZero $f[2]
        $dt = ''
        if ($f.Count -gt 31) { $dt = ([string]$f[30]) + ' ' + ([string]$f[31]) }
        return (New-Quote -Name ([string]$f[0]) -Code $script:Market.Code -Price $price -Prev $prev `
            -Open (To-DoubleOrZero $f[1]) -High (To-DoubleOrZero $f[4]) -Low (To-DoubleOrZero $f[5]) `
            -Vol (To-DoubleOrZero $f[8]) -Amt (To-DoubleOrZero $f[9]) -QTime $dt -Source 'sina')
    }
    if ($kind -eq 'US') {
        # 美股实际布局(实测): 0名称 1现价 2涨跌幅% 3时间 4涨跌额 5今开 6最高 7最低 8/9 52周高低 10成交量
        $price = To-DoubleOrZero $f[1]
        if ($price -le 0) { throw 'invalid price' }
        $chg = To-DoubleOrZero $f[4]
        $prev = 0.0
        if ($chg -ne 0) { $prev = $price - $chg }
        $vol = 0.0; if ($f.Count -gt 10) { $vol = To-DoubleOrZero $f[10] }
        return (New-Quote -Name ([string]$f[0]) -Code $script:Market.Code -Price $price -Prev $prev `
            -Open (To-DoubleOrZero $f[5]) -High (To-DoubleOrZero $f[6]) -Low (To-DoubleOrZero $f[7]) `
            -Vol $vol -Amt 0 -QTime ([string]$f[3]) -Source 'sina')
    }
    throw 'sina 不支持港股'
}

function Write-JsonAtomic([string]$path, [string]$json) {
    if (-not $path) { Write-Output $json; return }
    $tmp = $path + '.tmp'
    [System.IO.File]::WriteAllText($tmp, $json, (New-Object System.Text.UTF8Encoding($true)))
    try {
        if ([System.IO.File]::Exists($path)) { [System.IO.File]::Replace($tmp, $path, $null) }
        else { [System.IO.File]::Move($tmp, $path) }
    } catch {
        try { [System.IO.File]::Copy($tmp, $path, $true); [System.IO.File]::Delete($tmp) } catch {}
    }
}

function Emit-Result([bool]$ok, $q, [string]$src, [string]$err, [string]$out) {
    $obj = @{ ok = $ok }
    if ($q) { $obj.q = $q; $obj.src = $src }
    if ($err) { $obj.err = $err }
    $json = $obj | ConvertTo-Json -Compress -Depth 5
    Write-JsonAtomic $out $json
}

function Invoke-FetchRound([string]$ticker, [string[]]$sources, [string]$out) {
    $script:Market = Resolve-Market $ticker
    $errs = @()
    foreach ($src in $sources) {
        try {
            $q = switch ($src) {
                'eastmoney' { Get-Quote-Eastmoney }
                'tencent'   { Get-Quote-Tencent }
                'sina'      { Get-Quote-Sina }
                default     { throw "unknown source: $src" }
            }
            if ($null -ne $q -and $q.Price -gt 0) {
                Emit-Result $true $q $src '' $out
                return $true
            }
        } catch {
            $errs += "[$src] $($_.Exception.Message)"
        }
    }
    $errText = ($errs -join '; ')
    if ($errText.Length -gt 180) { $errText = $errText.Substring(0, 180) }
    Emit-Result $false $null '' $errText $out
    return $false
}

# ---- main ------------------------------------------------------------
if ($Serve) {
    # 常驻模式: 一次启动, 后续每轮只写一行命令, 省掉每次 ~400ms 的进程冷启动
    $stdin = [Console]::In
    while ($true) {
        $line = $stdin.ReadLine()
        if ($null -eq $line) { break }
        $line = $line.Trim()
        if ($line -eq '') { continue }
        if ($line -eq 'QUIT') { break }
        try {
            $p = $line -split '\|'
            $t = if ($p.Count -ge 1 -and $p[0]) { $p[0] } else { 'PDD' }
            $s = if ($p.Count -ge 2 -and $p[1]) { [string[]]($p[1] -split ',') } else { $script:DefaultSources }
            $o = if ($p.Count -ge 3) { $p[2] } else { '' }
            $null = Invoke-FetchRound $t $s $o
        } catch {
            try { Emit-Result $false $null '' ($_.Exception.Message) $o } catch {}
        }
    }
    try { if ($script:Http) { $script:Http.Dispose() } } catch {}
    exit 0
}

$ok = Invoke-FetchRound $Ticker $Sources $Out
try { if ($script:Http) { $script:Http.Dispose() } } catch {}
if ($ok) { exit 0 } else { exit 1 }
