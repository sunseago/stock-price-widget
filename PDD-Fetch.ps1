# =====================================================================
#  PDD-Fetch.ps1  -  Quote fetcher subprocess (called async by PDD-Widget.ps1)
#  Writes JSON to -Out file: {"ok":true,"q":{...},"src":"..."}
#  Supports US stocks (PDD), A-shares (600519/sh600519), HK stocks (00700/hk00700)
# =====================================================================
param(
    [string[]]$Sources = @('tencent', 'eastmoney', 'sina'),
    [string]$Ticker = 'PDD',
    [string]$Out = ''
)

# Windows PowerShell 5.1: accept -Sources 'a,b,c' as a single comma string
if ($Sources.Count -eq 1 -and $Sources[0] -like '*,*') {
    $Sources = [string[]]($Sources[0] -split ',')
}

# Force TLS 1.2 so quote servers do not reject the connection
try {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
} catch {}

$ErrorActionPreference = 'Continue'

# System.Net.Http is NOT loaded by default in Windows PowerShell 5.1
try { Add-Type -AssemblyName System.Net.Http -ErrorAction Stop } catch {}

function Invoke-HttpBytes {
    param([string]$Url, [int]$TimeoutMs = 6000, [hashtable]$Headers = @{})
    $handler = New-Object System.Net.Http.HttpClientHandler
    $handler.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
    $client = New-Object System.Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromMilliseconds($TimeoutMs)
    try {
        $client.DefaultRequestHeaders.UserAgent.ParseAdd('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36')
        foreach ($k in $Headers.Keys) { [void]$client.DefaultRequestHeaders.TryAddWithoutValidation($k, $Headers[$k]) }
        $resp = $client.GetAsync($Url).GetAwaiter().GetResult()
        if (-not $resp.IsSuccessStatusCode) { throw "HTTP $([int]$resp.StatusCode)" }
        return $resp.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
    } finally {
        $client.Dispose(); $handler.Dispose()
    }
}
function Invoke-HttpText {
    param([string]$Url, [string]$EncodingName = 'utf-8', [int]$TimeoutMs = 6000, [hashtable]$Headers = @{})
    $bytes = Invoke-HttpBytes -Url $Url -TimeoutMs $TimeoutMs -Headers $Headers
    return [System.Text.Encoding]::GetEncoding($EncodingName).GetString($bytes)
}

function To-DoubleOrZero([string]$s) {
    $d = 0.0
    if ([double]::TryParse($s, [ref]$d)) { return $d }
    return 0.0
}

# ---- market detection ------------------------------------------------
# Returns @{ Kind='US'|'ASH'|'HK'; Code; Tc; EmMarkets=int[]; Div=int; Cur; Sina }
function Resolve-Market([string]$raw) {
    $t = $raw.Trim().ToLower()
    if ($t -match '^(sh|sz|bj)([0-9]{6})$') {
        $ex = $Matches[1]; $num = $Matches[2]
        # 北交所: 腾讯 bj 前缀已验证可用 (qt.gtimg.cn/q=bj430047); 东财市场号按 0/1 兜底尝试
        $mk = if ($ex -eq 'sh') { @(1) } elseif ($ex -eq 'sz') { @(0) } else { @(0, 1) }
        return @{ Kind='ASH'; Code=$num; Tc="$ex$num"; EmMarkets=$mk; Div=100; Cur='CNY'; Sina="$ex$num"; Ex=$ex }
    }
    if ($t -match '^hk([0-9]{4,5})$') {
        $num = $Matches[1].PadLeft(5, '0')
        return @{ Kind='HK'; Code=$num; Tc="hk$num"; EmMarkets=@(116); Div=1000; Cur='HKD'; Sina="rt_hk$num"; Ex='hk' }
    }
    if ($t -match '^us([a-z][a-z0-9.\-]*)$') {
        $sym = $Matches[1].ToUpper()
        return @{ Kind='US'; Code=$sym; Tc="us$sym"; EmMarkets=@(105, 106, 107); Div=100; Cur='USD'; Sina="gb_$($sym.ToLower())"; Ex='us' }
    }
    if ($t -match '^([0-9]{6})$') {
        $num = $Matches[1]
        $ex = if ($num[0] -eq '6' -or $num[0] -eq '9') { 'sh' } elseif ($num[0] -eq '4' -or $num[0] -eq '8') { 'bj' } else { 'sz' }
        # 北交所: 腾讯 bj 前缀已验证可用 (qt.gtimg.cn/q=bj430047); 东财市场号按 0/1 兜底尝试
        $mk = if ($ex -eq 'sh') { @(1) } elseif ($ex -eq 'sz') { @(0) } else { @(0, 1) }
        return @{ Kind='ASH'; Code=$num; Tc="$ex$num"; EmMarkets=$mk; Div=100; Cur='CNY'; Sina="$ex$num"; Ex=$ex }
    }
    if ($t -match '^([0-9]{4,5})$') {
        $num = $Matches[1].PadLeft(5, '0')
        return @{ Kind='HK'; Code=$num; Tc="hk$num"; EmMarkets=@(116); Div=1000; Cur='HKD'; Sina="rt_hk$num"; Ex='hk' }
    }
    if ($t -match '^([a-z][a-z0-9.\-]*)$') {
        $sym = $t.ToUpper()
        return @{ Kind='US'; Code=$sym; Tc="us$sym"; EmMarkets=@(105, 106, 107); Div=100; Cur='USD'; Sina="gb_$($sym.ToLower())"; Ex='us' }
    }
    throw "无法识别的股票代码: $raw (支持 A股 600519 / 港股 00700 / 美股 PDD, 可加 sh/sz/bj/hk/us 前缀)"
}

$script:Market = Resolve-Market $Ticker

function New-Quote {
    param([string]$Name,[string]$Code,[double]$Price,[double]$Prev,[double]$Open,[double]$High,[double]$Low,
          [double]$Vol,[double]$Amt,[string]$QTime,[string]$Source)
    if ($Prev -le 0) { $Prev = 0 }
    $chg = 0.0; $pct = 0.0
    if ($Prev -gt 0) { $chg = $Price - $Prev; $pct = ($Price / $Prev - 1) * 100 }
    $up = 0
    if ($chg -gt 0.004) { $up = 1 } elseif ($chg -lt -0.004) { $up = -1 }
    return [pscustomobject]@{
        Name=$Name; Code=$Code; Price=$Price; PrevClose=$Prev; Open=$Open; High=$High; Low=$Low
        Volume=$Vol; Amount=$Amt; QuoteTime=$QTime; Source=$Source; Chg=$chg; Pct=$pct; Up=$up
        Currency=$script:Market.Cur; Market=$script:Market.Kind
    }
}

# ---- providers --------------------------------------------------------
function Get-Quote-Eastmoney {
    $errText = ''
    foreach ($mk in $script:Market.EmMarkets) {
        try {
            $url = "https://push2.eastmoney.com/api/qt/stock/get?secid=$mk.$($script:Market.Code)&fields=f43,f44,f45,f46,f47,f48,f57,f58,f60,f169,f170"
            $txt = Invoke-HttpText -Url $url -EncodingName 'utf-8' -TimeoutMs 5000
            $j = $txt | ConvertFrom-Json
            if ($j.rc -ne 0 -or $null -eq $j.data) { continue }
            $d = $j.data
            $div = [double]$script:Market.Div
            $price = [double]$d.f43 / $div
            if ($price -le 0) { continue }
            $vol = [double]$d.f47
            if ($script:Market.Kind -eq 'ASH') { $vol = $vol * 100 }   # 东财 A股成交量为"手", 统一归一为股
            return (New-Quote -Name ([string]$d.f58) -Code ([string]$d.f57) -Price $price `
                -Prev ([double]$d.f60/$div) -Open ([double]$d.f46/$div) -High ([double]$d.f44/$div) -Low ([double]$d.f45/$div) `
                -Vol $vol -Amt ([double]$d.f48) -QTime '' -Source 'eastmoney')
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
        return (New-Quote -Name $name -Code ([string]$f[2]) -Price $price `
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
    }
    return (New-Quote -Name $name -Code ([string]$f[2]) -Price $price `
        -Prev $prev -Open $open -High $high -Low $low `
        -Vol $vol -Amt $amt -QTime $qtime -Source 'tencent')
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
        $price = To-DoubleOrZero $f[2]
        $chg = To-DoubleOrZero $f[3]; $pct = To-DoubleOrZero $f[4]
        $prev = 0.0
        if ($price -gt 0 -and $pct -ne 0) { $prev = $price / (1 + $pct / 100) }
        if ($price -le 0) { throw 'invalid price' }
        return (New-Quote -Name ([string]$f[0]) -Code ([string]$f[1]) -Price $price -Prev $prev -Open 0 -High 0 -Low 0 `
            -Vol 0 -Amt 0 -QTime '' -Source 'sina')
    }
    throw 'sina not supported for HK'
}

function Emit([bool]$ok, $q, [string]$src, [string]$err) {
    $obj = @{ ok = $ok }
    if ($q) { $obj.q = $q; $obj.src = $src }
    if ($err) { $obj.err = $err }
    $json = $obj | ConvertTo-Json -Compress -Depth 5
    if ($Out) {
        try { [System.IO.File]::WriteAllText($Out, $json, (New-Object System.Text.UTF8Encoding($true))) }
        catch { try { Set-Content -Path $Out -Value $json -Encoding UTF8 } catch {} }
    } else {
        Write-Output $json
    }
    if ($ok) { exit 0 } else { exit 1 }
}

$errs = @()
foreach ($src in $Sources) {
    try {
        $q = switch ($src) {
            'eastmoney' { Get-Quote-Eastmoney }
            'tencent'   { Get-Quote-Tencent }
            'sina'      { Get-Quote-Sina }
            default     { throw "unknown source: $src" }
        }
        if ($null -ne $q -and $q.Price -gt 0) {
            Emit $true $q $src ''
        }
    } catch {
        $errs += "[$src] $($_.Exception.Message)"
    }
}
$errText = ($errs -join '; ')
if ($errText.Length -gt 180) { $errText = $errText.Substring(0, 180) }
Emit $false $null '' $errText
