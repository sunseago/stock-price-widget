# =====================================================================
#  tests\Test-Widget.ps1
#  不依赖 Pester 的自检脚本, 用目标运行时 Windows PowerShell 5.1 跑:
#    powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests\Test-Widget.ps1
#  覆盖: A股交易所推断 / 市场识别 / 数值格式化(区域无关) / 假期表跨年检测
#        以及 Draw-Card 的离屏渲染(输出 PNG 供人工核对排版)
# =====================================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$widget = Join-Path $root 'PDD-Widget.ps1'
$renderDir = Join-Path $env:TEMP 'spw-render'
if (-not (Test-Path -LiteralPath $renderDir)) { New-Item -ItemType Directory -Path $renderDir | Out-Null }

$script:Pass = 0; $script:Fail = 0
function Check([string]$name, $actual, $expected) {
    if ([string]$actual -eq [string]$expected) {
        $script:Pass++
        Write-Host ("  ok    {0} = {1}" -f $name, $actual)
    } else {
        $script:Fail++
        Write-Host ("  FAIL  {0}: got [{1}] want [{2}]" -f $name, $actual, $expected) -ForegroundColor Red
    }
}

Write-Host '[1] 加载 PDD-Widget.ps1 (-NoRun)'
. $widget -NoRun
Check 'loaded' $true $true

Write-Host '[2] A股交易所推断 (920 段/沪B/深市)'
Check '600519 -> sh' (Get-AShareExchange '600519') 'sh'
Check '688981 -> sh' (Get-AShareExchange '688981') 'sh'
Check '900901 -> sh(沪B)' (Get-AShareExchange '900901') 'sh'
Check '000001 -> sz' (Get-AShareExchange '000001') 'sz'
Check '300750 -> sz' (Get-AShareExchange '300750') 'sz'
Check '920047 -> bj' (Get-AShareExchange '920047') 'bj'
Check '430047 -> bj' (Get-AShareExchange '430047') 'bj'
Check '830799 -> bj' (Get-AShareExchange '830799') 'bj'

Write-Host '[3] 市场识别'
Check '920047 kind' (Get-MarketKind '920047') 'ASH'
Check 'bj920047 kind' (Get-MarketKind 'bj920047') 'ASH'
Check '00700 kind' (Get-MarketKind '00700') 'HK'
Check 'PDD kind' (Get-MarketKind 'PDD') 'US'

Write-Host '[4] 数值格式化不随系统区域漂移 (de-DE 小数点是逗号)'
$old = [System.Threading.Thread]::CurrentThread.CurrentCulture
try {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    Check 'Format-Price 1435.26' (Format-Price 1435.26) '1,435.26'
    Check 'Format-Vol 4178000' (Format-Vol 4178000) '417.80万'
    Check 'Format-Amt 5.95e9 CNY' (Format-Amt 5954830000 'CNY') '59.55亿元'
} finally {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $old
}

Write-Host '[5] 假期表跨年检测 (内置表只到 2026)'
Check 'ASH 2026 covered' (Test-HolidayTableFresh 'ASH' 2026) $true
Check 'ASH 2027 covered?' (Test-HolidayTableFresh 'ASH' 2027) $false
Check 'HK 2027 covered?' (Test-HolidayTableFresh 'HK' 2027) $false

Write-Host '[6] 成交额/涨跌方向'
$q = [pscustomobject]@{
    Name = '贵州茅台'; Code = '600519'; Price = 1435.26; PrevClose = 1410.27
    Open = 1408.00; High = 1442.50; Low = 1401.10; Volume = 4178000; Amount = 5954830000
    QuoteTime = '2026-11-13 16:14:55'; Chg = 24.99; Pct = 1.7724; Up = 1; Stale = $false
    Currency = 'CNY'; Market = 'ASH'; Warn = ''
}

function Render-Card([double]$scale, [string]$outFile) {
    $script:Scale = $scale
    Initialize-Assets
    $w = [int](350 * $scale); $h = [int](152 * $scale)
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $bmp.SetResolution([single](96 * $scale), [single](96 * $scale))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.Clear([System.Drawing.Color]::FromArgb(255, 40, 40, 40))
        Draw-Card $g $w $h
    } finally { $g.Dispose() }
    $bmp.Save($outFile, [System.Drawing.Imaging.ImageFormat]::Png)
    $size = (Get-Item -LiteralPath $outFile).Length
    $bmp.Dispose()
    return $size
}

Write-Host '[7] 离屏渲染 (排版自检, PNG 输出到 ' + $renderDir + ')'
$script:Ticker = '600519'; $script:StockName = '贵州茅台'
$script:Quote = $q; $script:Cfg.Ticker = '600519'; $script:Cfg.StockName = '贵州茅台'
$script:DemoOn = $false; $script:Fetching = $false; $script:FetchErr = ''
$script:SourceUsed = '腾讯'; $script:QuoteAt = Get-Date; $script:SesCache = $null
Check 'render 100% 涨' ((Render-Card 1.0 (Join-Path $renderDir 'a-96-up.png')) -gt 1000) $true

# 大跌 + 长名称(港股) — 检查长名字是否压到徽章
$q2 = $q | Select-Object *
$q2.Name = '阿里巴巴-W(港股通)'; $q2.Code = '09988'; $q2.Price = 118.4; $q2.PrevClose = 121.9
$q2.Chg = -3.5; $q2.Pct = -2.8712; $q2.Up = -1; $q2.Currency = 'HKD'; $q2.Open = 0; $q2.High = 0; $q2.Low = 0
$q2.Volume = 0; $q2.Amount = 0; $q2.QuoteTime = ''
$script:Ticker = '09988'; $script:StockName = '阿里巴巴-W'; $script:Cfg.Ticker = '09988'
$script:Quote = $q2
Check 'render 100% 跌+长名' ((Render-Card 1.0 (Join-Path $renderDir 'b-96-down.png')) -gt 1000) $true

# 失败态: 只有 err, 没有行情
$script:Quote = $null
$script:FetchErr = '[tencent] timeout after 4000ms; [eastmoney] eastmoney no quote for 920047 ()'
Check 'render 失败态' ((Render-Card 1.0 (Join-Path $renderDir 'c-96-error.png')) -gt 1000) $true

# 北交所老代码告警
$q3 = $q | Select-Object *
$q3.Code = '430047'; $q3.Name = '诺思兰德'; $q3.Price = 8.17; $q3.PrevClose = 8.17; $q3.Chg = 0; $q3.Pct = 0
$q3.Up = 0; $q3.Volume = 0; $q3.Amount = 0; $q3.Warn = '量额为 0, 可能已停用/迁移至 92xxxx'
$script:Quote = $q3; $script:Ticker = '430047'; $script:StockName = '诺思兰德'; $script:Cfg.Ticker = '430047'
$script:FetchErr = ''
Check 'render 北交所告警' ((Render-Card 1.0 (Join-Path $renderDir 'd-96-warn.png')) -gt 1000) $true

# 150% 缩放
$script:Quote = $q; $script:Ticker = '600519'; $script:Cfg.Ticker = '600519'
Check 'render 150%' ((Render-Card 1.5 (Join-Path $renderDir 'e-144-up.png')) -gt 1000) $true

Write-Host ''
Write-Host ("结果: {0} 通过 / {1} 失败" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 } else { exit 0 }
