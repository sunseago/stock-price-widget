# =====================================================================
#  build.ps1 - 一键重打包 dist\股价悬浮窗.exe
#  把 PDD-Widget.ps1 / PDD-Fetch.ps1 / src\app.ico 内嵌进单文件 EXE。
#  用 .NET Framework 自带的 csc, 不需要配置 PATH 或安装 SDK。
#  用法:  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
# =====================================================================
$ErrorActionPreference = 'Stop'
$root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }

$cscCandidates = @(
    (Join-Path $env:SystemRoot 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'),
    (Join-Path $env:SystemRoot 'Microsoft.NET\Framework\v4.0.30319\csc.exe')
)
$csc = $cscCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $csc) { throw 'csc.exe 不存在: 需要 .NET Framework 4.x (Windows 10/11 自带)' }

$widget = Join-Path $root 'PDD-Widget.ps1'
$fetch = Join-Path $root 'PDD-Fetch.ps1'
$icon = Join-Path $root 'src\app.ico'
$launcher = Join-Path $root 'src\Launcher.cs'
$assembly = Join-Path $root 'src\AssemblyInfo.cs'
foreach ($f in $widget, $fetch, $icon, $launcher, $assembly) {
    if (-not (Test-Path -LiteralPath $f)) { throw ('缺少文件: ' + $f) }
}

$outDir = Join-Path $root 'dist'
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
$out = Join-Path $outDir '股价悬浮窗.exe'

$cscArgs = @(
    '/nologo', '/target:winexe', '/platform:anycpu', '/optimize+',
    ('/out:' + $out),
    ('/win32icon:' + $icon),
    ('/resource:' + $widget + ',PDD-Widget.ps1'),
    ('/resource:' + $fetch + ',PDD-Fetch.ps1'),
    '/r:System.dll', '/r:System.Core.dll',
    $launcher, $assembly
)
& $csc @cscArgs
if ($LASTEXITCODE -ne 0) { throw ('csc 失败, exit=' + $LASTEXITCODE) }

$size = (Get-Item -LiteralPath $out).Length
$sha = (Get-FileHash -LiteralPath $out -Algorithm SHA256).Hash
"OK  $out"
"    $size bytes"
"    SHA256 $sha"
