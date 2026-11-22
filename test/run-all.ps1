# 底稿填充助手 · 全量回归
# 用法： powershell -NoProfile -ExecutionPolicy Bypass -File test\run-all.ps1
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

$localPaths = Join-Path $here "local-paths.ps1"
if (Test-Path $localPaths) { . $localPaths }

function Find-Exe($envName, $cmd) {
    if ($envName -and (Test-Path $envName)) { return $envName }
    $c = Get-Command $cmd -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    throw "找不到 $cmd：装到 PATH，或设置环境变量，或在 test\local-paths.ps1 里写绝对路径"
}

$node = Find-Exe $env:DSH_NODE "node"
Write-Host "node: $node"
Write-Host ""

$suites = @("core", "fixes", "limits", "export", "ui-smoke")
$total = 0
$failed = @()

foreach ($s in $suites) {
    $file = Join-Path $here "$s.test.js"
    $out = & {
        # Windows PowerShell turns redirected stderr into error records.
        $ErrorActionPreference = "Continue"
        $PSNativeCommandUseErrorActionPreference = $false
        & $node $file 2>&1
    }
    $code = $LASTEXITCODE
    $line = ($out | Where-Object { $_ -match "全部通过|失败 |跳过" }) -join " "
    if ($code -ne 0) {
        $failed += $s
        Write-Host ("  {0,-10} 失败" -f $s) -ForegroundColor Red
        $out | ForEach-Object {
            $message = if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { $_ }
            Write-Host "    $message"
        }
    } else {
        if ($line -match "(\d+) 项断言") { $total += [int]$Matches[1] }
        Write-Host ("  {0,-10} {1}" -f $s, $line)
    }
}

Write-Host ""
if ($failed.Count) {
    Write-Host ("失败套件：" + ($failed -join ", ")) -ForegroundColor Red
    exit 1
}
Write-Host ("全部通过，共 " + $total + " 项断言") -ForegroundColor Green
