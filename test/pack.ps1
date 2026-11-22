# 打发布包：把运行需要的文件装进一个 zip
# 用法： powershell -NoProfile -ExecutionPolicy Bypass -File test\pack.ps1
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here

$ver = (Select-String -Path (Join-Path $root "core.js") -Pattern "VERSION = '([\d.]+)'").Matches[0].Groups[1].Value
if (-not $ver) { throw "读不到 core.js 里的版本号" }

$staging = Join-Path $env:TEMP ("letter-pack-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $staging -Force | Out-Null

$files = @(
    "index.html", "app.js", "app.css", "core.js",
    "vendor\xlsx.full.min.js",
    "README.md", "LICENSE"
)

try {
    foreach ($f in $files) {
        $src = Join-Path $root $f
        if (-not (Test-Path $src)) { throw "缺文件：$f" }
        $dst = Join-Path $staging $f
        $dir = Split-Path -Parent $dst
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Copy-Item $src $dst -Force
    }

    $zip = Join-Path (Split-Path -Parent $root) ("底稿填充助手-v" + $ver + ".zip")
    if (Test-Path $zip) { Remove-Item $zip -Force }
    Compress-Archive -Path (Join-Path $staging "*") -DestinationPath $zip -CompressionLevel Optimal

    $item = Get-Item $zip
    $hash = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()
    Write-Host ""
    Write-Host ("发布包： " + $item.FullName)
    Write-Host ("大小：   " + [math]::Round($item.Length / 1KB, 1) + " KB (" + $item.Length + " 字节)")
    Write-Host ("SHA256： " + $hash)
    Write-Host ("文件数： " + $files.Count)
    Write-Host ""
    Write-Host "解压后双击 index.html 即可使用。"
} finally {
    if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
}
