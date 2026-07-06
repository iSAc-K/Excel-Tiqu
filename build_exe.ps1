$ErrorActionPreference = "Stop"

$depsDir = Join-Path $PSScriptRoot ".codex_test_deps"
if ($env:SKIP_DEP_INSTALL -ne "1") {
    python -m pip install --upgrade --target $depsDir -r requirements.txt
    python -m pip install --upgrade --target $depsDir pyinstaller
}

$oldPythonPath = $env:PYTHONPATH
if ([string]::IsNullOrWhiteSpace($oldPythonPath)) {
    $env:PYTHONPATH = $depsDir
} else {
    $env:PYTHONPATH = $depsDir + ";" + $oldPythonPath
}

$versionLine = Get-Content -LiteralPath (Join-Path $PSScriptRoot "VERSION.txt") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
if ([string]::IsNullOrWhiteSpace($versionLine)) {
    throw "VERSION.txt is missing a version line"
}
$version = $versionLine.Trim()
if ($version.StartsWith("v", [System.StringComparison]::OrdinalIgnoreCase)) {
    $version = $version.Substring(1)
}
$appName = "Excel" + [char]0x8BA2 + [char]0x5355 + [char]0x6570 + [char]0x636E + [char]0x63D0 + [char]0x53D6 + [char]0x5DE5 + [char]0x5177 + "_v" + $version

python -m PyInstaller --noconfirm --clean --windowed --onedir --name $appName --hidden-import extract_orders --collect-all customtkinter extract_orders_gui.py

$updaterName = "updater"
python -m PyInstaller --noconfirm --clean --windowed --onedir --name $updaterName --collect-all customtkinter updater.py

$env:PYTHONPATH = $oldPythonPath

$distDir = Join-Path ".\dist" $appName
$updaterDist = Join-Path ".\dist" $updaterName
$updaterExe = Join-Path $updaterDist "updater.exe"
if (-not (Test-Path $updaterExe)) {
    throw "updater.exe was not built: $updaterExe"
}
Copy-Item $updaterExe (Join-Path $distDir "updater.exe") -Force

if (Test-Path ".\README.md") {
    Copy-Item ".\README.md" $distDir -Force
}
if (Test-Path ".\category_config.json") {
    Copy-Item ".\category_config.json" $distDir -Force
}
if (Test-Path ".\app_settings.json") {
    Copy-Item ".\app_settings.json" $distDir -Force
}
if (Test-Path ".\VERSION.txt") {
    Copy-Item ".\VERSION.txt" $distDir -Force
}

if ($env:BUILD_RELEASE_ZIP -eq "1") {
    $zipName = "Excel-Tiqu-v$version.zip"
    $zipPath = Join-Path ".\dist" $zipName
    if (Test-Path $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }
    Compress-Archive -LiteralPath $distDir -DestinationPath $zipPath -Force
    $sha = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $manifest = [ordered]@{
        version = $version
        download_url = "https://github.com/iSAc-K/Excel-Tiqu/releases/download/v$version/$zipName"
        sha256 = $sha
        notes = @("Added automatic update support.")
    }
    $manifestPath = Join-Path ".\dist" "update.json"
    $manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    Write-Host ("Release ZIP: " + $zipPath)
    Write-Host ("Manifest: " + $manifestPath)
}

Write-Host ""
Write-Host ("Build complete: " + $distDir)
Write-Host ("EXE path: " + (Join-Path $distDir ($appName + ".exe")))

if ($env:CODEX_NO_OPEN_EXPLORER -ne "1") {
    explorer $distDir
}
