$ErrorActionPreference = "Stop"

$depsDir = Join-Path $PSScriptRoot ".codex_test_deps"
python -m pip install --upgrade --target $depsDir -r requirements.txt
python -m pip install --upgrade --target $depsDir pyinstaller

$oldPythonPath = $env:PYTHONPATH
if ([string]::IsNullOrWhiteSpace($oldPythonPath)) {
    $env:PYTHONPATH = $depsDir
} else {
    $env:PYTHONPATH = $depsDir + ";" + $oldPythonPath
}

$appName = "Excel" + [char]0x8BA2 + [char]0x5355 + [char]0x6570 + [char]0x636E + [char]0x63D0 + [char]0x53D6 + [char]0x5DE5 + [char]0x5177 + "_v1.4"

python -m PyInstaller --noconfirm --clean --windowed --onedir --name $appName --hidden-import extract_orders --collect-all customtkinter extract_orders_gui.py

$env:PYTHONPATH = $oldPythonPath

$distDir = Join-Path ".\dist" $appName

if (Test-Path ".\README.md") {
    Copy-Item ".\README.md" $distDir -Force
}
if (Test-Path ".\category_config.json") {
    Copy-Item ".\category_config.json" $distDir -Force
}
if (Test-Path ".\app_settings.json") {
    Copy-Item ".\app_settings.json" $distDir -Force
}

Write-Host ""
Write-Host ("Build complete: " + $distDir)
Write-Host ("EXE path: " + (Join-Path $distDir ($appName + ".exe")))

if ($env:CODEX_NO_OPEN_EXPLORER -ne "1") {
    explorer $distDir
}
