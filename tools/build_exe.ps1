param(
    [ValidateSet("OneFile", "Clean")]
    [string]$Mode = "OneFile"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

if ($Mode -eq "Clean") {
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue build, dist
    Write-Host "Removed build and dist folders."
    exit 0
}

python .\tools\generate_icon.py
python -m PyInstaller .\SoundManager.spec --noconfirm --clean

$ExePath = Join-Path $ProjectRoot "dist\SoundManager.exe"
if (Test-Path $ExePath) {
    Write-Host "Built $ExePath"
} else {
    throw "Build finished, but dist\SoundManager.exe was not found."
}
