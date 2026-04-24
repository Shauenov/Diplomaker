$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
$specFile = Join-Path $PSScriptRoot "desktop_app.spec"

function Invoke-Step {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Description,
        [Parameter(Mandatory = $true)]
        [scriptblock]$Script
    )

    & $Script
    if ($LASTEXITCODE -ne 0) {
        throw "$Description failed with exit code $LASTEXITCODE"
    }
}

if (-not (Test-Path $venvPython)) {
    throw "Python executable not found: $venvPython"
}

Push-Location $projectRoot
try {
    if (-not (Test-Path (Join-Path $projectRoot "build"))) {
        New-Item -ItemType Directory -Path (Join-Path $projectRoot "build") | Out-Null
    }

    if (-not (Test-Path (Join-Path $projectRoot "dist"))) {
        New-Item -ItemType Directory -Path (Join-Path $projectRoot "dist") | Out-Null
    }

    Invoke-Step "pip upgrade" { & $venvPython -m pip install --upgrade pip }
    Invoke-Step "PyInstaller install" { & $venvPython -m pip install pyinstaller }
    Invoke-Step "PyInstaller build" { & $venvPython -m PyInstaller --noconfirm --clean $specFile }

    Write-Host "Portable build created in dist\DiplomaGenerator" -ForegroundColor Green
}
finally {
    Pop-Location
}
