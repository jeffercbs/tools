param(
    [switch]$Clean,
    [switch]$Console,
    [switch]$Run,
    [switch]$RecreateVenv,
    [switch]$NoInstall
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $root

function Write-Step($msg) { Write-Host "`n==== $msg ====" -ForegroundColor Cyan }
function Try-StopProcess($name) {
    try {
        $p = Get-Process -Name $name -ErrorAction SilentlyContinue
        if ($null -ne $p) {
            Write-Host "Stopping process: $name (PID(s): $($p.Id -join ', '))" -ForegroundColor Yellow
            $p | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 300
        }
    } catch {}
}
function Try-Delete($path, $retries=10) {
    for ($i=0; $i -lt $retries; $i++) {
        try { if (Test-Path $path) { Remove-Item -Force $path -ErrorAction Stop } return }
        catch { Start-Sleep -Milliseconds 300 }
    }
    if (Test-Path $path) { Write-Host "Warning: could not delete $path (locked)" -ForegroundColor Yellow }
}

# Resolve Python in venv (create if missing)
$venvPython = Join-Path $root 'venv\Scripts\python.exe'
if ($RecreateVenv -and (Test-Path (Join-Path $root 'venv'))) {
    Write-Step 'Removing existing venv'
    Remove-Item -Recurse -Force (Join-Path $root 'venv')
}
if (-not (Test-Path $venvPython)) {
    Write-Step 'Creating virtual environment (venv)'
    if (Get-Command py -ErrorAction SilentlyContinue) {
        py -3 -m venv "$root\venv"
    } else {
        python -m venv "$root\venv"
    }
}

# Install/upgrade dependencies
if (-not $NoInstall) {
    Write-Step 'Upgrading pip'
    & $venvPython -m pip install --upgrade pip

    $requirements = Join-Path $root 'requirements.txt'
    if (Test-Path $requirements) {
        Write-Step 'Installing from requirements.txt'
        & $venvPython -m pip install -r $requirements
    } else {
        Write-Step 'Installing minimal build/runtime deps (PyInstaller, PyMuPDF, openpyxl)'
        & $venvPython -m pip install PyInstaller PyMuPDF openpyxl
    }
}

# Optional clean build artifacts
if ($Clean) {
    Write-Step 'Cleaning build/ and dist/'
    if (Test-Path (Join-Path $root 'build')) { Remove-Item -Recurse -Force (Join-Path $root 'build') }
    if (Test-Path (Join-Path $root 'dist')) { Remove-Item -Recurse -Force (Join-Path $root 'dist') }
}

# Ensure PyMuPDF is importable in the venv
Write-Step 'Sanity check: importing fitz (PyMuPDF) in venv'
& $venvPython -c "import fitz; print('fitz OK', fitz.__doc__[:20])"

# Choose spec file; optionally generate console variant
$specPath = Join-Path $root 'ProcesarPDF.spec'
if (-not (Test-Path $specPath)) {
    throw "Spec file not found: $specPath"
}
$specToUse = $specPath
if ($Console) {
    Write-Step 'Generating console spec variant'
    $consoleSpec = Join-Path $root 'ProcesarPDF.console.spec'
    $content = Get-Content $specPath -Raw
    $content = $content -replace "console=False", "console=True"
    $content = $content -replace "name='ProcesarPDF'", "name='ProcesarPDFConsole'"
    Set-Content -Path $consoleSpec -Value $content -Encoding UTF8
    $specToUse = $consoleSpec
}

# Build
Write-Step "Building with PyInstaller spec: $(Split-Path -Leaf $specToUse)"

# Ensure previous EXEs are not holding a lock
Try-StopProcess 'ProcesarPDF'
Try-StopProcess 'ProcesarPDFConsole'

$distDir = Join-Path $root 'dist'
$guiExe = Join-Path $distDir 'ProcesarPDF.exe'
$consoleExe = Join-Path $distDir 'ProcesarPDFConsole.exe'
Try-Delete $guiExe
Try-Delete $consoleExe

& $venvPython -m PyInstaller --noconfirm $specToUse

# Locate output
$exeName = if ($Console) { 'ProcesarPDFConsole.exe' } else { 'ProcesarPDF.exe' }
$exePath = Join-Path $distDir $exeName

if (Test-Path $exePath) {
    Write-Host "`nBuild complete: $exePath" -ForegroundColor Green
} else {
    Write-Host "`nBuild finished but expected EXE not found in dist/. Check PyInstaller output." -ForegroundColor Yellow
}

if ($Run -and (Test-Path $exePath)) {
    Write-Step 'Launching EXE'
    & $exePath
}
