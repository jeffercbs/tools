$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$venvPath = Join-Path $scriptDir "venv"
$scriptPath = Join-Path $scriptDir "main.py"
$requirementsPath = Join-Path $scriptDir "requirements.txt"
$allInstalled = $true

if (-not (Test-Path $venvPath)) {
    Write-Host "Creating virtual environment..." -ForegroundColor Yellow
    python -m venv $venvPath
}

Write-Host "Environment created successfully" -ForegroundColor Green

if (-not(Test-Path $scriptDir\renombrados)) {
    Write-Host "Create directory for saving files renamed..." -ForegroundColor Yellow
    mkdir $scriptDir\renombrados 
}

Write-Host "Directory renombrados created successfully" -ForegroundColor Green

Write-Host "validating if the excel file exists..." -ForegroundColor Yellow

if (-not (Test-Path $scriptDir\soportes.xlsx)) {
    Write-Host "The file soportes.xlsx does not exist, please add it to the root of the project" -ForegroundColor Red
    exit
}

Write-Host "Activating virtual environment..." -ForegroundColor Yellow
Write-Host "venvPath: $venvPath"
. $venvPath\Scripts\Activate.ps1
Write-Host "Virtual environment activated" -ForegroundColor Green

foreach ($line in Get-Content $requirementsPath) {
    $package = $line.Split('==')[0]
    $installed = pip show $package

    if (-not $installed) {
        Write-Host "Installing $package..." -ForegroundColor Yellow
        pip install $line
        $allInstalled = $false
    }

}

if ($allInstalled) {
    Write-Host "All packages are already installed" -ForegroundColor Green
}
else {
    Write-Host "All packages installed successfully" -ForegroundColor Green
    pip install -r $requirementsPath
}

Write-Host "Running script..." -ForegroundColor Yellow
python $scriptPath

Write-Host "Deactivating virtual environment..."
deactivate