
# This script creates a Python virtual environment using pyenv and sets it up for use in Visual Studio Code.
# It also installs common packages and generates a settings.json file for VS Code.
# Make sure to run this script in a PowerShell terminal with administrative privileges.


# You need to have pyenv installed and available in your PATH for this script to work.
if (-not (Get-Command pyenv -ErrorAction SilentlyContinue)) {
    Write-Error "pyenv is not installed or not available in your PATH. Please install pyenv and try again."
    exit 1
}

# Get the list of installed Python versions
$installedVersions = pyenv versions | ForEach-Object { $_.Trim() }
Write-Host "Installed Python versions:"

# if there are no Python versions installed, exit the script
if ($installedVersions.Count -eq 0) {
    Write-Error "No Python versions installed. Please install a Python version using pyenv and try again."
    exit 1
}

# Initialize the counter outside the loop
$i = 1
$installedVersions | ForEach-Object {
    Write-Host "[$i] $_"
    $i++
}

# Prompt the user to select a Python version by number
Write-Host "Enter the number corresponding to the desired Python version (starting from 1)"
$selection = Read-Host

# Validate the selection
if (-not ($selection -as [int]) -or $selection -lt 1 -or $selection -gt $installedVersions.Count) {
    Write-Error "Invalid selection. Please enter a number between 1 and $($installedVersions.Count)."
    exit 1
}

# Get the selected Python version
$PythonVersion = $installedVersions[$selection - 1]

# Confirm current project directory
$projectPath = Get-Location
Write-Host "Project directory: $projectPath"

# Check if the selected version is valid
if ($installedVersions -notcontains $PythonVersion) {
    Write-Error "Python version $PythonVersion is not installed. Please install it using pyenv and try again."
    exit 1
}

Write-Host "Using Python version $PythonVersion"

# Set local Python version
pyenv local $PythonVersion
if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to set local Python version to $PythonVersion. Exiting."
    exit 1
}
Write-Host "Set local Python version to $PythonVersion"

# Create virtual environment
Write-Host "Creating virtual environment in .venv..."
if (Test-Path ".venv") {
    Write-Host "The .venv folder already exists. Please remove it manually first to avoid deleting any existing virtual environment."
    Write-Host "If you want to remove the existing virtual environment, run the following command:"
    Write-Host "Remove-Item -Recurse -Force .venv"
    exit 1
}
python -m venv .venv
if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to create virtual environment. Exiting."
    exit 1
}

# Generate VS Code settings
$vsCodeFolder = ".vscode"
if (-not (Test-Path $vsCodeFolder)) {
    New-Item -ItemType Directory -Path $vsCodeFolder | Out-Null
}

$settingsPath = Join-Path $vsCodeFolder "settings.json"
$interpreterPath = '${workspaceFolder}\\.venv\\Scripts\\python.exe'

# Create or update settings.json
if (Test-Path $settingsPath) {
    try {
        $existingSettings = Get-Content -Path $settingsPath -Raw | ConvertFrom-Json
        $existingSettings["python.defaultInterpreterPath"] = $interpreterPath
        $settingsJson = $existingSettings | ConvertTo-Json -Depth 10
    } catch {
        Write-Host "Could not parse existing settings.json. Creating new file."
        $settingsMap = @{ "python.defaultInterpreterPath" = $interpreterPath }
        $settingsJson = $settingsMap | ConvertTo-Json -Depth 10
    }
} else {
    $settingsMap = @{ "python.defaultInterpreterPath" = $interpreterPath }
    $settingsJson = $settingsMap | ConvertTo-Json -Depth 10
}

$settingsJson | Out-File -FilePath $settingsPath -Encoding UTF8 -Force

# Activate the virtual environment and install common packages
Write-Host "Activating virtual environment and installing base packages..."
& .\.venv\Scripts\Activate.ps1

# remove the .python-version file if it exists
if (Test-Path ".python-version") {
    Remove-Item ".python-version"
}