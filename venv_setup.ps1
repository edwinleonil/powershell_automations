
#Requires -Version 5.1
<#
.SYNOPSIS
    Creates a Python virtual environment with VS Code integration.

.DESCRIPTION
    This script creates a Python virtual environment using either the pyenv tool (if available) 
    or the system Python. It sets up VS Code settings and provides a robust, 
    user-friendly experience with comprehensive error handling.

.PARAMETER Force
    Force recreation of virtual environment if it already exists.

.EXAMPLE
    .\env_gen.ps1
    .\env_gen.ps1 -Force
#>

param(
    [switch]$Force
)

# Set strict mode for better error handling 
Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

# Function to write colored output
function Write-Info {
    param([string]$Message)
    Write-Host "✓ $Message" -ForegroundColor Green
}

function Write-Warning {
    param([string]$Message)
    Write-Host "⚠ $Message" -ForegroundColor Yellow
}

function Write-ErrorMsg {
    param([string]$Message)
    Write-Host "✗ $Message" -ForegroundColor Red
}

# Function to safely execute commands with error handling
function Invoke-SafeCommand {
    param(
        [string]$Command,
        [string]$Arguments = "",
        [string]$ErrorMessage = "Command failed"
    )
    
    try {
        if ($Arguments) {
            $argArray = $Arguments.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
            $result = & $Command @argArray
        } else {
            $result = & $Command
        }
        
        if ($LASTEXITCODE -ne 0) {
            throw "$ErrorMessage. Exit code: $LASTEXITCODE"
        }
        return $result
    }
    catch {
        Write-ErrorMsg "$ErrorMessage : $_"
        exit 1
    }
}

# Function to get Python executable
function Get-PythonExecutable {
    $allInstalledVersions = @()
    $allVersionDetails = @{}
    $sourceCounter = 1
    
    # Try py launcher first (Windows)
    if (Get-Command py -ErrorAction SilentlyContinue) {
        Write-Info "Found Python Launcher. Getting available Python versions..."
        
        try {
            # Try py -0 first (more detailed output)
            $pyOutput = $null
            try {
                $pyOutput = py -0 2>&1
                Write-Host "Using 'py -0' for version discovery..." -ForegroundColor Gray
            } catch {
                # Fallback to py --list
                try {
                    $pyOutput = py --list 2>&1
                    Write-Host "Using 'py --list' for version discovery..." -ForegroundColor Gray
                } catch {
                    Write-Warning "Both 'py -0' and 'py --list' failed"
                    $pyOutput = $null
                }
            }
            
            if ($pyOutput) {
                $installedVersions = @()
                $versionDetails = @{}

                # Ensure $pyOutput is always an array of lines
                if ($pyOutput -is [string]) {
                    $pyOutput = $pyOutput -split "`r?`n"
                }

                foreach ($line in $pyOutput) {
                    # Parse py -0 output - handle both old and new formats
                    # New format: " -V:3.13          Python 3.13 (64-bit)"
                    # Old format: " -3.11-64          C:\Python311\python.exe *"
                    # Active venv: "  *               Active venv"

                    # Skip active venv line
                    if ($line -match '^\s*\*\s+Active venv') {
                        continue
                    }

                    # First try new format with -V: prefix
                    if ($line -match '^\s*-V:([0-9]+\.[0-9]+)\s*(\*)?\s*Python\s+[0-9.]+\s+\((\d+)-bit\)\s*(\*?)') {
                        $version = $matches[1]
                        $architecture = $matches[3]
                        $isDefault = ($matches[2] -eq '*') -or ($matches[4] -eq '*')  # Check both positions for *
                        $pythonPath = ""

                        $versionKey = "py-$version-$architecture"
                        if ($installedVersions -notcontains $versionKey) {
                            $installedVersions += $versionKey
                            $versionDetails[$versionKey] = @{
                                Version = $version
                                Architecture = $architecture
                                Path = $pythonPath
                                IsDefault = $isDefault
                                Command = "py -$version"
                                Source = "Python Launcher"
                            }
                        }
                    }
                    # Then try old format for backwards compatibility
                    elseif ($line -match '^\s*-([0-9]+\.[0-9]+)(?:-(\d+))?\s*(.*)') {
                        $version = $matches[1]
                        $architecture = if ($matches[2]) { $matches[2] } else { "64" }
                        $pathAndDefault = $matches[3].Trim()
                        $isDefault = $pathAndDefault -match '\*'

                        # Extract path if available (from py -0)
                        $pythonPath = ""
                        if ($pathAndDefault -match '([A-Z]:\\[^*]+\.exe)') {
                            $pythonPath = $matches[1].Trim()
                        }

                        $versionKey = "py-$version-$architecture"
                        if ($installedVersions -notcontains $versionKey) {
                            $installedVersions += $versionKey
                            $versionDetails[$versionKey] = @{
                                Version = $version
                                Architecture = $architecture
                                Path = $pythonPath
                                IsDefault = $isDefault
                                Command = "py -$version"
                                Source = "Python Launcher"
                            }
                        }
                    }
                }

                # Add py launcher versions to combined list
                foreach ($versionKey in $installedVersions) {
                    $allInstalledVersions += $versionKey
                    $allVersionDetails[$versionKey] = $versionDetails[$versionKey]
                }
                
                Write-Info "Found $($installedVersions.Count) Python version(s) via Python Launcher"
            }
        }
        catch {
            Write-Warning "Error with py launcher: $_"
        }
    }

    # Try pyenv second
    if (Get-Command pyenv -ErrorAction SilentlyContinue) {
        Write-Info "Found pyenv. Getting available Python versions..."
        
        try {
            # Test if pyenv is working properly by checking for common Windows Script Host errors
            $pyenvTest = pyenv versions --bare 2>&1
            $pyenvOutput = $pyenvTest | Out-String
            
            # Check for cscript errors or other Windows Script Host issues
            if ($pyenvOutput -match 'cscript.*is not recognized' -or $pyenvOutput -match 'Windows Script Host' -or $LASTEXITCODE -ne 0) {
                throw "pyenv appears to be misconfigured (Windows Script Host issue)"
            }
            
            $allVersions = $pyenvTest
            $installedVersions = @()
            foreach ($version in $allVersions) {
                if (($version -notlike "*system*") -and ($version -notlike "*envs/*") -and ($version -notlike "*cscript*") -and ($version.Trim() -ne "")) {
                    $cleanVersion = $version.Trim()
                    $versionKey = "pyenv-$cleanVersion"
                    if ($installedVersions -notcontains $versionKey) {
                        $installedVersions += $versionKey
                        $allInstalledVersions += $versionKey
                        $allVersionDetails[$versionKey] = @{
                            Version = $cleanVersion
                            Architecture = ""
                            Path = ""
                            IsDefault = $false
                            Command = "python"  # Will use pyenv's python after setting local version
                            Source = "pyenv"
                            PyenvVersion = $cleanVersion
                        }
                    }
                }
            }
            
            Write-Info "Found $($installedVersions.Count) Python version(s) via pyenv"
        }
        catch {
            Write-Warning "Error with pyenv: $_"
        }
    }

    # If we found versions from any source, present them all
    if ($allInstalledVersions.Count -gt 0) {
        # Sort versions by source, then by version
        $allInstalledVersions = $allInstalledVersions | Sort-Object {
            $details = $allVersionDetails[$_]
            $sourceOrder = if ($details.Source -eq "Python Launcher") { 0 } else { 1 }
            $versionParts = $details.Version -split '\.'
            $majorMinor = [int]$versionParts[0] * 100 + [int]$versionParts[1]
            "$sourceOrder-$(-$majorMinor)"  # Sort by source first, then version descending
        }
        
        Write-Host "`nAvailable Python versions:"
        for ($i = 0; $i -lt $allInstalledVersions.Count; $i++) {
            $versionKey = $allInstalledVersions[$i]
            $details = $allVersionDetails[$versionKey]
            $defaultMarker = if ($details.IsDefault) { " (default)" } else { "" }
            $archInfo = if ($details.Architecture) { " ($($details.Architecture)-bit)" } else { "" }
            $pathInfo = if ($details.Path) { " - $($details.Path)" } else { "" }
            Write-Host "[$($i + 1)] Python $($details.Version)$archInfo$defaultMarker - via $($details.Source)$pathInfo"
        }
        
        do {
            $selection = Read-Host "`nSelect Python version (1-$($allInstalledVersions.Count)) or press Enter for default"
            
            if ([string]::IsNullOrWhiteSpace($selection)) {
                # Find default version or use first
                $defaultVersion = $allInstalledVersions | Where-Object { $allVersionDetails[$_].IsDefault } | Select-Object -First 1
                $selectedVersionKey = if ($defaultVersion) { $defaultVersion } else { $allInstalledVersions[0] }
                break
            }
            
            if ($selection -match '^\d+$') {
                $selectionNum = [int]$selection
                if (($selectionNum -ge 1) -and ($selectionNum -le $allInstalledVersions.Count)) {
                    $selectedVersionKey = $allInstalledVersions[$selection - 1]
                    break
                }
            }
            
            Write-Warning "Invalid selection. Please enter a number between 1 and $($allInstalledVersions.Count)."
        } while ($true)
        
        $selectedDetails = $allVersionDetails[$selectedVersionKey]
        
        # Handle pyenv selection differently
        if ($selectedDetails.Source -eq "pyenv") {
            Write-Info "Setting local Python version to $($selectedDetails.PyenvVersion) via pyenv"
            try {
                & pyenv local $selectedDetails.PyenvVersion
                if ($LASTEXITCODE -ne 0) {
                    throw "Failed to set Python version"
                }
            } catch {
                Write-ErrorMsg "Failed to set Python version: $_"
                exit 1
            }
        }
        
        $archText = if ($selectedDetails.Architecture) { " ($($selectedDetails.Architecture)-bit)" } else { "" }
        Write-Info "Using Python $($selectedDetails.Version)$archText via $($selectedDetails.Source)"
        return $selectedDetails.Command
    }
    
    # Try system Python last
    $pythonCommands = @("python", "python3", "py")
    foreach ($cmd in $pythonCommands) {
        if (Get-Command $cmd -ErrorAction SilentlyContinue) {
            try {
                $version = & $cmd --version 2>&1
                Write-Info "Found system Python: $version"
                return $cmd
            }
            catch {
                continue
            }
        }
    }
    
    Write-ErrorMsg "No Python installation found. Please install Python or use the py launcher."
    exit 1
}

# Function to handle existing virtual environment
function Test-VirtualEnvironment {
    if (Test-Path ".venv") {
        if ($Force) {
            Write-Warning "Removing existing virtual environment..."
            try {
                Remove-Item -Recurse -Force ".venv"
                Write-Info "Existing virtual environment removed."
            }
            catch {
                Write-ErrorMsg "Failed to remove existing virtual environment: $_"
                exit 1
            }
        } else {
            Write-Warning "Virtual environment already exists in .venv"
            $choice = Read-Host "Do you want to recreate it? (y/N)"
            if ($choice -match '^[Yy]') {
                try {
                    Remove-Item -Recurse -Force ".venv"
                    Write-Info "Existing virtual environment removed."
                }
                catch {
                    Write-ErrorMsg "Failed to remove existing virtual environment: $_"
                    exit 1
                }
            } else {
                Write-Info "Keeping existing virtual environment."
                return $false
            }
        }
    }
    return $true
}

# Function to create VS Code settings
function New-VSCodeSettings {
    try {
        $vsCodeFolder = ".vscode"
        if (-not (Test-Path $vsCodeFolder)) {
            New-Item -ItemType Directory -Path $vsCodeFolder -Force | Out-Null
        }

        $settingsPath = Join-Path $vsCodeFolder "settings.json"
        $interpreterPath = '${workspaceFolder}/.venv/Scripts/python.exe'
        
        $settings = @{
            "python.defaultInterpreterPath" = $interpreterPath
            "python.terminal.activateEnvironment" = $true
            "python.linting.enabled" = $true
            "python.linting.pylintEnabled" = $false
            "python.linting.flake8Enabled" = $true
            "python.formatting.provider" = "black"
        }

        if (Test-Path $settingsPath) {
            try {
                $existingSettings = Get-Content -Path $settingsPath -Raw | ConvertFrom-Json
                $settingsHashtable = @{}
                $existingSettings.PSObject.Properties | ForEach-Object {
                    $settingsHashtable[$_.Name] = $_.Value
                }
                foreach ($key in $settings.Keys) {
                    $settingsHashtable[$key] = $settings[$key]
                }
                $settings = $settingsHashtable
            }
            catch {
                Write-Warning "Could not parse existing settings.json. Creating new file."
            }
        }

        $settings | ConvertTo-Json -Depth 10 | Out-File -FilePath $settingsPath -Encoding UTF8 -Force
        Write-Info "VS Code settings configured."
    }
    catch {
        Write-Warning "Failed to create VS Code settings: $_"
    }
}

# Main execution
try {
    Write-Host "🐍 Python Virtual Environment Generator" -ForegroundColor Cyan
    Write-Host "=====================================" -ForegroundColor Cyan
    
    $projectPath = Get-Location
    Write-Info "Project directory: $projectPath"
    
    # Check if we should create virtual environment
    if (-not (Test-VirtualEnvironment)) {
        Write-Info "Using existing virtual environment."
        New-VSCodeSettings
        Write-Info "Setup completed successfully!"
        exit 0
    }
    
    # Get Python executable
    $pythonExe = Get-PythonExecutable
    
    # Create virtual environment
    Write-Info "Creating virtual environment..."
    try {
        & $pythonExe -m venv .venv
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to create virtual environment"
        }
    } catch {
        Write-ErrorMsg "Failed to create virtual environment: $_"
        exit 1
    }
    
    # Create VS Code settings
    New-VSCodeSettings
    
    # Clean up pyenv local file if using system Python
    if ((Test-Path ".python-version") -and (-not (Get-Command pyenv -ErrorAction SilentlyContinue))) {
        Remove-Item ".python-version" -Force
        Write-Info "Removed unnecessary .python-version file."
    }
    
    Write-Info "Virtual environment created successfully!"
    Write-Host "`n🎉 Setup completed! To activate the environment, run:" -ForegroundColor Green
    Write-Host "   .\.venv\Scripts\Activate.ps1" -ForegroundColor Yellow
    
} catch {
    Write-ErrorMsg "Setup failed: $($_.Exception.Message)"
    exit 1
}
