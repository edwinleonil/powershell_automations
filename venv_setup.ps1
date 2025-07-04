
#Requires -Version 5.1
<#
.SYNOPSIS
    Creates a Python virtual environment with VS Code integration.

.DESCRIPTION
    This script creates a Python virtual environment using either pyenv (if available) 
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
                    throw "Both 'py -0' and 'py --list' failed"
                }
            }
            
            $installedVersions = @()
            $versionDetails = @{}
            
            foreach ($line in $pyOutput) {
                # Parse py -0 output (format: " -3.11-64          C:\Python311\python.exe *")
                # Parse py --list output (format: " -3.11-64 *" or " -3.10-64")
                if ($line -match '^\s*-([0-9]+\.[0-9]+)(?:-(\d+))?\s*(.*)') {
                    $version = $matches[1]
                    $architecture = if ($matches[2]) { $matches[2] } else { "64" }
                    $pathAndDefault = $matches[3].Trim()
                    $isDefault = $pathAndDefault -match '\*'
                    
                    # Extract path if available (from py -0)
                    $pythonPath = ""
                    if ($pathAndDefault -match '([A-Z]:\\[^*]+\.exe)') {
                        $pythonPath = $matches[1].Trim()
                    }
                    
                    $versionKey = "$version-$architecture"
                    if ($installedVersions -notcontains $versionKey) {
                        $installedVersions += $versionKey
                        $versionDetails[$versionKey] = @{
                            Version = $version
                            Architecture = $architecture
                            Path = $pythonPath
                            IsDefault = $isDefault
                            Command = "py -$version"
                        }
                    }
                }
            }
            
            # Sort versions (newest first, then by architecture)
            $installedVersions = $installedVersions | Sort-Object {
                $parts = $_ -split '-'
                $versionParts = $parts[0] -split '\.'
                $majorMinor = [int]$versionParts[0] * 100 + [int]$versionParts[1]
                $arch = [int]$parts[1]
                -$majorMinor * 1000 - $arch  # Negative for descending sort
            }
            
            if ($installedVersions.Count -eq 0) {
                Write-Warning "No Python versions found via py launcher. Trying pyenv..."
            } else {
                Write-Host "`nAvailable Python versions (via py launcher):"
                for ($i = 0; $i -lt $installedVersions.Count; $i++) {
                    $versionKey = $installedVersions[$i]
                    $details = $versionDetails[$versionKey]
                    $defaultMarker = if ($details.IsDefault) { " (default)" } else { "" }
                    $pathInfo = if ($details.Path) { " - $($details.Path)" } else { "" }
                    Write-Host "[$($i + 1)] Python $($details.Version) ($($details.Architecture)-bit)$defaultMarker$pathInfo"
                }
                
                do {
                    $selection = Read-Host "`nSelect Python version (1-$($installedVersions.Count)) or press Enter for default"
                    
                    if ([string]::IsNullOrWhiteSpace($selection)) {
                        # Find default version or use first
                        $defaultVersion = $installedVersions | Where-Object { $versionDetails[$_].IsDefault } | Select-Object -First 1
                        $selectedVersionKey = if ($defaultVersion) { $defaultVersion } else { $installedVersions[0] }
                        break
                    }
                    
                    if ($selection -match '^\d+$') {
                        $selectionNum = [int]$selection
                        if (($selectionNum -ge 1) -and ($selectionNum -le $installedVersions.Count)) {
                            $selectedVersionKey = $installedVersions[$selection - 1]
                            break
                        }
                    }
                    
                    Write-Warning "Invalid selection. Please enter a number between 1 and $($installedVersions.Count)."
                } while ($true)
                
                $selectedDetails = $versionDetails[$selectedVersionKey]
                Write-Info "Using Python $($selectedDetails.Version) ($($selectedDetails.Architecture)-bit)"
                return $selectedDetails.Command
            }
        }
        catch {
            Write-Warning "Error with py launcher: $_. Trying pyenv..."
        }
    }

    # Try pyenv second
    if (Get-Command pyenv -ErrorAction SilentlyContinue) {
        Write-Info "Found pyenv. Getting available Python versions..."
        
        try {
            $allVersions = pyenv versions --bare
            $installedVersions = @()
            foreach ($version in $allVersions) {
                if (($version -notlike "*system*") -and ($version -notlike "*envs/*")) {
                    $installedVersions += $version
                }
            }
            
            if ($installedVersions.Count -eq 0) {
                Write-Warning "No Python versions found in pyenv. Falling back to system Python."
            } else {
                Write-Host "`nAvailable Python versions (via pyenv):"
                for ($i = 0; $i -lt $installedVersions.Count; $i++) {
                    Write-Host "[$($i + 1)] $($installedVersions[$i])"
                }
                
                do {
                    $selection = Read-Host "`nSelect Python version (1-$($installedVersions.Count)) or press Enter for latest"
                    
                    if ([string]::IsNullOrWhiteSpace($selection)) {
                        $selectedVersion = $installedVersions[0]  # Use first (usually latest)
                        break
                    }
                    
                    if ($selection -match '^\d+$') {
                        $selectionNum = [int]$selection
                        if (($selectionNum -ge 1) -and ($selectionNum -le $installedVersions.Count)) {
                            $selectedVersion = $installedVersions[$selection - 1]
                            break
                        }
                    }
                    
                    Write-Warning "Invalid selection. Please enter a number between 1 and $($installedVersions.Count)."
                } while ($true)
                
                Write-Info "Setting local Python version to $selectedVersion"
                try {
                    & pyenv local $selectedVersion
                    if ($LASTEXITCODE -ne 0) {
                        throw "Failed to set Python version"
                    }
                } catch {
                    Write-ErrorMsg "Failed to set Python version: $_"
                    exit 1
                }
                return "python"
            }
        }
        catch {
            Write-Warning "Error with pyenv: $_. Falling back to system Python."
        }
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
