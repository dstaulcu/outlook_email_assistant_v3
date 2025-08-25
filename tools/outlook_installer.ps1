# Outlook Email Assistant - Windows Installer Script
# This script downloads the manifest from S3, stops Outlook, clears cache, and configures registry for sideloading

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Dev", "Prd")]
    [string]$Environment = "Prd",
    
    [Parameter(Mandatory=$false)]
    [string]$ManifestUrl = "",
    
    [Parameter(Mandatory=$false)]
    [string]$InstallPath = "$env:LOCALAPPDATA\OutlookEmailAssistant",
    
    [Parameter(Mandatory=$false)]
    [switch]$Silent = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$UninstallOnly = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$Help = $false
)

# Function to write colored output
function Write-Status {
    param([string]$Message, [string]$Color = "White")
    if (-not $Silent) {
        Write-Host $Message -ForegroundColor $Color
    }
}

# Function to show help
function Show-Help {
    Write-Status "Outlook Email Assistant - Windows Installer" "Blue"
    Write-Status "=============================================" "Blue"
    Write-Status ""
    Write-Status "This script installs the Outlook Email Assistant add-in by:" "White"
    Write-Status "1. Downloading the manifest from S3" "White"
    Write-Status "2. Stopping Outlook processes" "White"
    Write-Status "3. Clearing Outlook add-in cache" "White"
    Write-Status "4. Adding registry keys for sideloading" "White"
    Write-Status ""
    Write-Status "Usage:" "Yellow"
    Write-Status "  .\outlook_installer.ps1 -Environment Prd" "White"
    Write-Status "  .\outlook_installer.ps1 -ManifestUrl 'https://custom-url.com/manifest.xml'" "White"
    Write-Status "  .\outlook_installer.ps1 -UninstallOnly" "White"
    Write-Status ""
    Write-Status "Parameters:" "Yellow"
    Write-Status "  -Environment    Environment to install from (Dev, Prd) [Default: Prd]" "White"
    Write-Status "  -ManifestUrl    Custom manifest URL (overrides environment)" "White"
    Write-Status "  -InstallPath    Installation directory [Default: %LOCALAPPDATA%\OutlookEmailAssistant]" "White"
    Write-Status "  -Silent         Run silently without user prompts" "White"
    Write-Status "  -UninstallOnly  Only remove the add-in, don't install" "White"
    Write-Status "  -Help           Show this help message" "White"
    Write-Status ""
}

# Function to check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to get manifest URL from environment configuration
function Get-ManifestUrl {
    param([string]$Env)
    
    $configPath = Join-Path (Split-Path $PSScriptRoot) "tools\deployment-environments.json"
    if (-not (Test-Path $configPath)) {
        Write-Status "Warning: deployment-environments.json not found, using default URLs" "Yellow"
        switch ($Env) {
            "Dev" { return "https://293354421824-outlook-email-assistant-dev.s3.us-east-1.amazonaws.com/manifest.xml" }
            "Prd" { return "https://293354421824-outlook-email-assistant-prd.s3.us-east-1.amazonaws.com/manifest.xml" }
        }
    }
    
    try {
        $config = Get-Content $configPath | ConvertFrom-Json
        $envConfig = $config.environments.$Env
        if ($envConfig -and $envConfig.publicUri) {
            return "$($envConfig.publicUri.protocol)://$($envConfig.publicUri.host)/manifest.xml"
        }
    } catch {
        Write-Status "Warning: Could not parse deployment configuration: $($_.Exception.Message)" "Yellow"
    }
    
    # Fallback URLs
    switch ($Env) {
        "Dev" { return "https://293354421824-outlook-email-assistant-dev.s3.us-east-1.amazonaws.com/manifest.xml" }
        "Prd" { return "https://293354421824-outlook-email-assistant-prd.s3.us-east-1.amazonaws.com/manifest.xml" }
    }
}

# Function to download manifest file
function Download-Manifest {
    param([string]$Url, [string]$Destination)
    
    Write-Status "Downloading manifest from: $Url" "Blue"
    
    try {
        # Create destination directory if it doesn't exist
        $destinationDir = Split-Path $Destination -Parent
        if (-not (Test-Path $destinationDir)) {
            New-Item -Path $destinationDir -ItemType Directory -Force | Out-Null
        }
        
        # Download with error handling
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($Url, $Destination)
        
        Write-Status "Manifest downloaded successfully to: $Destination" "Green"
        return $true
    } catch {
        Write-Status "Failed to download manifest: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to validate manifest file
function Test-ManifestFile {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        return $false
    }
    
    try {
        [xml]$manifest = Get-Content $Path
        $officeApp = $manifest.SelectSingleNode("//OfficeApp")
        if ($officeApp) {
            Write-Status "Manifest validation passed" "Green"
            return $true
        } else {
            Write-Status "Invalid manifest format" "Red"
            return $false
        }
    } catch {
        Write-Status "Manifest validation failed: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to stop Outlook processes
function Stop-OutlookProcesses {
    Write-Status "Stopping Outlook processes..." "Blue"
    
    $outlookProcesses = Get-Process OUTLOOK -ErrorAction SilentlyContinue
    if ($outlookProcesses) {
        $outlookProcesses | Stop-Process -Force
        Write-Status "Outlook processes stopped" "Green"
        Start-Sleep -Seconds 3
    } else {
        Write-Status "No Outlook processes found" "Gray"
    }
}

# Function to clear Outlook cache
function Clear-OutlookCache {
    Write-Status "Clearing Outlook add-in cache..." "Blue"
    
    $cachePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
        "$env:LOCALAPPDATA\Microsoft\Office\15.0\Wef",
        "$env:LOCALAPPDATA\Microsoft\Office\14.0\Wef"
    )
    
    foreach ($cachePath in $cachePaths) {
        if (Test-Path $cachePath) {
            $maxRetries = 5
            $retryDelay = 2
            $attempt = 0
            $success = $false
            
            while (-not $success -and $attempt -lt $maxRetries) {
                try {
                    Remove-Item "$cachePath\*" -Recurse -Force -ErrorAction Stop
                    $success = $true
                    Write-Status "Cleared cache: $cachePath" "Green"
                } catch {
                    Write-Status "Attempt $($attempt + 1) to clear cache failed: $($_.Exception.Message)" "Yellow"
                    Start-Sleep -Seconds $retryDelay
                    $attempt++
                }
            }
            
            if (-not $success) {
                Write-Status "Warning: Some cache files could not be deleted: $cachePath" "Yellow"
            }
        }
    }
}

# Function to add registry keys for sideloading
function Add-SideloadRegistryKeys {
    param([string]$ManifestPath, [string]$AddInName = "OutlookEmailAssistant")
    
    Write-Status "Adding registry keys for sideloading..." "Blue"
    
    # Registry paths for different Office versions
    $registryPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer", 
        "HKCU:\SOFTWARE\Microsoft\Office\14.0\WEF\Developer"
    )
    
    $success = $false
    
    foreach ($regPath in $registryPaths) {
        try {
            # Create the Developer key if it doesn't exist
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
                Write-Status "Created registry path: $regPath" "Green"
            }
            
            # Add the manifest path
            Set-ItemProperty -Path $regPath -Name $AddInName -Value $ManifestPath -Type String
            Write-Status "Added registry entry: $regPath\$AddInName = $ManifestPath" "Green"
            $success = $true
        } catch {
            Write-Status "Warning: Could not add registry key at $regPath`: $($_.Exception.Message)" "Yellow"
        }
    }
    
    if ($success) {
        Write-Status "Registry keys added successfully" "Green"
        return $true
    } else {
        Write-Status "Failed to add registry keys" "Red"
        return $false
    }
}

# Function to remove registry keys (for uninstall)
function Remove-SideloadRegistryKeys {
    param([string]$AddInName = "OutlookEmailAssistant")
    
    Write-Status "Removing registry keys..." "Blue"
    
    $registryPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer",
        "HKCU:\SOFTWARE\Microsoft\Office\14.0\WEF\Developer"
    )
    
    foreach ($regPath in $registryPaths) {
        try {
            if (Test-Path $regPath) {
                $property = Get-ItemProperty -Path $regPath -Name $AddInName -ErrorAction SilentlyContinue
                if ($property) {
                    Remove-ItemProperty -Path $regPath -Name $AddInName -ErrorAction Stop
                    Write-Status "Removed registry entry: $regPath\$AddInName" "Green"
                }
            }
        } catch {
            Write-Status "Warning: Could not remove registry key: $($_.Exception.Message)" "Yellow"
        }
    }
}

# Function to cleanup installation files
function Remove-InstallationFiles {
    param([string]$InstallPath)
    
    Write-Status "Cleaning up installation files..." "Blue"
    
    try {
        if (Test-Path $InstallPath) {
            Remove-Item $InstallPath -Recurse -Force
            Write-Status "Installation directory removed: $InstallPath" "Green"
        }
    } catch {
        Write-Status "Warning: Could not remove installation directory: $($_.Exception.Message)" "Yellow"
    }
}

# Function to verify installation
function Test-Installation {
    param([string]$ManifestPath)
    
    Write-Status "Verifying installation..." "Blue"
    
    # Check if manifest file exists
    if (-not (Test-Path $ManifestPath)) {
        Write-Status "Installation verification failed: Manifest file not found" "Red"
        return $false
    }
    
    # Check if registry keys exist
    $registryPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer"
    )
    
    $registryFound = $false
    foreach ($regPath in $registryPaths) {
        try {
            if (Test-Path $regPath) {
                $property = Get-ItemProperty -Path $regPath -Name "OutlookEmailAssistant" -ErrorAction SilentlyContinue
                if ($property) {
                    $registryFound = $true
                    break
                }
            }
        } catch {
            # Continue checking other paths
        }
    }
    
    if ($registryFound) {
        Write-Status "Installation verification passed" "Green"
        return $true
    } else {
        Write-Status "Installation verification failed: Registry keys not found" "Red"
        return $false
    }
}

# Main installation function
function Install-OutlookEmailAssistant {
    param([string]$ManifestUrl, [string]$InstallPath)
    
    Write-Status "Starting Outlook Email Assistant installation..." "Blue"
    Write-Status "Environment: $Environment" "Gray"
    Write-Status "Manifest URL: $ManifestUrl" "Gray"
    Write-Status "Install Path: $InstallPath" "Gray"
    Write-Status ""
    
    # Step 1: Download manifest
    $manifestPath = Join-Path $InstallPath "manifest.xml"
    if (-not (Download-Manifest -Url $ManifestUrl -Destination $manifestPath)) {
        throw "Failed to download manifest file"
    }
    
    # Step 2: Validate manifest
    if (-not (Test-ManifestFile -Path $manifestPath)) {
        throw "Invalid manifest file"
    }
    
    # Step 3: Stop Outlook processes
    Stop-OutlookProcesses
    
    # Step 4: Clear cache
    Clear-OutlookCache
    
    # Step 5: Add registry keys
    if (-not (Add-SideloadRegistryKeys -ManifestPath $manifestPath)) {
        Write-Status "Warning: Registry keys may not have been added properly" "Yellow"
    }
    
    # Step 6: Verify installation
    if (Test-Installation -ManifestPath $manifestPath) {
        Write-Status "Installation completed successfully!" "Green"
        Write-Status ""
        Write-Status "The Outlook Email Assistant add-in has been installed." "White"
        Write-Status "Start Outlook to see the add-in in the ribbon." "White"
        
        if (-not $Silent) {
            $response = Read-Host "Would you like to start Outlook now? (y/n)"
            if ($response -eq 'y' -or $response -eq 'Y' -or $response -eq '') {
                Write-Status "Starting Outlook..." "Blue"
                Start-Process OUTLOOK.EXE
            }
        }
        return $true
    } else {
        throw "Installation verification failed"
    }
}

# Main uninstall function
function Uninstall-OutlookEmailAssistant {
    param([string]$InstallPath)
    
    Write-Status "Starting Outlook Email Assistant uninstallation..." "Blue"
    
    # Step 1: Stop Outlook processes
    Stop-OutlookProcesses
    
    # Step 2: Clear cache
    Clear-OutlookCache
    
    # Step 3: Remove registry keys
    Remove-SideloadRegistryKeys
    
    # Step 4: Remove installation files
    Remove-InstallationFiles -InstallPath $InstallPath
    
    Write-Status "Uninstallation completed successfully!" "Green"
    return $true
}

# Main script execution
try {
    # Show help if requested
    if ($Help) {
        Show-Help
        exit 0
    }
    
    # Check if running with sufficient privileges
    if (-not (Test-Administrator)) {
        Write-Status "Warning: Not running as Administrator. Some operations may fail." "Yellow"
        if (-not $Silent) {
            $response = Read-Host "Continue anyway? (y/n)"
            if ($response -ne 'y' -and $response -ne 'Y' -and $response -ne '') {
                exit 1
            }
        }
    }
    
    # Handle uninstall-only mode
    if ($UninstallOnly) {
        Uninstall-OutlookEmailAssistant -InstallPath $InstallPath
        exit 0
    }
    
    # Determine manifest URL
    if ([string]::IsNullOrEmpty($ManifestUrl)) {
        $ManifestUrl = Get-ManifestUrl -Env $Environment
    }
    
    # Confirm installation if not silent
    if (-not $Silent) {
        Write-Status "Outlook Email Assistant Installer" "Blue"
        Write-Status "=================================" "Blue"
        Write-Status ""
        Write-Status "This will install the Outlook Email Assistant add-in:" "White"
        Write-Status "- Environment: $Environment" "Gray"
        Write-Status "- Manifest URL: $ManifestUrl" "Gray"
        Write-Status "- Install Path: $InstallPath" "Gray"
        Write-Status ""
        Write-Status "The installer will:" "White"
        Write-Status "1. Download the manifest file from S3" "White"
        Write-Status "2. Stop any running Outlook processes" "White"
        Write-Status "3. Clear the Outlook add-in cache" "White"
        Write-Status "4. Add registry keys for sideloading" "White"
        Write-Status ""
        
        $response = Read-Host "Continue with installation? (y/n)"
        if ($response -ne 'y' -and $response -ne 'Y' -and $response -ne '') {
            Write-Status "Installation cancelled by user" "Yellow"
            exit 0
        }
        Write-Status ""
    }
    
    # Perform installation
    Install-OutlookEmailAssistant -ManifestUrl $ManifestUrl -InstallPath $InstallPath
    
} catch {
    Write-Status "Installation failed: $($_.Exception.Message)" "Red"
    Write-Status ""
    Write-Status "Troubleshooting tips:" "Yellow"
    Write-Status "1. Run as Administrator" "White"
    Write-Status "2. Ensure internet connectivity" "White"
    Write-Status "3. Close all Office applications" "White"
    Write-Status "4. Check if Windows Defender or antivirus is blocking the script" "White"
    
    exit 1
}
