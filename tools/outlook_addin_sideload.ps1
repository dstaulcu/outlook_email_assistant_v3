# Outlook Add-in Registry Installation Script
# This script installs Office add-ins via Windows registry keys for environments
# where Outlook Web Access is not available

param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestUrl,
    
    [string]$Environment = "Prd",
    [switch]$UseLocalManifest = $false,
    [switch]$SkipCacheClear = $false,
    [switch]$Uninstall = $false,
    [switch]$CurrentUserOnly = $false,
    [switch]$Help = $false
)

function Write-Status {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

function Show-Help {
    Write-Status "Outlook Add-in Registry Installation Script" "Blue"
    Write-Status "===========================================" "Blue"
    Write-Status ""
    Write-Status "Installs Office add-ins via Windows registry keys" "White"
    Write-Status ""
    Write-Status "Usage:" "Yellow"
    Write-Status "  .\outlook_addin_sideload.ps1 -ManifestUrl <url>" "White"
    Write-Status "  .\outlook_addin_sideload.ps1 -ManifestUrl <url> -Environment Dev" "White"
    Write-Status "  .\outlook_addin_sideload.ps1 -ManifestUrl <path> -UseLocalManifest" "White"
    Write-Status "  .\outlook_addin_sideload.ps1 -ManifestUrl <url> -Uninstall" "White"
    Write-Status ""
    Write-Status "Parameters:" "Yellow"
    Write-Status "  -ManifestUrl       URL or path to the manifest.xml file (required)" "White"
    Write-Status "  -Environment       Environment name for logging (Dev/Prd, default: Prd)" "White"
    Write-Status "  -UseLocalManifest  Use local file path instead of URL" "White"
    Write-Status "  -SkipCacheClear    Skip clearing Office add-in cache" "White"
    Write-Status "  -Uninstall         Remove the add-in from registry" "White"
    Write-Status "  -CurrentUserOnly   Install only for current user (HKCU)" "White"
    Write-Status "  -Help              Show this help message" "White"
    Write-Status ""
    Write-Status "Examples:" "Yellow"
    Write-Status "  # Install production manifest from S3:" "Cyan"
    Write-Status "  .\outlook_addin_sideload.ps1 -ManifestUrl 'https://your-bucket.s3.region.amazonaws.com/manifest.xml'" "White"
    Write-Status ""
    Write-Status "  # Install for current user only:" "Cyan"
    Write-Status "  .\outlook_addin_sideload.ps1 -ManifestUrl 'https://your-bucket.s3.region.amazonaws.com/manifest.xml' -CurrentUserOnly" "White"
    Write-Status ""
    Write-Status "  # Uninstall add-in:" "Cyan"
    Write-Status "  .\outlook_addin_sideload.ps1 -ManifestUrl 'https://your-bucket.s3.region.amazonaws.com/manifest.xml' -Uninstall" "White"
    Write-Status ""
}

function Get-OutlookVersions {
    Write-Status "🔍 Detecting installed Office versions..." "Cyan"
    
    $officeVersions = @()
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office"
    )
    
    foreach ($basePath in $registryPaths) {
        if (Test-Path $basePath) {
            $versions = Get-ChildItem $basePath -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '\d+\.\d+$' }
            foreach ($version in $versions) {
                $versionNumber = $version.PSChildName
                $outlookPath = Join-Path $version.PSPath "Outlook"
                if (Test-Path $outlookPath) {
                    $officeVersions += $versionNumber
                    Write-Status "   Found Office version: $versionNumber" "Gray"
                }
            }
        }
    }
    
    if ($officeVersions.Count -eq 0) {
        Write-Status "⚠️  No Office installations found" "Yellow"
        # Default to common versions
        $officeVersions = @("16.0", "15.0")
        Write-Status "   Using default versions: $($officeVersions -join ', ')" "Gray"
    }
    
    return $officeVersions | Sort-Object -Unique -Descending
}

function Get-AddinInfoFromManifest {
    param([string]$ManifestPath, [bool]$IsLocal)
    
    Write-Status "📋 Extracting add-in information from manifest..." "Cyan"
    
    try {
        if ($IsLocal) {
            if (-not (Test-Path $ManifestPath)) {
                throw "Local manifest file not found: $ManifestPath"
            }
            [xml]$manifest = Get-Content $ManifestPath
        } else {
            $manifestContent = Invoke-WebRequest -Uri $ManifestPath -TimeoutSec 10 -ErrorAction Stop
            [xml]$manifest = $manifestContent.Content
        }
        
        $addinId = $manifest.OfficeApp.Id
        $displayName = $manifest.OfficeApp.DisplayName.DefaultValue
        $description = $manifest.OfficeApp.Description.DefaultValue
        
        if (-not $addinId) {
            throw "Manifest does not contain a valid add-in ID"
        }
        
        Write-Status "✅ Manifest parsed successfully" "Green"
        Write-Status "   Add-in ID: $addinId" "Gray"
        Write-Status "   Display Name: $displayName" "Gray"
        Write-Status "   Description: $description" "Gray"
        
        return @{
            Id = $addinId
            DisplayName = $displayName
            Description = $description
            ManifestUrl = $ManifestPath
        }
        
    } catch {
        Write-Status "❌ Failed to parse manifest: $($_.Exception.Message)" "Red"
        throw
    }
}

function Test-RegistryWritePermission {
    param([string]$RegistryPath)
    
    try {
        # Try to create a temporary test key
        $testKeyPath = Join-Path $RegistryPath "TempTestKey"
        New-Item -Path $testKeyPath -Force -ErrorAction Stop | Out-Null
        Remove-Item -Path $testKeyPath -Force -ErrorAction SilentlyContinue
        return $true
    } catch {
        return $false
    }
}

function Install-AddinToRegistry {
    param(
        [hashtable]$AddinInfo,
        [string[]]$OfficeVersions,
        [bool]$CurrentUserOnly
    )
    
    Write-Status "📝 Installing add-in to Windows registry..." "Cyan"
    
    $registryRoots = @()
    if ($CurrentUserOnly) {
        $registryRoots = @("HKCU:")
        Write-Status "   Installing for current user only" "Gray"
    } else {
        $registryRoots = @("HKLM:", "HKCU:")
        Write-Status "   Installing for all users and current user" "Gray"
    }
    
    $installCount = 0
    $failCount = 0
    
    foreach ($root in $registryRoots) {
        foreach ($version in $OfficeVersions) {
            $registryPath = "$root\Software\Microsoft\Office\$version\Outlook\Addins\$($AddinInfo.Id)"
            
            try {
                # Check if we have write permission
                $parentPath = "$root\Software\Microsoft\Office\$version\Outlook\Addins"
                if (-not (Test-Path $parentPath)) {
                    New-Item -Path $parentPath -Force | Out-Null
                }
                
                if (-not (Test-RegistryWritePermission -RegistryPath $parentPath)) {
                    if ($root -eq "HKLM:") {
                        Write-Status "⚠️  No permission to write to $root for Office $version (run as Administrator)" "Yellow"
                    } else {
                        Write-Status "⚠️  No permission to write to $root for Office $version" "Yellow"
                    }
                    $failCount++
                    continue
                }
                
                # Create the add-in registry key
                New-Item -Path $registryPath -Force | Out-Null
                
                # Set registry values
                Set-ItemProperty -Path $registryPath -Name "Description" -Value $AddinInfo.Description -Type String
                Set-ItemProperty -Path $registryPath -Name "FriendlyName" -Value $AddinInfo.DisplayName -Type String
                Set-ItemProperty -Path $registryPath -Name "Manifest" -Value $AddinInfo.ManifestUrl -Type String
                Set-ItemProperty -Path $registryPath -Name "LoadBehavior" -Value 3 -Type DWord
                
                Write-Status "✅ Installed to: $registryPath" "Green"
                $installCount++
                
            } catch {
                Write-Status "❌ Failed to install to $registryPath`: $($_.Exception.Message)" "Red"
                $failCount++
            }
        }
    }
    
    Write-Status ""
    Write-Status "📊 Installation Summary:" "Yellow"
    Write-Status "   Successful installs: $installCount" "Green"
    Write-Status "   Failed installs: $failCount" "Red"
    
    if ($installCount -eq 0) {
        throw "No successful installations completed"
    }
    
    if ($failCount -gt 0 -and -not $CurrentUserOnly) {
        Write-Status ""
        Write-Status "💡 Tip: If HKLM installations failed, try running as Administrator" "Yellow"
        Write-Status "   Or use -CurrentUserOnly to install only for the current user" "Yellow"
    }
}

function Uninstall-AddinFromRegistry {
    param(
        [hashtable]$AddinInfo,
        [string[]]$OfficeVersions
    )
    
    Write-Status "🗑️  Uninstalling add-in from Windows registry..." "Cyan"
    
    $registryRoots = @("HKLM:", "HKCU:")
    $uninstallCount = 0
    
    foreach ($root in $registryRoots) {
        foreach ($version in $OfficeVersions) {
            $registryPath = "$root\Software\Microsoft\Office\$version\Outlook\Addins\$($AddinInfo.Id)"
            
            if (Test-Path $registryPath) {
                try {
                    Remove-Item -Path $registryPath -Recurse -Force -ErrorAction Stop
                    Write-Status "✅ Uninstalled from: $registryPath" "Green"
                    $uninstallCount++
                } catch {
                    Write-Status "❌ Failed to uninstall from $registryPath`: $($_.Exception.Message)" "Red"
                }
            }
        }
    }
    
    if ($uninstallCount -eq 0) {
        Write-Status "ℹ️  Add-in was not found in registry" "Gray"
    } else {
        Write-Status "✅ Uninstalled from $uninstallCount location(s)" "Green"
    }
}

function Test-ManifestAccessibility {
    param([string]$ManifestPath, [bool]$IsLocal)
    
    Write-Status "🔍 Validating manifest accessibility..." "Cyan"
    
    if ($IsLocal) {
        if (-not (Test-Path $ManifestPath)) {
            Write-Status "❌ Local manifest file not found: $ManifestPath" "Red"
            return $false
        }
        
        try {
            [xml]$manifest = Get-Content $ManifestPath
            $displayName = $manifest.OfficeApp.DisplayName.DefaultValue
            $id = $manifest.OfficeApp.Id
            Write-Status "✅ Local manifest is valid XML" "Green"
            Write-Status "   Display Name: $displayName" "Gray"
            Write-Status "   ID: $id" "Gray"
            return $true
        } catch {
            Write-Status "❌ Local manifest has invalid XML: $($_.Exception.Message)" "Red"
            return $false
        }
    } else {
        try {
            $response = Invoke-WebRequest -Uri $ManifestPath -Method Head -TimeoutSec 10 -ErrorAction Stop
            Write-Status "✅ Manifest URL is accessible [Status: $($response.StatusCode)]" "Green"
            
            # Download and validate XML
            $manifestContent = Invoke-WebRequest -Uri $ManifestPath -TimeoutSec 10 -ErrorAction Stop
            [xml]$manifest = $manifestContent.Content
            $displayName = $manifest.OfficeApp.DisplayName.DefaultValue
            $id = $manifest.OfficeApp.Id
            Write-Status "✅ Manifest has valid XML" "Green"
            Write-Status "   Display Name: $displayName" "Gray"
            Write-Status "   ID: $id" "Gray"
            return $true
        } catch {
            Write-Status "❌ Manifest not accessible: $($_.Exception.Message)" "Red"
            Write-Status "💡 Check URL, permissions, and network connectivity" "Yellow"
            return $false
        }
    }
}

function Stop-OutlookProcess {
    Write-Status "📧 Checking Outlook process..." "Cyan"
    
    $outlookProcess = Get-Process "OUTLOOK" -ErrorAction SilentlyContinue
    if ($outlookProcess) {
        Write-Status "⚠️  Outlook is currently running (PID: $($outlookProcess.Id))" "Yellow"
        Write-Status "   Registry changes require Outlook restart to take effect" "White"
        Write-Status ""
        
        $response = Read-Host "Close Outlook automatically? (y/n)"
        if ($response -eq 'y' -or $response -eq 'Y') {
            try {
                Stop-Process -Name "OUTLOOK" -Force -ErrorAction Stop
                Start-Sleep -Seconds 3
                Write-Status "✅ Outlook closed successfully" "Green"
            } catch {
                Write-Status "❌ Failed to close Outlook: $($_.Exception.Message)" "Red"
                Write-Status "Please close Outlook manually and restart it after installation" "Yellow"
                return $false
            }
        } else {
            Write-Status "⚠️  Please restart Outlook after installation to load the add-in" "Yellow"
        }
    } else {
        Write-Status "✅ Outlook is not running" "Green"
    }
    return $true
}

function Clear-OfficeAddInCache {
    if ($SkipCacheClear) {
        Write-Status "⏭️  Skipping cache clear (as requested)" "Yellow"
        return
    }
    
    Write-Status "🧹 Clearing Office add-in cache..." "Cyan"
    
    $cachePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
        "$env:LOCALAPPDATA\Microsoft\Office\Wef"
    )
    
    $clearedPaths = 0
    foreach ($cachePath in $cachePaths) {
        if (Test-Path $cachePath) {
            try {
                Remove-Item -Path $cachePath -Recurse -Force -ErrorAction Stop
                Write-Status "✅ Cleared cache: $cachePath" "Green"
                $clearedPaths++
            } catch {
                Write-Status "⚠️  Could not clear cache: $cachePath ($($_.Exception.Message))" "Yellow"
            }
        }
    }
    
    if ($clearedPaths -eq 0) {
        Write-Status "ℹ️  No cache directories found to clear" "Gray"
    }
}

function Show-CompletionInstructions {
    param([bool]$IsUninstall, [string]$AddinName)
    
    Write-Status ""
    if ($IsUninstall) {
        Write-Status "✅ Add-in uninstallation complete!" "Green"
        Write-Status ""
        Write-Status "📝 Next Steps:" "Yellow"
        Write-Status "🔹 Restart Outlook to remove the add-in from the ribbon" "White"
        Write-Status "� The add-in should no longer appear in Outlook's ribbon" "White"
    } else {
        Write-Status "✅ Add-in installation complete!" "Green"
        Write-Status ""
        Write-Status "� Next Steps:" "Yellow"
        Write-Status "🔹 Start or restart Outlook" "White"
        Write-Status "� Look for '$AddinName' in the ribbon" "White"
        Write-Status "🔹 The add-in should appear in the Home or Message tab" "White"
        Write-Status ""
        Write-Status "🔧 Troubleshooting:" "Yellow"
        Write-Status "• If add-in doesn't appear, check File → Options → Add-ins" "Gray"
        Write-Status "• Ensure 'Optional Connected Experiences' is enabled" "Gray"
        Write-Status "• Run: .\outlook_addin_diagnostics.ps1 for detailed diagnostics" "Gray"
    }
    Write-Status ""
}

# Main execution
if ($Help) {
    Show-Help
    exit 0
}

# Validate parameters
if ([string]::IsNullOrWhiteSpace($ManifestUrl)) {
    Write-Status "❌ ManifestUrl parameter is required" "Red"
    Write-Status "Use -Help for usage information" "Yellow"
    exit 1
}

# Check for Administrator privileges when installing to HKLM
if (-not $CurrentUserOnly -and -not $Uninstall) {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin) {
        Write-Status "⚠️  Running without Administrator privileges" "Yellow"
        Write-Status "   HKLM installations may fail. Consider running as Administrator" "Yellow"
        Write-Status "   or use -CurrentUserOnly for current user installation only" "Yellow"
        Write-Status ""
    }
}

# Main script execution
$actionText = if ($Uninstall) { "Uninstallation" } else { "Installation" }
Write-Status "🚀 Outlook Add-in Registry $actionText" "Blue"
Write-Status "======================================" "Blue"
Write-Status ""
Write-Status "Environment: $Environment" "Gray"
Write-Status "Manifest Source: $ManifestUrl" "Gray"
Write-Status "Use Local File: $UseLocalManifest" "Gray"
Write-Status "Action: $actionText" "Gray"
Write-Status ""

try {
    # Step 1: Validate manifest and extract add-in info
    if (-not (Test-ManifestAccessibility -ManifestPath $ManifestUrl -IsLocal $UseLocalManifest)) {
        Write-Status ""
        Write-Status "❌ Cannot proceed - manifest validation failed" "Red"
        exit 1
    }
    
    $addinInfo = Get-AddinInfoFromManifest -ManifestPath $ManifestUrl -IsLocal $UseLocalManifest
    
    # Step 2: Detect Office versions
    $officeVersions = Get-OutlookVersions
    
    # Step 3: Handle Outlook process
    if (-not (Stop-OutlookProcess)) {
        Write-Status "⚠️  Continuing with Outlook running - restart required after installation" "Yellow"
    }
    
    # Step 4: Clear cache (if not uninstalling)
    if (-not $Uninstall) {
        Clear-OfficeAddInCache
    }
    
    # Step 5: Install or uninstall from registry
    if ($Uninstall) {
        Uninstall-AddinFromRegistry -AddinInfo $addinInfo -OfficeVersions $officeVersions
    } else {
        Install-AddinToRegistry -AddinInfo $addinInfo -OfficeVersions $officeVersions -CurrentUserOnly $CurrentUserOnly
    }
    
    # Step 6: Show completion instructions
    Show-CompletionInstructions -IsUninstall $Uninstall -AddinName $addinInfo.DisplayName
    
} catch {
    Write-Status ""
    Write-Status "❌ $actionText failed: $($_.Exception.Message)" "Red"
    Write-Status ""
    Write-Status "📞 Need help? Run: .\outlook_addin_diagnostics.ps1" "Yellow"
    exit 1
}
