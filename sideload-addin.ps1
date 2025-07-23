# PromptEmail Outlook Add-in Sideloader
# This script automates the sideloading of the PromptEmail add-in into Outlook
# Requires Administrator privileges for registry changes

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ManifestUrl = "",
    
    [Parameter(Mandatory=$false)]
    [string]$LocalManifestPath = ".\manifest.xml",
    
    [Parameter(Mandatory=$false)]
    [switch]$RemoveAddin = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$ClearCacheOnly = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$Verbose = $false
)

# Configuration
$AddinId = "12345678-1234-1234-1234-123456789012"  # Must match manifest ID
$AddinName = "PromptEmail"
$DefaultManifestUrl = "https://your-promptemail-bucket-name.s3.amazonaws.com/manifest.xml"

# Colors for output
$Red = "Red"
$Green = "Green"
$Yellow = "Yellow"
$Blue = "Blue"
$Cyan = "Cyan"

function Write-Status {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Stop-OutlookProcesses {
    Write-Status "Stopping Outlook processes..." $Blue
    
    $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($outlookProcesses) {
        Write-Status "Found $($outlookProcesses.Count) Outlook process(es). Stopping..." $Yellow
        
        foreach ($process in $outlookProcesses) {
            try {
                $process.CloseMainWindow()
                Start-Sleep -Seconds 2
                
                if (!$process.HasExited) {
                    $process.Kill()
                }
                Write-Status "✓ Stopped Outlook process (PID: $($process.Id))" $Green
            }
            catch {
                Write-Status "✗ Failed to stop Outlook process: $($_.Exception.Message)" $Red
            }
        }
        
        # Wait for processes to fully terminate
        Start-Sleep -Seconds 3
    }
    else {
        Write-Status "✓ No Outlook processes running" $Green
    }
}

function Clear-OutlookCache {
    Write-Status "Clearing Outlook add-in cache..." $Blue
    
    $cachePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
        "$env:LOCALAPPDATA\Microsoft\Office\Wef",
        "$env:APPDATA\Microsoft\Office\16.0\Wef",
        "$env:TEMP\Outlook WebView",
        "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache"
    )
    
    foreach ($path in $cachePaths) {
        if (Test-Path $path) {
            try {
                Write-Status "Clearing cache: $path" $Cyan
                Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
                Write-Status "✓ Cleared: $path" $Green
            }
            catch {
                Write-Status "⚠ Could not clear: $path - $($_.Exception.Message)" $Yellow
            }
        }
    }
    
    # Clear browser cache for Office
    Write-Status "Clearing Office browser cache..." $Cyan
    $officeCachePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Office\16.0\WebView2",
        "$env:LOCALAPPDATA\Microsoft\Office\WebView2"
    )
    
    foreach ($path in $officeCachePaths) {
        if (Test-Path $path) {
            try {
                Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
                Write-Status "✓ Cleared Office WebView cache: $path" $Green
            }
            catch {
                Write-Status "⚠ Could not clear WebView cache: $path" $Yellow
            }
        }
    }
}

function Get-OutlookRegistryPaths {
    $basePaths = @(
        "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer",
        "HKCU:\Software\Microsoft\Office\WEF\Developer"
    )
    
    $validPaths = @()
    foreach ($path in $basePaths) {
        if (Test-Path $path -ErrorAction SilentlyContinue) {
            $validPaths += $path
        }
        else {
            # Create the registry path if it doesn't exist
            try {
                New-Item -Path $path -Force | Out-Null
                $validPaths += $path
                Write-Status "✓ Created registry path: $path" $Green
            }
            catch {
                Write-Status "⚠ Could not create registry path: $path" $Yellow
            }
        }
    }
    
    return $validPaths
}

function Remove-AddinFromRegistry {
    Write-Status "Removing PromptEmail add-in from registry..." $Blue
    
    $registryPaths = Get-OutlookRegistryPaths
    $removed = $false
    
    foreach ($basePath in $registryPaths) {
        try {
            $addinPath = "$basePath\$AddinId"
            if (Test-Path $addinPath) {
                Remove-Item -Path $addinPath -Recurse -Force
                Write-Status "✓ Removed add-in from: $addinPath" $Green
                $removed = $true
            }
        }
        catch {
            Write-Status "✗ Failed to remove from registry: $($_.Exception.Message)" $Red
        }
    }
    
    if (!$removed) {
        Write-Status "ℹ No existing add-in registration found" $Cyan
    }
}

function Add-AddinToRegistry {
    param([string]$ManifestLocation)
    
    Write-Status "Adding PromptEmail add-in to registry..." $Blue
    
    $registryPaths = Get-OutlookRegistryPaths
    $added = $false
    
    foreach ($basePath in $registryPaths) {
        try {
            $addinPath = "$basePath\$AddinId"
            
            # Create the add-in registry key
            New-Item -Path $addinPath -Force | Out-Null
            
            # Set the manifest location
            Set-ItemProperty -Path $addinPath -Name "Location" -Value $ManifestLocation
            Set-ItemProperty -Path $addinPath -Name "Name" -Value $AddinName
            
            Write-Status "✓ Added add-in to: $addinPath" $Green
            Write-Status "  Location: $ManifestLocation" $Cyan
            $added = $true
        }
        catch {
            Write-Status "✗ Failed to add to registry: $($_.Exception.Message)" $Red
        }
    }
    
    return $added
}

function Test-ManifestAccessibility {
    param([string]$ManifestLocation)
    
    Write-Status "Testing manifest accessibility..." $Blue
    
    if ($ManifestLocation.StartsWith("http")) {
        # Test URL accessibility
        try {
            $response = Invoke-WebRequest -Uri $ManifestLocation -Method Head -TimeoutSec 10 -UseBasicParsing
            if ($response.StatusCode -eq 200) {
                Write-Status "✓ Manifest URL is accessible: $ManifestLocation" $Green
                return $true
            }
            else {
                Write-Status "✗ Manifest URL returned status: $($response.StatusCode)" $Red
                return $false
            }
        }
        catch {
            Write-Status "✗ Cannot access manifest URL: $($_.Exception.Message)" $Red
            return $false
        }
    }
    else {
        # Test local file
        if (Test-Path $ManifestLocation) {
            Write-Status "✓ Local manifest file found: $ManifestLocation" $Green
            return $true
        }
        else {
            Write-Status "✗ Local manifest file not found: $ManifestLocation" $Red
            return $false
        }
    }
}

function Validate-Manifest {
    param([string]$ManifestLocation)
    
    Write-Status "Validating manifest content..." $Blue
    
    try {
        if ($ManifestLocation.StartsWith("http")) {
            $content = Invoke-WebRequest -Uri $ManifestLocation -UseBasicParsing
            $xmlContent = $content.Content
        }
        else {
            $xmlContent = Get-Content -Path $ManifestLocation -Raw
        }
        
        # Parse XML
        $xml = [xml]$xmlContent
        
        # Basic validation
        if ($xml.OfficeApp) {
            $manifestId = $xml.OfficeApp.Id
            if ($manifestId -eq $AddinId) {
                Write-Status "✓ Manifest ID matches expected value" $Green
            }
            else {
                Write-Status "⚠ Warning: Manifest ID ($manifestId) doesn't match script configuration ($AddinId)" $Yellow
            }
            
            $displayName = $xml.OfficeApp.DisplayName.DefaultValue
            Write-Status "✓ Add-in name: $displayName" $Green
            
            return $true
        }
        else {
            Write-Status "✗ Invalid manifest format" $Red
            return $false
        }
    }
    catch {
        Write-Status "✗ Manifest validation failed: $($_.Exception.Message)" $Red
        return $false
    }
}

function Start-OutlookSafely {
    Write-Status "Starting Outlook..." $Blue
    
    try {
        # Try to start Outlook
        $outlookPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" -ErrorAction SilentlyContinue).Path
        
        if ($outlookPath -and (Test-Path $outlookPath)) {
            Start-Process -FilePath $outlookPath
            Write-Status "✓ Outlook started successfully" $Green
            Start-Sleep -Seconds 3
        }
        else {
            Write-Status "⚠ Could not find Outlook executable. Please start Outlook manually." $Yellow
        }
    }
    catch {
        Write-Status "⚠ Could not start Outlook automatically: $($_.Exception.Message)" $Yellow
        Write-Status "Please start Outlook manually." $Cyan
    }
}

function Show-Instructions {
    Write-Status "`n=== Next Steps ===" $Blue
    Write-Status "1. Open Outlook Desktop" $Cyan
    Write-Status "2. Look for the PromptEmail button in the ribbon" $Cyan
    Write-Status "3. If you don't see it, go to File > Manage Add-ins > My Add-ins" $Cyan
    Write-Status "4. The PromptEmail add-in should appear in the list" $Cyan
    Write-Status "5. If issues persist, try restarting Outlook completely" $Cyan
    Write-Status "`n=== Troubleshooting ===" $Blue
    Write-Status "• Check Windows Application Log for PromptEmail events" $Cyan
    Write-Status "• Verify manifest URL is accessible from your network" $Cyan
    Write-Status "• Ensure Outlook is updated to the latest version" $Cyan
    Write-Status "• Try running this script again with -Force parameter" $Cyan
}

# Main execution
function Main {
    Write-Status "PromptEmail Outlook Add-in Sideloader" $Blue
    Write-Status "========================================" $Blue
    
    # Check administrator privileges
    if (-not (Test-Administrator)) {
        Write-Status "✗ This script requires Administrator privileges for registry changes." $Red
        Write-Status "Please run PowerShell as Administrator and try again." $Yellow
        exit 1
    }
    
    # Handle cache clear only
    if ($ClearCacheOnly) {
        Stop-OutlookProcesses
        Clear-OutlookCache
        Write-Status "`n✓ Cache clearing completed!" $Green
        return
    }
    
    # Handle removal
    if ($RemoveAddin) {
        Stop-OutlookProcesses
        Remove-AddinFromRegistry
        Clear-OutlookCache
        Write-Status "`n✓ PromptEmail add-in removed successfully!" $Green
        Start-OutlookSafely
        return
    }
    
    # Determine manifest location
    $manifestLocation = ""
    if ($ManifestUrl -ne "") {
        $manifestLocation = $ManifestUrl
    }
    elseif (Test-Path $LocalManifestPath) {
        $manifestLocation = (Resolve-Path $LocalManifestPath).Path
    }
    else {
        $manifestLocation = $DefaultManifestUrl
    }
    
    Write-Status "Using manifest: $manifestLocation" $Cyan
    
    # Test manifest accessibility
    if (-not (Test-ManifestAccessibility -ManifestLocation $manifestLocation)) {
        if (-not $Force) {
            Write-Status "✗ Cannot access manifest. Use -Force to proceed anyway." $Red
            exit 1
        }
        else {
            Write-Status "⚠ Proceeding despite inaccessible manifest (Force mode)" $Yellow
        }
    }
    
    # Validate manifest
    if (-not (Validate-Manifest -ManifestLocation $manifestLocation)) {
        if (-not $Force) {
            Write-Status "✗ Manifest validation failed. Use -Force to proceed anyway." $Red
            exit 1
        }
        else {
            Write-Status "⚠ Proceeding despite validation failure (Force mode)" $Yellow
        }
    }
    
    # Stop Outlook
    Stop-OutlookProcesses
    
    # Remove existing registration
    Remove-AddinFromRegistry
    
    # Clear cache
    Clear-OutlookCache
    
    # Add new registration
    if (Add-AddinToRegistry -ManifestLocation $manifestLocation) {
        Write-Status "`n✓ PromptEmail add-in sideloaded successfully!" $Green
        
        # Start Outlook
        Start-OutlookSafely
        
        # Show instructions
        Show-Instructions
    }
    else {
        Write-Status "`n✗ Failed to sideload add-in" $Red
        exit 1
    }
}

# Help function
function Show-Help {
    Write-Host @"
PromptEmail Outlook Add-in Sideloader

USAGE:
    .\sideload-addin.ps1 [OPTIONS]

OPTIONS:
    -ManifestUrl <url>       URL to manifest.xml on S3 or web server
    -LocalManifestPath <path> Path to local manifest.xml file (default: .\manifest.xml)
    -RemoveAddin            Remove the add-in instead of installing
    -ClearCacheOnly         Only clear Outlook cache, don't install
    -Force                  Proceed even if manifest is inaccessible or invalid
    -Verbose                Show detailed output
    -Help                   Show this help message

EXAMPLES:
    # Install from S3 URL
    .\sideload-addin.ps1 -ManifestUrl "https://mybucket.s3.amazonaws.com/manifest.xml"
    
    # Install from local file
    .\sideload-addin.ps1 -LocalManifestPath "C:\MyProject\manifest.xml"
    
    # Remove the add-in
    .\sideload-addin.ps1 -RemoveAddin
    
    # Clear cache only
    .\sideload-addin.ps1 -ClearCacheOnly
    
    # Force install even if validation fails
    .\sideload-addin.ps1 -Force

REQUIREMENTS:
    • Windows PowerShell 5.1+ or PowerShell Core 7+
    • Administrator privileges
    • Outlook Desktop (Microsoft 365)
    • Network access to manifest location

NOTE: This script requires Administrator privileges to modify the Windows registry.
"@ -ForegroundColor Cyan
}

# Check for help parameter
if ($args -contains "-Help" -or $args -contains "/?" -or $args -contains "--help") {
    Show-Help
    exit 0
}

# Run main function
try {
    Main
}
catch {
    Write-Status "`n✗ Script execution failed: $($_.Exception.Message)" $Red
    Write-Status "Stack trace: $($_.ScriptStackTrace)" $Red
    exit 1
}
