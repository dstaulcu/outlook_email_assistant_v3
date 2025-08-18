# Outlook Add-in Production Sideload Helper
# This script helps sideload production Office add-ins from S3 in Outlook Desktop

param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestUrl,
    
    [string]$Environment = "Prd",
    [switch]$UseLocalManifest = $false,
    [switch]$SkipCacheClear = $false,
    [switch]$Help = $false
)

function Write-Status {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

function Show-Help {
    Write-Status "Outlook Add-in Production Sideload Helper" "Blue"
    Write-Status "=========================================" "Blue"
    Write-Status ""
    Write-Status "Helps sideload production Office add-ins from S3 or local manifests" "White"
    Write-Status ""
    Write-Status "Usage:" "Yellow"
    Write-Status "  .\outlook_addin_sideload_helper.ps1 -ManifestUrl <url>" "White"
    Write-Status "  .\outlook_addin_sideload_helper.ps1 -ManifestUrl <url> -Environment Dev" "White"
    Write-Status "  .\outlook_addin_sideload_helper.ps1 -ManifestUrl <path> -UseLocalManifest" "White"
    Write-Status ""
    Write-Status "Parameters:" "Yellow"
    Write-Status "  -ManifestUrl       S3 URL to the manifest.xml file (required)" "White"
    Write-Status "  -Environment       Environment name for logging (Dev/Prd, default: Prd)" "White"
    Write-Status "  -UseLocalManifest  Use local file path instead of S3 URL" "White"
    Write-Status "  -SkipCacheClear    Skip clearing Office add-in cache" "White"
    Write-Status "  -Help              Show this help message" "White"
    Write-Status ""
    Write-Status "Examples:" "Yellow"
    Write-Status "  # Production manifest from S3:" "Cyan"
    Write-Status "  .\outlook_addin_sideload_helper.ps1 -ManifestUrl 'https://your-bucket.s3.region.amazonaws.com/manifest.xml'" "White"
    Write-Status ""
    Write-Status "  # Development manifest from S3:" "Cyan"
    Write-Status "  .\outlook_addin_sideload_helper.ps1 -ManifestUrl 'https://dev-bucket.s3.region.amazonaws.com/manifest.xml' -Environment Dev" "White"
    Write-Status ""
    Write-Status "  # Local manifest file:" "Cyan"
    Write-Status "  .\outlook_addin_sideload_helper.ps1 -ManifestUrl '.\public\manifest.xml' -UseLocalManifest" "White"
    Write-Status ""
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
            Write-Status "✅ S3 manifest is accessible [Status: $($response.StatusCode)]" "Green"
            
            # Download and validate XML
            $manifestContent = Invoke-WebRequest -Uri $ManifestPath -TimeoutSec 10 -ErrorAction Stop
            [xml]$manifest = $manifestContent.Content
            $displayName = $manifest.OfficeApp.DisplayName.DefaultValue
            $id = $manifest.OfficeApp.Id
            Write-Status "✅ S3 manifest has valid XML" "Green"
            Write-Status "   Display Name: $displayName" "Gray"
            Write-Status "   ID: $id" "Gray"
            return $true
        } catch {
            Write-Status "❌ S3 manifest not accessible: $($_.Exception.Message)" "Red"
            Write-Status "💡 Check URL, S3 bucket permissions, and network connectivity" "Yellow"
            return $false
        }
    }
}

function Stop-OutlookProcess {
    Write-Status "📧 Checking Outlook process..." "Cyan"
    
    $outlookProcess = Get-Process "OUTLOOK" -ErrorAction SilentlyContinue
    if ($outlookProcess) {
        Write-Status "⚠️  Outlook is currently running (PID: $($outlookProcess.Id))" "Yellow"
        Write-Status "   Office add-ins require Outlook restart for proper sideloading" "White"
        Write-Status ""
        
        $response = Read-Host "Close Outlook automatically? (y/n)"
        if ($response -eq 'y' -or $response -eq 'Y') {
            try {
                Stop-Process -Name "OUTLOOK" -Force -ErrorAction Stop
                Start-Sleep -Seconds 3
                Write-Status "✅ Outlook closed successfully" "Green"
            } catch {
                Write-Status "❌ Failed to close Outlook: $($_.Exception.Message)" "Red"
                Write-Status "Please close Outlook manually and run this script again" "Yellow"
                return $false
            }
        } else {
            Write-Status "Please close Outlook manually before continuing" "Yellow"
            Read-Host "Press Enter when Outlook is closed..."
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

function Start-OutlookWithInstructions {
    param([string]$ManifestPath, [string]$Environment, [bool]$IsLocal)
    
    Write-Status ""
    Write-Status "🚀 Starting Outlook..." "Cyan"
    
    try {
        Start-Process "outlook.exe" -ErrorAction Stop
        Write-Status "✅ Outlook started successfully" "Green"
        Start-Sleep -Seconds 5
    } catch {
        Write-Status "❌ Failed to start Outlook: $($_.Exception.Message)" "Red"
        Write-Status "Please start Outlook manually" "Yellow"
    }
    
    # Show sideloading instructions
    Write-Status ""
    Write-Status "📝 Production Sideloading Instructions:" "Yellow"
    Write-Status "=======================================" "Yellow"
    Write-Status ""
    Write-Status "🔵 Step 1: Wait for Outlook to fully load" "White"
    Write-Status "🔵 Step 2: Open any email or create a new email" "White"
    Write-Status "🔵 Step 3: Look for 'Get Add-ins' or 'Store' in the ribbon" "White"
    Write-Status "   • In Message tab, or" "Gray"
    Write-Status "   • In Home tab" "Gray"
    Write-Status ""
    Write-Status "🔵 Step 4: Click 'Get Add-ins' → 'My Add-ins' (left sidebar)" "White"
    Write-Status "🔵 Step 5: Click 'Add a custom add-in' → 'Add from URL...'" "White"
    Write-Status ""
    Write-Status "🔵 Step 6: Enter the manifest URL:" "White"
    Write-Status "   $ManifestPath" "Cyan"
    Write-Status ""
    Write-Status "🔵 Step 7: Click 'OK' and accept security warnings" "White"
    Write-Status "🔵 Step 8: Look for the add-in button in the ribbon" "White"
    Write-Status ""
    
    if ($IsLocal) {
        Write-Status "⚠️  Note: You're using a local manifest file." "Yellow"
        Write-Status "   For production, use the S3 URL instead." "Yellow"
    } else {
        Write-Status "✅ Using production manifest from S3" "Green"
        Write-Status "   Environment: $Environment" "Gray"
    }
    
    Write-Status ""
    Write-Status "🔧 Troubleshooting:" "Yellow"
    Write-Status "• If 'Add from URL' is not available, try 'Add from file' with downloaded manifest" "Gray"
    Write-Status "• If add-in doesn't appear, check 'Optional Connected Experiences' in Outlook settings" "Gray"
    Write-Status "• If sideloading fails, run: .\outlook_addin_diagnostics.ps1" "Gray"
    Write-Status ""
}

function Download-ManifestForFileMethod {
    param([string]$ManifestUrl)
    
    Write-Status ""
    Write-Status "📥 Alternative: Download manifest for 'Add from file' method" "Yellow"
    
    $tempManifest = Join-Path $env:TEMP "outlook-addin-manifest.xml"
    
    try {
        Invoke-WebRequest -Uri $ManifestUrl -OutFile $tempManifest -TimeoutSec 10
        Write-Status "✅ Downloaded manifest to: $tempManifest" "Green"
        Write-Status ""
        Write-Status "If 'Add from URL' doesn't work, use 'Add from file' with:" "White"
        Write-Status "   $tempManifest" "Cyan"
        
        # Open file location
        $openLocation = Read-Host "Open download location? (y/n)"
        if ($openLocation -eq 'y' -or $openLocation -eq 'Y') {
            Start-Process "explorer.exe" "/select,`"$tempManifest`""
        }
        
    } catch {
        Write-Status "⚠️  Could not download manifest: $($_.Exception.Message)" "Yellow"
    }
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

# Main script execution
Write-Status "🚀 Outlook Add-in Production Sideload Helper" "Blue"
Write-Status "============================================" "Blue"
Write-Status ""
Write-Status "Environment: $Environment" "Gray"
Write-Status "Manifest Source: $ManifestUrl" "Gray"
Write-Status "Use Local File: $UseLocalManifest" "Gray"
Write-Status ""

# Step 1: Validate manifest
if (-not (Test-ManifestAccessibility -ManifestPath $ManifestUrl -IsLocal $UseLocalManifest)) {
    Write-Status ""
    Write-Status "❌ Cannot proceed - manifest validation failed" "Red"
    exit 1
}

# Step 2: Handle Outlook process
if (-not (Stop-OutlookProcess)) {
    exit 1
}

# Step 3: Clear cache
Clear-OfficeAddInCache

# Step 4: Start Outlook and show instructions
Start-OutlookWithInstructions -ManifestPath $ManifestUrl -Environment $Environment -IsLocal $UseLocalManifest

# Step 5: Offer backup method for S3 manifests
if (-not $UseLocalManifest) {
    Download-ManifestForFileMethod -ManifestUrl $ManifestUrl
}

Write-Status ""
Write-Status "✅ Sideload helper complete!" "Green"
Write-Status ""
Write-Status "📞 Need help? Run: .\outlook_addin_diagnostics.ps1" "Yellow"
