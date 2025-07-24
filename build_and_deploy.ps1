# PromptEmail Outlook Add-in Build & Deployment Script
# This script builds the project and safely uploads only build assets to S3

param(
    [string]$BucketName = "293354421824-outlook-email-assistant",
    [string]$Region = "us-east-1",
    [switch]$DryRun = $false
)


function Write-Status {
    param([string]$Message, [string]$Color = "White")
    $validColors = @('Black','DarkBlue','DarkGreen','DarkCyan','DarkRed','DarkMagenta','DarkYellow','Gray','DarkGray','Blue','Green','Cyan','Red','Magenta','Yellow','White')
    if (-not $Color -or ($validColors -notcontains $Color)) {
        $Color = 'White'
    }
    Write-Host $Message -ForegroundColor $Color
}

Write-Status "Starting build process..." "Blue"

# Run npm build and capture output
$buildOutput = $null
$buildError = $null
$buildSucceeded = $false
try {
    $buildOutput = & npm run build 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Status "✓ Build completed successfully" "Green"
        $buildSucceeded = $true
    } else {
        Write-Status "✗ Build failed. See details below:" "Red"
        Write-Host $buildOutput
        exit 1
    }
} catch {
    Write-Status "✗ Exception during build: $_" "Red"
    if ($buildOutput) { Write-Host $buildOutput }
    exit 1
}

# Configuration
$BuildDir = "public"
$ManifestFile = "manifest.xml"
$RequiredFiles = @(
    "$BuildDir/index.html",
    "$BuildDir/taskpane.html",
    "$BuildDir/taskpane.bundle.js",
    "$BuildDir/commands.bundle.js",
    "$BuildDir/taskpane.css"
)

# Colors for output
 # Colors are now set in Write-Status, so these variables are not needed

function Test-Prerequisites {
    Write-Status "Checking prerequisites..." $Blue
    
    # Check if AWS CLI is installed
    try {
        aws --version | Out-Null
        Write-Status "✓ AWS CLI found" $Green
    } catch {
        Write-Status "✗ AWS CLI not found. Please install AWS CLI." $Red
        exit 1
    }
    
    # Check if build directory exists
    if (-not (Test-Path $BuildDir)) {
        Write-Status "✗ Build directory '$BuildDir' not found. Run 'npm run build' first." $Red
        exit 1
    }
    
    # Check required files
    foreach ($file in $RequiredFiles) {
        if (-not (Test-Path $file)) {
            Write-Status "✗ Required file not found: $file" $Red
            exit 1
        }
    }
    
    Write-Status "✓ All prerequisites met" $Green
}

function Backup-Manifest {
    if (Test-Path $ManifestFile) {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $backupFile = "manifest-backup-$timestamp.xml"
        Copy-Item $ManifestFile $backupFile
        Write-Status "✓ Manifest backed up to $backupFile" $Green
    }
}

function Deploy-Assets {
    Write-Status "Deploying assets to S3 bucket: $BucketName" $Blue
    
    if ($DryRun) {
        Write-Status "DRY RUN MODE - No files will be uploaded" $Yellow
    }
    
    # Upload HTML files with correct content-type
    $htmlFiles = Get-ChildItem -Path $BuildDir -Filter "*.html"
    foreach ($file in $htmlFiles) {
        $s3Key = $file.Name
        $localPath = $file.FullName
        
        if ($DryRun) {
            Write-Status "Would upload: $localPath -> s3://$BucketName/$s3Key" $Yellow
        }
        else {
            try {
                aws s3 cp $localPath "s3://$BucketName/$s3Key" --content-type "text/html" --region $Region
                Write-Status "✓ Uploaded $s3Key" $Green
            }
            catch {
                Write-Status "✗ Failed to upload $s3Key" $Red
                Write-Status $_.Exception.Message $Red
            }
        }
    }
    
    # Upload JS files
    $jsFiles = Get-ChildItem -Path $BuildDir -Filter "*.js"
    foreach ($file in $jsFiles) {
        $s3Key = $file.Name
        $localPath = $file.FullName
        
        if ($DryRun) {
            Write-Status "Would upload: $localPath -> s3://$BucketName/$s3Key" $Yellow
        }
        else {
            try {
                aws s3 cp $localPath "s3://$BucketName/$s3Key" --content-type "application/javascript" --region $Region
                Write-Status "✓ Uploaded $s3Key" $Green
            }
            catch {
                Write-Status "✗ Failed to upload $s3Key" $Red
                Write-Status $_.Exception.Message $Red
            }
        }
    }
    
    # Upload CSS files
    $cssFiles = Get-ChildItem -Path $BuildDir -Filter "*.css"
    foreach ($file in $cssFiles) {
        $s3Key = $file.Name
        $localPath = $file.FullName
        
        if ($DryRun) {
            Write-Status "Would upload: $localPath -> s3://$BucketName/$s3Key" $Yellow
        }
        else {
            try {
                aws s3 cp $localPath "s3://$BucketName/$s3Key" --content-type "text/css" --region $Region
                Write-Status "✓ Uploaded $s3Key" $Green
            }
            catch {
                Write-Status "✗ Failed to upload $s3Key" $Red
                Write-Status $_.Exception.Message $Red
            }
        }
    }
    
    # Upload icon files if they exist
    $iconDir = "$BuildDir/icons"
    if (Test-Path $iconDir) {
        $iconFiles = Get-ChildItem -Path $iconDir -File
        foreach ($file in $iconFiles) {
            $s3Key = "icons/$($file.Name)"
            $localPath = $file.FullName
            if ($DryRun) {
                Write-Status "Would upload: $localPath -> s3://$BucketName/$s3Key" $Yellow
            } else {
                try {
                    aws s3 cp $localPath "s3://$BucketName/$s3Key" --region $Region
                    Write-Status "✓ Uploaded $s3Key" $Green
                } catch {
                    Write-Status "✗ Failed to upload $s3Key" $Red
                    Write-Status $_.Exception.Message $Red
                }
            }
        }
    }
}

function Verify-Deployment {
    if ($DryRun) {
        Write-Status "Skipping verification in dry run mode" $Yellow
        return
    }
    
    Write-Status "Verifying deployment..." $Blue
    
    $baseUrl = "https://$BucketName.s3.amazonaws.com"
    
    # Test index.html accessibility
    try {
        $response = Invoke-WebRequest -Uri "$baseUrl/index.html" -Method Head -TimeoutSec 10
        if ($response.StatusCode -eq 200) {
            Write-Status "✓ index.html is accessible" $Green
        }
        else {
            Write-Status "✗ index.html returned status: $($response.StatusCode)" $Red
        }
    }
    catch {
        Write-Status "✗ Failed to verify index.html accessibility" $Red
        Write-Status $_.Exception.Message $Red
    }
}

function Show-NextSteps {
    Write-Status "`nDeployment Summary:" $Blue
    Write-Status "Bucket: $BucketName" $Blue
    Write-Status "Base URL: https://$BucketName.s3.amazonaws.com" $Blue
    Write-Status "`nNext Steps:" $Blue
    Write-Status "1. Update manifest.xml with the correct S3 URLs" $Yellow
    Write-Status "2. Validate the manifest: npm run validate-manifest" $Yellow
    Write-Status "3. Sideload the manifest in Outlook" $Yellow
    Write-Status "4. Test the add-in functionality" $Yellow
}

# Main execution
Write-Status "PromptEmail Outlook Add-in Deployment" $Blue
Write-Status "=====================================" $Blue

# Run deployment steps
Test-Prerequisites
Backup-Manifest
Deploy-Assets
Verify-Deployment
Show-NextSteps

Write-Status "`nDeployment completed!" $Green
