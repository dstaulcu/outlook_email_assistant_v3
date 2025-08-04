param(
    [ValidateSet('Dev', 'Prd')]
    [string]$Environment = 'Dev',
    [switch]$DryRun = $false
)

# Ensure $ProjectRoot is set before any usage
$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

# Update manifest.xml URLs for the selected environment
function Update-ManifestUrls {
    $srcManifestPath = Join-Path $ProjectRoot 'src/manifest.xml'
    $publicManifestPath = Join-Path $ProjectRoot 'public/manifest.xml'
    if (-not (Test-Path $srcManifestPath)) {
        Write-Status "src/manifest.xml not found, skipping URL update." $Yellow
        return
    }
    # Copy manifest.xml to public before updating URLs
    Copy-Item $srcManifestPath $publicManifestPath -Force
    $manifestContent = Get-Content $publicManifestPath -Raw
    # Replace all S3 URLs (including s3-website endpoints) with the correct base for this environment
    $pattern = 'https?://[a-zA-Z0-9\-]+\.(s3|s3-website)([.-][a-z0-9-]+)?\.amazonaws\.com'
    $updatedContent = $manifestContent -replace $pattern, $HttpBaseUrl
    if ($manifestContent -ne $updatedContent) {
        Set-Content $publicManifestPath $updatedContent
        Write-Status "Updated manifest.xml URLs for $Environment environment." $Green
    } else {
        Write-Status "No manifest.xml URLs needed updating for $Environment environment." $Blue
    }
}



# Load environment config and construct URLs dynamically
$ConfigPath = Join-Path $PSScriptRoot 'deployment-environments.json'
if (Test-Path $ConfigPath) {
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    $envConfig = $config.environments.$Environment
    if ($envConfig.bucketName -and $envConfig.region -and $envConfig.publicDnsSuffix -and $envConfig.s3DnsSuffix) {
        $BucketName = $envConfig.bucketName
        $Region = $envConfig.region
        $PublicDnsSuffix = $envConfig.publicDnsSuffix
        $S3DnsSuffix = $envConfig.s3DnsSuffix
        $PublicBaseUrl = "https://$BucketName.$PublicDnsSuffix"
        $S3BucketUri = "s3://$BucketName"
        $S3RestUrl = "https://$BucketName.$S3DnsSuffix"
    } else {
        throw "Missing bucketName, region, publicDnsSuffix, or s3DnsSuffix for environment '$Environment' in deployment-environments.json."
    }
} else {
    throw "deployment-environments.json not found in tools/."
}

# For compatibility with rest of script
$HttpBaseUrl = $PublicBaseUrl
$S3BaseUrl = $S3BucketUri

$BucketName = ($S3BucketUri -replace '^s3://', '')

# Update references in online.orig and copy to online
function Update-OnlineReferences {
    param(
        [string]$S3BaseUrlOverride = $null
    )
    $S3BaseUrlFinal = if ($S3BaseUrlOverride) { $S3BaseUrlOverride } else { "$HttpBaseUrl/online" }
    $origDir = Join-Path $ProjectRoot 'src/online.orig'
    $destDir = Join-Path $ProjectRoot 'public/online'

    Write-Status "Copying files from $origDir to $destDir ..." $Blue
    if (Test-Path $destDir) {
        Remove-Item $destDir -Recurse -Force
    }
    Copy-Item $origDir $destDir -Recurse -Force

    $files = Get-ChildItem -Path $destDir -Recurse -Include *.js,*.html,*.css
    foreach ($file in $files) {
        $content = Get-Content $file.FullName -Raw
        $content = $content -replace '//appsforoffice\.microsoft\.com', "$S3BaseUrlFinal/appsforoffice.microsoft.com"
        $content = $content -replace '//ajax\.aspnetcdn\.com', "$S3BaseUrlFinal/ajax.aspnetcdn.com"
        Set-Content $file.FullName $content
        Write-Status "Updated references in $($file.FullName)" $Green
    }
    Write-Status "All references updated in 'public/online'." $Green
}


function Write-Status {
    param([string]$Message, [string]$Color = "White")
    $validColors = @('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')
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
    }
    else {
        Write-Status "✗ Build failed. See details below:" "Red"
        Write-Host $buildOutput
        exit 1
    }
}
catch {
    Write-Status "✗ Exception during build: $_" "Red"
    if ($buildOutput) { Write-Host $buildOutput }
    exit 1
}


# Ensure styles.css is copied to public before deployment
$srcStyles = Join-Path $ProjectRoot 'src/assets/css/styles.css'
$destStyles = Join-Path $ProjectRoot 'public/styles.css'
if (Test-Path $srcStyles) {
    Copy-Item $srcStyles $destStyles -Force
    Write-Status "✓ Copied styles.css to public/" $Green
} else {
    Write-Status "✗ src/assets/css/styles.css not found!" $Red
    exit 1
}

# Ensure default-models.json is copied to public before deployment
$srcDefaultModels = Join-Path $ProjectRoot 'src/default-models.json'
$destDefaultModels = Join-Path $ProjectRoot 'public/default-models.json'
if (Test-Path $srcDefaultModels) {
    Copy-Item $srcDefaultModels $destDefaultModels -Force
    Write-Status "✓ Copied default-models.json to public/" $Green
} else {
    Write-Status "✗ src/default-models.json not found!" $Red
    exit 1
}

# Ensure default-providers.json is copied to public before deployment
$srcDefaultProviders = Join-Path $ProjectRoot 'src/default-providers.json'
$destDefaultProviders = Join-Path $ProjectRoot 'public/default-providers.json'
if (Test-Path $srcDefaultProviders) {
    Copy-Item $srcDefaultProviders $destDefaultProviders -Force
    Write-Status "✓ Copied default-providers.json to public/" $Green
} else {
    Write-Status "✗ src/default-providers.json not found!" $Red
    exit 1
}

# Configuration
$BuildDir = Join-Path $ProjectRoot 'public'
$ManifestFile = Join-Path $BuildDir 'manifest.xml'
$RequiredFiles = @(
    (Join-Path $BuildDir 'index.html'),
    (Join-Path $BuildDir 'taskpane.html'),
    (Join-Path $BuildDir 'taskpane.bundle.js'),
    (Join-Path $BuildDir 'commands.bundle.js'),
    (Join-Path $BuildDir 'taskpane.css')
)

# Colors for output
# Colors are now set in Write-Status, so these variables are not needed

function Test-Prerequisites {
    Write-Status "Checking prerequisites..." $Blue
    
    # Check if AWS CLI is installed
    try {
        aws --version | Out-Null
        Write-Status "✓ AWS CLI found" $Green
    }
    catch {
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
    Write-Status "Deploying assets to S3 bucket: $BucketName (region: $Region)" $Blue
    
    if ($DryRun) {
        Write-Status "DRY RUN MODE - No files will be uploaded" $Yellow
    }
    
    # Upload HTML files with correct content-type
    $htmlFiles = Get-ChildItem -Path $BuildDir -Filter "*.html"
    foreach ($file in $htmlFiles) {
        $s3Key = $file.Name
        $localPath = $file.FullName
        
        if ($DryRun) {
            Write-Status "Would upload: $localPath -> $S3BaseUrl/$s3Key" $Yellow
        }
        else {
            try {
                aws s3 cp $localPath "$S3BaseUrl/$s3Key" --content-type "text/html" --region $Region
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
            Write-Status "Would upload: $localPath -> $S3BaseUrl/$s3Key" $Yellow
        }
        else {
            try {
                aws s3 cp $localPath "$S3BaseUrl/$s3Key" --content-type "application/javascript" --region $Region
                Write-Status "✓ Uploaded $s3Key" $Green
            }
            catch {
                Write-Status "✗ Failed to upload $s3Key" $Red
                Write-Status $_.Exception.Message $Red
            }
        }
    }

    # Upload default-models.json and default-providers.json
    $jsonFiles = @('default-models.json', 'default-providers.json')
    foreach ($jsonFile in $jsonFiles) {
        $localPath = Join-Path $BuildDir $jsonFile
        if (Test-Path $localPath) {
            if ($DryRun) {
                Write-Status "Would upload: $localPath -> $S3BaseUrl/$jsonFile" $Yellow
            } else {
                try {
                    aws s3 cp $localPath "$S3BaseUrl/$jsonFile" --content-type "application/json" --region $Region
                    Write-Status "✓ Uploaded $jsonFile" $Green
                } catch {
                    Write-Status "✗ Failed to upload $jsonFile" $Red
                    Write-Status $_.Exception.Message $Red
                }
            }
        } else {
            Write-Status "✗ $localPath not found, skipping upload." $Red
        }
    }

    # Upload manifest.xml
    $manifestPath = Join-Path $BuildDir 'manifest.xml'
    if (Test-Path $manifestPath) {
        if ($DryRun) {
            Write-Status "Would upload: $manifestPath -> $S3BaseUrl/manifest.xml" $Yellow
        } else {
            try {
                aws s3 cp $manifestPath "$S3BaseUrl/manifest.xml" --content-type "text/xml" --region $Region
                Write-Status "✓ Uploaded manifest.xml" $Green
            } catch {
                Write-Status "✗ Failed to upload manifest.xml" $Red
                Write-Status $_.Exception.Message $Red
            }
        }
    } else {
        Write-Status "✗ $manifestPath not found, skipping upload." $Red
    }
    
    # Upload CSS files
    $cssFiles = Get-ChildItem -Path $BuildDir -Filter "*.css"
    foreach ($file in $cssFiles) {
        $s3Key = $file.Name
        $localPath = $file.FullName
        
        if ($DryRun) {
            Write-Status "Would upload: $localPath -> $S3BaseUrl/$s3Key" $Yellow
        }
        else {
            try {
                aws s3 cp $localPath "$S3BaseUrl/$s3Key" --content-type "text/css" --region $Region
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
                Write-Status "Would upload: $localPath -> $S3BaseUrl/$s3Key" $Yellow
            }
            else {
                try {
                    aws s3 cp $localPath "$S3BaseUrl/$s3Key" --region $Region
                    Write-Status "✓ Uploaded $s3Key" $Green
                }
                catch {
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
    $baseUrl = $HttpBaseUrl
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
    Write-Status "Environment: $Environment" $Blue
    Write-Status "Bucket: $BucketName" $Blue
    Write-Status "Region: $Region" $Blue
    Write-Status "Base URL: $HttpBaseUrl" $Blue
    Write-Status "`nNext Steps:" $Blue
    Write-Status "1. Validate the manifest: npm run validate-manifest" $Yellow
    Write-Status "2. Sideload the manifest in Outlook" $Yellow
    Write-Status "3. Test the add-in functionality" $Yellow
}

# Main execution
Write-Status "PromptEmail Outlook Add-in Deployment" $Blue
Write-Status "=====================================" $Blue

# Run deployment steps
Test-Prerequisites
Backup-Manifest

# Update manifest.xml URLs for this environment
Update-ManifestUrls

# Update online references before deployment
Write-Status "Updating online references..." $Blue
try {
    Update-OnlineReferences
    Write-Status "✓ Online references updated" $Green
} catch {
    Write-Status "✗ Failed to update online references: $_" $Red
    exit 1
}

Deploy-Assets

# Upload all files in ./public/online to S3, preserving folder structure
$onlineDir = Join-Path $ProjectRoot 'public/online'
if (Test-Path $onlineDir) {
    Write-Status "Uploading ./public/online assets to S3..." $Blue
    $onlineFiles = Get-ChildItem -Path $onlineDir -Recurse -File
    foreach ($file in $onlineFiles) {
        $relativePath = $file.FullName.Substring($onlineDir.Length + 1) -replace '\\','/'
        $s3Key = "online/$relativePath"
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

Verify-Deployment
Show-NextSteps

Write-Status "`nDeployment completed!" $Green

