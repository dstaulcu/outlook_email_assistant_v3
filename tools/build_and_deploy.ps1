param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Dev', 'Prd')]
    [string]$Environment,
    [switch]$DryRun,
    [switch]$Force
)

function Update-EmbeddedUrls {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory)]
        [string]$RootPath,
        [string]$NewHost,
        [string]$NewScheme = "https"
    )
    
    # Comprehensive URL pattern to match HTTP/HTTPS/S3 URLs
    $urlPattern = '(?i)(?:https?://|s3://)[a-zA-Z0-9\-\.]+(?:\.[a-zA-Z0-9\-\.]+)*(?:\:[0-9]+)?(?:/[^\s"''<>]*)?'
    
    # (Removed stray $updatedContent normalization from here)
    $files = Get-ChildItem -Path $RootPath -Recurse -File
    $urlResults = @()
    foreach ($file in $files) {
        try {
            $content = Get-Content -Path $file.FullName -Raw -ErrorAction Stop
            $matches = [regex]::Matches($content, $urlPattern)
            foreach ($match in $matches) {
                $urlResults += [PSCustomObject]@{
                    File = $file.FullName
                    URL  = $match.Value
                }
            }
        }
        catch {
            Write-Warning "Could not read file: $($file.FullName)"
        }
    }

    # evaluate each url
    foreach ($url in $urlResults) {

        # Check if any name is contained in the URL
        $fileNames = $files | Select-Object -ExpandProperty name
        $otherStrings = "(outlook-email-assistant)"
        $matchFound = $fileNames | Where-Object { $url.URL -like "*$_*" }

        # prepare url for POTENTIAL replacement
        $do_replacement = $false
        if ($url.URL -like "s3://*") {
            $originalUri = [System.Uri]("https://" + $url.URL.Substring(6))
            $scheme = "s3"
        }
        else {
            $originalUri = [System.Uri]$url.URL
            $scheme = $originalUri.Scheme
        }        

        if ($matchFound) {
            $do_replacement = $true
            Write-Host "Public file match found: $($matchFound -join ', ') for url: $($url.URL) in file: $($url.File)"
            # we need to get the fullpath to file from matching filename
            $matchFoundFullName = ($files | Where-Object { $_.name -eq $matchFound }[0]).FullName
            # Normalize path for reference in URL
            $AbsolutePath_new = $matchFoundFullName -replace ".*\\public", ""
            $AbsolutePath_new = $AbsolutePath_new -replace "\\", "/"
            $AbsolutePath_new = $AbsolutePath_new -replace "^([^/])", "/$1"
        }
        elseif ($url.URL -match $otherStrings) {
            $do_replacement = $true
            Write-Host "String match found for url: $($url.URL) in file: $($url.File)"
            $AbsolutePath_new = $originalUri.AbsolutePath
        }
        # Normalize double slashes (except after protocol) for all replacements
        if ($do_replacement -and $AbsolutePath_new) {
            $AbsolutePath_new = $AbsolutePath_new -replace '^/+', '/'
            $AbsolutePath_new = $AbsolutePath_new -replace '://', '___PROTOCOL_SLASH___'
            $AbsolutePath_new = $AbsolutePath_new -replace '/{2,}', '/'
            $AbsolutePath_new = $AbsolutePath_new -replace '___PROTOCOL_SLASH___', '://'
        }
        else {
            $do_replacement = $false
            if ($url.url -match '\.[^\.]{0,3}$') {
                Write-status "⚠️  No public file name or match found in SEEMINGLY file-oriented url: $($url.URL) in file: $($url.File)" 'yellow'
            } else {
                # Write-status "⚠️  No public file name or match found in SEEMLINGLY folder-oriented url: $($url.URL) in file: $($url.File)"
            }
        }

        if ($do_replacement -eq $true) {
            # Build new URI
            $builder = New-Object System.UriBuilder $originalUri
            $builder.Host = $NewHost
            $builder.Path = $AbsolutePath_new
            $builder.Query = $originalUri.Query
            $newUri = $builder.Uri            


            # Restore s3:// scheme if needed
            $newUriString = $newUri.AbsoluteUri
            if ($scheme -eq "s3") {
                $newUriString = $newUriString -replace "^https://", "s3://"
            }

            # Normalize double slashes in the final URL (except after protocol)
            $newUriString = $newUriString -replace '://', '___PROTOCOL_SLASH___'
            $newUriString = $newUriString -replace '/{2,}', '/'
            $newUriString = $newUriString -replace '___PROTOCOL_SLASH___', '://'

            write-status "`tReplacing originalUri: `"$($url.URL)`" with:"
            write-status "`t               newUri: `"$($newUriString)`""

            if ($PSCmdlet.ShouldProcess($url.File, "Replace '$($url.URL)' with '$newUriString'")) {
                $content = Get-Content -Path $url.File -Raw
                $contentUpdated = $content -replace [regex]::Escape($url.URL), $newUriString
                Set-Content -Path $url.File -Value $contentUpdated
            }

        }

    }

}

# Update references in online.orig and copy to online
function Write-Status {
    param([string]$Message, [string]$Color = "White")
    $validColors = @('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')
    if (-not $Color -or ($validColors -notcontains $Color)) {
        $Color = 'White'
    }
    Write-Host $Message -ForegroundColor $Color
}

function Test-Prerequisites {
    Write-Status "Checking prerequisites..." 'Blue'
    
    # Check if AWS CLI is installed
    try {
        aws --version | Out-Null
        Write-Status "✓ AWS CLI found" 'Green'
    }
    catch {
        Write-Status "✗ AWS CLI not found. Please install AWS CLI." 'Red'
        exit 1
    }
        
    Write-Status "✓ All prerequisites met" 'Green'
}

function Deploy-Assets {
    param(
        [Parameter(Mandatory)]
        [string]$BuildDir
    )
    
    Write-Status "Deploying assets to S3 bucket: $BucketName (region: $Region)" $Blue

    if ($DryRun) {
        Write-Status "DRY RUN MODE - No files will be uploaded" 
    } elseif (-not $DryRun) {
        if ($Force) {
            $confirm = 'YES'
        } else {
            $confirm = Read-Host "Are you sure you want to delete ALL contents from the S3 bucket '$BucketName'? This cannot be undone. Type 'YES' to confirm"
        }
        if ($confirm -eq 'YES') {
            try {
                Write-Status "Clearing all items from S3 bucket: $BucketName"
                aws s3 rm "$S3BaseUrl/" --recursive --region $Region
                Write-Status "✓ Cleared all items from S3 bucket: $BucketName" 'Green'
            }
            catch {
                Write-Status "✗ Failed to clear S3 bucket: $BucketName" 'Red'
                Write-Status $_.Exception.Message $Red
                exit 1
            }
        } else {
            Write-Status "Aborted: S3 bucket will NOT be cleared. Deployment cancelled." 'Red'
            exit 1
        }
    }

    # Upload all files in public
    $allFiles = Get-ChildItem -Path $BuildDir -Recurse -File

    Write-Status "Found $($allFiles.Count) files to upload from: $BuildDir" $Blue

    foreach ($file in $allFiles) {
        # Compute S3 key relative to $BuildDir
        $s3Key = $file.FullName.Substring($BuildDir.Length + 1) -replace '\\', '/'
        $localPath = $file.FullName

        # Determine content-type by extension
        $ext = [System.IO.Path]::GetExtension($file.Name).ToLower()
        switch ($ext) {
            ".html" { $contentType = "text/html" }
            ".js" { $contentType = "application/javascript" }
            ".json" { $contentType = "application/json" }
            ".xml" { $contentType = "text/xml" }
            ".css" { $contentType = "text/css" }
            ".png" { $contentType = "image/png" }
            default { $contentType = $null }
        }

        if ($DryRun) {
            Write-Status "Would upload: $localPath -> $S3BaseUrl/$s3Key ($contentType)" 'Yellow'
        }
        else {
            try {
                # Compose aws s3 cp command
                if ($contentType) {
                    aws s3 cp $localPath "$S3BaseUrl/$s3Key" --region $Region --content-type $contentType
                }
                else {
                    aws s3 cp $localPath "$S3BaseUrl/$s3Key" --region $Region
                }
                Write-Status "✓ Uploaded $s3Key" 'Green'
            }
            catch {
                Write-Status "✗ Failed to upload $s3Key" 'Red'
                Write-Status $_.Exception.Message 'Red'
            }
        }
    }
}

function Test-Deployment {

    Write-Status "Verifying deployment..." 'Blue'
    $baseUrl = $HttpBaseUrl
    # Test index.html accessibility
    try {
        $response = Invoke-WebRequest -Uri "$baseUrl/index.html" -Method Head -TimeoutSec 10
        if ($response.StatusCode -eq 200) {
            Write-Status "✓ index.html is accessible" 'Green'
        }
        else {
            Write-Status "✗ index.html returned status: $($response.StatusCode)" 'Red'
        }
    }
    catch {
        Write-Status "✗ Failed to verify index.html accessibility" 'Red'
        Write-Status $_.Exception.Message 'Red'
    }
}

function Show-NextSteps {
    Write-Status "`nDeployment Summary:" 'Blue'
    Write-Status "Environment: $Environment" 'Blue'
    Write-Status "Bucket: $BucketName" 'Blue'
    Write-Status "Region: $Region" 'Blue'
    Write-Status "Base URL: $HttpBaseUrl" 'Blue'
    Write-Status "`nNext Steps:" 'Blue'
    Write-Status "1. Validate the manifest: npm run validate-manifest" 
    Write-Status "2. Sideload the manifest in Outlook" 
    Write-Status "3. Test the add-in functionality" 
}

# Main execution
Write-Status "PromptEmail Outlook Add-in Deployment" 'Blue'
Write-Status "=====================================" 'Blue'

# Ensure $ProjectRoot is set before any usage
$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$srcDir = Join-Path $ProjectRoot 'src'
$publicDir = Join-Path $ProjectRoot 'public'

# Initialize environment and config 
$ConfigPath = Join-Path $PSScriptRoot 'deployment-environments.json'
if (Test-Path $ConfigPath) {
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    $envConfig = $config.environments.$Environment
    if ($envConfig.publicUri -and $envConfig.s3Uri -and $envConfig.region) {
        $Region = $envConfig.region
        $PublicBaseUrl = "$($envConfig.publicUri.protocol)://$($envConfig.publicUri.host)"
        $S3BaseUrl = "$($envConfig.s3Uri.protocol)://$($envConfig.s3Uri.host)"
        # For compatibility with rest of script
        $HttpBaseUrl = $PublicBaseUrl
        $BucketName = $envConfig.s3Uri.host.Split('.')[0]
    }
    else {
        throw "Missing publicUri, s3Uri, or region for environment '$Environment' in deployment-environments.json."
    }
}
else {
    throw "deployment-environments.json not found in tools/."
}


# Run prerequisites check as early as possible
Test-Prerequisites

# clear the public folder if it already exists
if (test-path -path $publicDir) {
    remove-item -Path $publicDir -recurse
    mkdir -Path $publicDir | Out-Null
}

# Run npm build and capture output
Write-Status "Starting build process..." "Blue"
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

# copy src\manifest.xml to .\public
$srcManifest = Join-Path $srcDir 'manifest.xml'
$publicManifest = Join-Path $publicDir 'manifest.xml'
if ($DryRun) {
    Write-Status "[DryRun] Would copy manifest.xml to public/" 'Yellow'
}
elseif (Test-Path $srcManifest) {
    Copy-Item $srcManifest $publicManifest -Force
    Write-Status "✓ Copied manifest.xml to public/" 'Green'
}
else {
    Write-Status "✗ src/manifest.xml not found!" 'Red'
}

# copy src\online to .\public\online
$srcOnline = Join-Path $srcDir 'online'
$publicOnline = Join-Path $publicDir 'online'
if ($DryRun) {
    Write-Status "[DryRun] Would copy online.orig to public/online/" 'Yellow'
}
elseif (Test-Path $srcOnline) {
    if (Test-Path $publicOnline) {
        Remove-Item $publicOnline -Recurse -Force
    }
    Copy-Item $srcOnline $publicOnline -Recurse -Force
    Write-Status "✓ Copied online.orig to public/online/" 'Green'
}
else {
    Write-Status "✗ src/online.orig not found!" 'Red'
}


# Update online references before deployment
if ($DryRun) {
    Write-Status "[DryRun] Would update embedded URLs and normalize manifest.xml" 'Yellow'
}
else {
    Write-Status "Updating Urls in public folder files..." 'Blue'
    try {
        # Update embedded URLs in public files using new URI-spec config
        Update-EmbeddedUrls -RootPath (Join-Path $ProjectRoot 'public') -NewHost $envConfig.publicUri.host -NewScheme $envConfig.publicUri.protocol
        Write-Status "✓ Urls in public folder files updated" 'Green'

        # Post-process manifest.xml to normalize all URLs (remove double slashes except after protocol)
        $publicManifestPath = Join-Path $publicDir 'manifest.xml'
        if (Test-Path $publicManifestPath) {
            $manifestContent = Get-Content $publicManifestPath -Raw
            $normalizedManifest = $manifestContent -replace '://', '___PROTOCOL_SLASH___'
            $normalizedManifest = $normalizedManifest -replace '/{2,}', '/'
            $normalizedManifest = $normalizedManifest -replace '___PROTOCOL_SLASH___', '://'
            if ($manifestContent -ne $normalizedManifest) {
                Set-Content $publicManifestPath $normalizedManifest
                Write-Status "✓ Normalized double slashes in manifest.xml URLs" 'Green'
            }
        }
    }
    catch {
        Write-Status "✗ Failed to update Urls in publif folder files: $_" 'Red'
        exit 1
    }
}

# deploy content of .\public to target web server (e.g. s3)
Deploy-Assets -BuildDir $publicDir
# verify index.html is web-accessible in web server
Test-Deployment
Show-NextSteps
