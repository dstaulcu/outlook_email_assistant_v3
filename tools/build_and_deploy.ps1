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
    
    # Patterns for runtime concatenation - these need to be updated
    $concatenationPatterns = @(
        # Pattern: (protocol.toLowerCase() === "https:" ? "https:" : "http:") + "//domain.com/path"
        '(?i)\(\s*[^)]*\.toLowerCase\(\)\s*===\s*[''"]https:[''"].*?\?\s*[''"]https:[''"].*?:\s*[''"]http:[''"].*?\)\s*\+\s*[''"]//([^''"]+)[''"]',
        
        # Pattern: "https:" ? "https:" : "http:") + "//domain.com/path"
        '(?i)[''"]https:[''"].*?\?\s*[''"]https:[''"].*?:\s*[''"]http:[''"].*?\)\s*\+\s*[''"]//([^''"]+)[''"]',
        
        # Pattern: window.location.protocol.toLowerCase() === "https:" ? "https:" : "http:") + "//domain.com
        '(?i)window\.location\.protocol[^)]*\)\s*\+\s*[''"]//([^''"]+)[''"]',
        
        # Pattern: (condition) + "//domain.com/path"  
        '(?i)\([^)]+\)\s*\+\s*[''"]//([^''"]+)[''"]'
    )
    
    $files = Get-ChildItem -Path $RootPath -Recurse -File
    $urlResults = @()
    $concatenationResults = @()
    
    foreach ($file in $files) {
        try {
            $content = Get-Content -Path $file.FullName -Raw -ErrorAction Stop
            
            # Find regular URLs
            $matches = [regex]::Matches($content, $urlPattern)
            foreach ($match in $matches) {
                $urlResults += [PSCustomObject]@{
                    File = $file.FullName
                    URL  = $match.Value
                    Type = "DirectURL"
                }
            }
            
            # Find concatenation patterns
            foreach ($pattern in $concatenationPatterns) {
                $concatMatches = [regex]::Matches($content, $pattern)
                foreach ($concatMatch in $concatMatches) {
                    if ($concatMatch.Groups.Count -gt 1) {
                        $domain = $concatMatch.Groups[1].Value
                        $concatenationResults += [PSCustomObject]@{
                            File = $file.FullName
                            OriginalPattern = $concatMatch.Value
                            Domain = $domain
                            Type = "Concatenation"
                        }
                    }
                }
            }
        }
        catch {
            Write-Warning "Could not read file: $($file.FullName)"
        }
    }

    Write-Host "Found $($urlResults.Count) direct URLs and $($concatenationResults.Count) concatenation patterns"

    # Process direct URLs (existing logic)
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

            write-status "`tReplacing direct URL:"
            write-status "`t  Original: `"$($url.URL)`""
            write-status "`t       New: `"$($newUriString)`""
            write-status "`t  Mapping: $($originalUri.Host)$($originalUri.AbsolutePath) → $($NewHost)$($AbsolutePath_new)"

            if ($PSCmdlet.ShouldProcess($url.File, "Replace '$($url.URL)' with '$newUriString'")) {
                $content = Get-Content -Path $url.File -Raw
                $contentUpdated = $content -replace [regex]::Escape($url.URL), $newUriString
                Set-Content -Path $url.File -Value $contentUpdated
            }
        }
    }

    # Process concatenation patterns (enhanced logic with path mapping)
    foreach ($concat in $concatenationResults) {
        Write-Host "Processing concatenation pattern in: $($concat.File)"
        Write-Host "  Original: $($concat.OriginalPattern)"
        Write-Host "  Domain: $($concat.Domain)"
        
        # Parse the domain and path from the concatenated URL
        $domainAndPath = $concat.Domain
        $pathParts = $domainAndPath -split "/"
        $actualDomain = $pathParts[0]
        $originalPath = if ($pathParts.Length -gt 1) { "/" + ($pathParts[1..($pathParts.Length-1)] -join "/") } else { "" }
        
        Write-Host "  Parsed Domain: $actualDomain"
        Write-Host "  Original Path: $originalPath"
        
        # Check if we should replace this domain and map to local files
        $shouldReplace = $false
        $localPath = ""
        
        # Map external resources to local paths in our S3 bucket
        switch -Regex ($actualDomain) {
            "ajax\.aspnetcdn\.com" {
                $shouldReplace = $true
                # Map ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js -> /online/ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js
                $localPath = "/online/ajax.aspnetcdn.com$originalPath"
                Write-Host "  → Mapped to local path: $localPath"
            }
            "alcdn\.msauth\.net" {
                $shouldReplace = $true
                # Map alcdn.msauth.net/browser-1p/... -> /online/alcdn.msauth.net/browser-1p/...
                $localPath = "/online/alcdn.msauth.net$originalPath"
                Write-Host "  → Mapped to local path: $localPath"
            }
            "appsforoffice\.microsoft\.com" {
                $shouldReplace = $true
                # Map appsforoffice.microsoft.com/lib/... -> /online/appsforoffice.microsoft.com/lib/...
                $localPath = "/online/appsforoffice.microsoft.com$originalPath"
                Write-Host "  → Mapped to local path: $localPath"
            }
            "raw\.githubusercontent\.com" {
                $shouldReplace = $true
                # Map raw.githubusercontent.com/... -> /online/raw.githubusercontent.com/...
                $localPath = "/online/raw.githubusercontent.com$originalPath"
                Write-Host "  → Mapped to local path: $localPath"
            }
            default {
                Write-Host "  → No mapping defined for domain: $actualDomain"
            }
        }
        
        if ($shouldReplace -and $localPath -and $PSCmdlet.ShouldProcess($concat.File, "Replace concatenation pattern")) {
            $content = Get-Content -Path $concat.File -Raw
            
            # Create the new pattern with our host and the mapped local path
            $newDomainAndPath = $NewHost + $localPath
            $newPattern = $concat.OriginalPattern -replace [regex]::Escape($concat.Domain), $newDomainAndPath
            
            write-status "`tReplacing concatenation pattern:"
            write-status "`t  Original: `"$($concat.OriginalPattern)`""
            write-status "`t       New: `"$($newPattern)`""
            write-status "`t  Mapping: $actualDomain$originalPath → $NewHost$localPath"
            
            $contentUpdated = $content -replace [regex]::Escape($concat.OriginalPattern), $newPattern
            Set-Content -Path $concat.File -Value $contentUpdated
        }
        elseif ($shouldReplace -and -not $localPath) {
            Write-Host "  ⚠️  Domain recognized but no local path mapping defined" -ForegroundColor Yellow
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
    
    # Check AWS credentials and permissions
    Write-Status "Validating AWS credentials and permissions..." 'Yellow'
    try {
        # Test basic AWS credential validity
        $whoamiOutput = aws sts get-caller-identity 2>&1
        if ($LASTEXITCODE -eq 0) {
            $identity = $whoamiOutput | ConvertFrom-Json
            Write-Status "✓ AWS credentials are valid" 'Green'
            Write-Status "  Account: $($identity.Account)" 'Cyan'
            Write-Status "  User/Role: $($identity.Arn.Split('/')[-1])" 'Cyan'
        }
        else {
            Write-Status "✗ AWS credentials are invalid or expired." 'Red'
            Write-Status "Error details: $whoamiOutput" 'Red'
            Write-Status "Please run 'aws configure' or refresh your temporary credentials." 'Yellow'
            exit 1
        }
        
        # Test S3 access to the target bucket
        $bucketName = $envConfig.s3Uri.host
        Write-Status "Testing S3 bucket access: $bucketName..." 'Yellow'
        
        $s3TestOutput = aws s3 ls "s3://$bucketName" --region $envConfig.region 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Status "✓ S3 bucket access confirmed" 'Green'
        }
        else {
            Write-Status "✗ Cannot access S3 bucket: $bucketName" 'Red'
            Write-Status "Error details: $s3TestOutput" 'Red'
            Write-Status "Please verify:" 'Yellow'
            Write-Status "  - Bucket exists and you have access" 'Yellow'
            Write-Status "  - Your AWS credentials have S3 permissions" 'Yellow'
            Write-Status "  - The specified region ($($envConfig.region)) is correct" 'Yellow'
            exit 1
        }
    }
    catch {
        Write-Status "✗ Exception during AWS credential validation: $_" 'Red'
        Write-Status "Please verify your AWS configuration and try again." 'Yellow'
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
                $clearOutput = aws s3 rm "$S3BaseUrl/" --recursive --region $Region 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Status "✓ Cleared all items from S3 bucket: $BucketName" 'Green'
                }
                else {
                    Write-Status "✗ Failed to clear S3 bucket: $BucketName" 'Red'
                    Write-Status "AWS CLI output: $clearOutput" 'Red'
                    
                    # Check if this looks like a credential issue
                    if ($clearOutput -match "expired|invalid|credentials|token|authentication|authorization") {
                        Write-Status "This appears to be an AWS credential issue." 'Yellow'
                        Write-Status "Please refresh your AWS credentials and try again." 'Yellow'
                    }
                    exit 1
                }
            }
            catch {
                Write-Status "✗ Exception during S3 bucket clear: $_" 'Red'
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
                $uploadOutput = $null
                if ($contentType) {
                    $uploadOutput = aws s3 cp $localPath "$S3BaseUrl/$s3Key" --region $Region --content-type $contentType 2>&1
                }
                else {
                    $uploadOutput = aws s3 cp $localPath "$S3BaseUrl/$s3Key" --region $Region 2>&1
                }
                
                if ($LASTEXITCODE -eq 0) {
                    Write-Status "✓ Uploaded $s3Key" 'Green'
                }
                else {
                    Write-Status "✗ Failed to upload $s3Key" 'Red'
                    Write-Status "AWS CLI output: $uploadOutput" 'Red'
                    
                    # Check if this looks like a credential issue
                    if ($uploadOutput -match "expired|invalid|credentials|token|authentication|authorization") {
                        Write-Status "This appears to be an AWS credential issue." 'Yellow'
                        Write-Status "Your credentials may have expired during the upload process." 'Yellow'
                        Write-Status "Please refresh your AWS credentials and try again." 'Yellow'
                        exit 1
                    }
                    exit 1
                }
            }
            catch {
                Write-Status "✗ Exception during upload of $s3Key : $_" 'Red'
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

# Ensure required npm packages are installed
Write-Status "Checking webpack installation..." "Blue"
try {
    # First check if webpack is already available and working
    $webpackVersion = $null
    $webpackCliVersion = $null
    $webpackWorking = $false
    
    try {
        $webpackVersion = & webpack --version 2>$null
        $webpackCliVersion = & webpack-cli --version 2>$null
        if ($webpackVersion -and $webpackCliVersion) {
            $webpackWorking = $true
            Write-Status "✓ webpack is already available (webpack: $webpackVersion, webpack-cli: $webpackCliVersion)" "Green"
        }
    }
    catch {
        Write-Status "webpack not found or not working, will install..." "Yellow"
    }
    
    # Install webpack if not working
    if (-not $webpackWorking) {
        Write-Status "Installing webpack and webpack-cli..." "Yellow"
        $installOutput = & npm install -g webpack webpack-cli --save-dev 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Status "✓ Global webpack packages installed successfully" "Green"
            
            # Verify the installation worked
            try {
                $newWebpackVersion = & webpack --version 2>$null
                $newWebpackCliVersion = & webpack-cli --version 2>$null
                if ($newWebpackVersion -and $newWebpackCliVersion) {
                    Write-Status "✓ Verified webpack installation (webpack: $newWebpackVersion, webpack-cli: $newWebpackCliVersion)" "Green"
                }
                else {
                    Write-Status "⚠️  Webpack installed but verification failed, continuing anyway..." "Yellow"
                }
            }
            catch {
                Write-Status "⚠️  Could not verify webpack installation, continuing anyway..." "Yellow"
            }
        }
        else {
            Write-Status "⚠️  Global webpack install had issues, trying local install..." "Yellow"
            Write-Host $installOutput
            
            # Try local install as fallback
            $localInstallOutput = & npm install webpack webpack-cli --save-dev 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Status "✓ Local webpack packages installed successfully" "Green"
            }
            else {
                Write-Status "✗ Both global and local webpack installs failed. See details below:" "Red"
                Write-Host $localInstallOutput
                exit 1
            }
        }
    }
    
    # Install/update local project dependencies
    Write-Status "Installing/updating local project dependencies..." "Yellow"
    $depsOutput = & npm install 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Status "✓ Local dependencies installed successfully" "Green"
    }
    else {
        Write-Status "✗ Failed to install local dependencies. See details below:" "Red"
        Write-Host $depsOutput
        exit 1
    }
}
catch {
    Write-Status "✗ Exception during package installation: $_" "Red"
    exit 1
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

# copy src\telemetry-config.json to .\public
$srcTelemetryConfig = Join-Path $srcDir 'telemetry-config.json'
$publicTelemetryConfig = Join-Path $publicDir 'telemetry-config.json'
if ($DryRun) {
    Write-Status "[DryRun] Would copy telemetry-config.json to public/" 'Yellow'
}
elseif (Test-Path $srcTelemetryConfig) {
    Copy-Item $srcTelemetryConfig $publicTelemetryConfig -Force
    Write-Status "✓ Copied telemetry-config.json to public/" 'Green'
}
else {
    Write-Status "⚠️ src/telemetry-config.json not found - telemetry configuration will use defaults" 'Yellow'
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
