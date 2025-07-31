


param(
    [string]$ManifestSource = "C:\code\outlook_email_assistant_v3\manifest.xml"
)

# Kill all Outlook processes
Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

# Clear Office add-in cache (for Office 2016/2019/365)
$cachePath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
if (Test-Path $cachePath) {
    $maxRetries = 5
    $retryDelay = 2 # seconds
    $attempt = 0
    $success = $false
    while (-not $success -and $attempt -lt $maxRetries) {
        try {
            Remove-Item "$cachePath\*" -Recurse -Force -ErrorAction Stop
            $success = $true
        } catch {
            Write-Host "Attempt $($attempt + 1) to clear cache failed: $($_.Exception.Message)" -ForegroundColor Yellow
            Start-Sleep -Seconds $retryDelay
            $attempt++
        }
    }
    if (-not $success) {
        Write-Host "Some cache files could not be deleted after $maxRetries attempts. You may need to close other processes or run as administrator." -ForegroundColor Red
    }
    Write-Host "Office add-in cache cleared: $cachePath"
} else {
    Write-Host "Cache path not found: $cachePath"
}




# Copy manifest.xml from local file and sideload if ManifestSource is provided
if ($ManifestSource) {
    $localManifest = Join-Path $env:USERPROFILE "manifest.xml"
    try {
        Write-Host "Copying manifest.xml from $ManifestSource to $localManifest ..."
        Copy-Item -Path $ManifestSource -Destination $localManifest -Force
        if (-not (Test-Path $localManifest)) {
            Write-Host "Failed to copy manifest.xml from $ManifestSource" -ForegroundColor Red
            exit 1
        }

        $regBase = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
        $manifestValue = (Resolve-Path $localManifest).Path
        New-ItemProperty -Path $regBase -Name "ManifestPath" -Value $manifestValue -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $regBase -Name "EnableLocalManifest" -Value 1 -PropertyType DWord -Force | Out-Null
        Write-Host "Manifest sideloaded to registry: $regBase" -ForegroundColor Green
        Write-Host "  ManifestPath = $manifestValue"
        Write-Host "  EnableLocalManifest = 1"
    } catch {
        Write-Host "Failed to copy or sideload manifest: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

Write-Host "Outlook closed and cache cleared. Launching Outlook..."
Start-Process OUTLOOK.EXE
