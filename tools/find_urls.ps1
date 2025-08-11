function Show-EmbeddedUrls {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory)]
        [string]$RootPath
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

        if ($matchFound) {
            $do_replacement = $true
            Write-Host "✅ Public file match found: $($matchFound -join ', ') for url: $($url.URL) in file: $($url.File)"
        }
        elseif ($url.URL -match $otherStrings) {
            Write-Host "✅ String match found for url: $($url.URL) in file: $($url.File)"
        }
        else {
            Write-Host "⚠️  No public file name or match found for url: $($url.URL) in file: $($url.File)"
        }

    }

}

Show-EmbeddedUrls -RootPath .\public