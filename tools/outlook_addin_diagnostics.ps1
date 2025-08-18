# Office Add-in Comprehensive Diagnostics Toolkit
# Consolidated from debug-addin-loading.ps1, diagnose-addin.ps1, and diagnose-office-environment.ps1

param(
    [switch]$Help = $false
)

function Write-Status {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

function Show-Help {
    Write-Status "Office Add-in Comprehensive Diagnostics Toolkit" "Blue"
    Write-Status "===============================================" "Blue"
    Write-Status ""
    Write-Status "A comprehensive diagnostic and debugging tool for Office add-ins" "White"
    Write-Status "Combines environment analysis, debugging controls, and troubleshooting" "White"
    Write-Status ""
    Write-Status "Usage:" "Yellow"
    Write-Status "  .\office-addin-diagnostics.ps1" "White"
    Write-Status "  .\office-addin-diagnostics.ps1 -Help" "White"
    Write-Status ""
}

# === ENVIRONMENT ANALYSIS FUNCTIONS ===

function Get-OfficeEnvironmentInfo {
    Write-Status "🏢 Office Environment Analysis" "Cyan"
    Write-Status "=============================" "Cyan"
    
    # Office version detection
    Write-Status ""
    Write-Status "📋 Office Version Information:" "Yellow"
    try {
        $outlookVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue).VersionToReport
        if ($outlookVersion) {
            Write-Status "  ✅ Office Click-to-Run Version: $outlookVersion" "Green"
        } else {
            $outlookVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot" -ErrorAction SilentlyContinue).Path
            if ($outlookVersion) {
                Write-Status "  ✅ Office MSI Installation Path: $outlookVersion" "Green"
            } else {
                Write-Status "  ❌ Cannot determine Office version" "Red"
            }
        }
    } catch {
        Write-Status "  ❌ Error detecting Office version: $($_.Exception.Message)" "Red"
    }

    # Update channel detection
    try {
        $updateChannel = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue).CDNBaseUrl
        if ($updateChannel -like "*insidersfast*") {
            Write-Status "  📺 Update Channel: Insider Fast" "Yellow"
        } elseif ($updateChannel -like "*insidersslow*") {
            Write-Status "  📺 Update Channel: Insider Slow" "Yellow"
        } elseif ($updateChannel -like "*monthly*") {
            Write-Status "  📺 Update Channel: Monthly Enterprise" "Green"
        } elseif ($updateChannel) {
            Write-Status "  📺 Update Channel: Custom/Enterprise" "White"
        } else {
            Write-Status "  ❓ Update Channel: Unknown" "Yellow"
        }
    } catch {
        Write-Status "  ❌ Cannot determine update channel" "Red"
    }
    
    # Outlook process check
    Write-Status ""
    Write-Status "📧 Outlook Process Status:" "Yellow"
    $outlookProcess = Get-Process "OUTLOOK" -ErrorAction SilentlyContinue
    if ($outlookProcess) {
        Write-Status "  ✅ Outlook is running (PID: $($outlookProcess.Id))" "Green"
        Write-Status "  📊 Memory Usage: $([math]::Round($outlookProcess.WorkingSet64/1MB, 2)) MB" "White"
        
        # Check if running in safe mode
        try {
            $commandLine = (Get-CimInstance Win32_Process -Filter "ProcessId = $($outlookProcess.Id)").CommandLine
            if ($commandLine -like "*safe*") {
                Write-Status "  ⚠️  WARNING: Outlook is running in Safe Mode - this disables add-ins!" "Red"
            } else {
                Write-Status "  ✅ Outlook is running in normal mode" "Green"
            }
        } catch {
            Write-Status "  ❓ Cannot determine Outlook startup mode" "Yellow"
        }
    } else {
        Write-Status "  ❌ Outlook is not currently running" "Red"
    }
}

function Check-CriticalOfficeSettings {
    Write-Status ""
    Write-Status "❗ CRITICAL: Optional Connected Experiences Check" "Red"
    Write-Status "===============================================" "Red"
    Write-Status "This is the #1 cause of add-in buttons not appearing!" "Yellow"
    Write-Status ""
    Write-Status "🔍 Manual Check Required:" "White"
    Write-Status "1. Open Outlook → File → Options → Trust Center → Trust Center Settings → Privacy Options" "Cyan"
    Write-Status "2. Verify 'Optional connected experiences' is CHECKED/ENABLED" "Cyan"
    Write-Status "3. If disabled, enable it and restart Outlook" "Cyan"
    
    # Check add-in registry settings
    Write-Status ""
    Write-Status "📋 Add-in Registry Configuration:" "Yellow"
    
    $addinRegPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs"
    try {
        if (Test-Path $addinRegPath) {
            Write-Status "  ✅ Add-in registry path exists" "Green"
            $catalogs = Get-ChildItem $addinRegPath -ErrorAction SilentlyContinue
            Write-Status "  📁 Trusted catalogs: $($catalogs.Count)" "White"
            
            if ($catalogs.Count -gt 0) {
                Write-Status "  📦 Registered Add-in Catalogs:" "White"
                foreach ($catalog in $catalogs) {
                    $props = Get-ItemProperty $catalog.PSPath -ErrorAction SilentlyContinue
                    if ($props.Id) {
                        Write-Status "    - ID: $($props.Id)" "Gray"
                        if ($props.Url) {
                            Write-Status "      URL: $($props.Url)" "Gray"
                        }
                    }
                }
            }
        } else {
            Write-Status "  ⚠️  No trusted catalogs found" "Yellow"
        }
    } catch {
        Write-Status "  ❌ Error checking add-in registry: $($_.Exception.Message)" "Red"
    }

    # Check for restrictive policies
    Write-Status ""
    Write-Status "🔒 Office Security Policies:" "Yellow"
    
    $policyPaths = @(
        @{Path = "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\outlook\security"; Name = "User Outlook Security"},
        @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\outlook\security"; Name = "Machine Outlook Security"},
        @{Path = "HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\common\security"; Name = "User Office Security"},
        @{Path = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\common\security"; Name = "Machine Office Security"}
    )

    $foundRestrictive = $false
    foreach ($policy in $policyPaths) {
        if (Test-Path $policy.Path) {
            try {
                $settings = Get-ItemProperty $policy.Path -ErrorAction SilentlyContinue
                $restrictiveSettings = $settings.PSObject.Properties | Where-Object { 
                    $_.Name -like "*addin*" -or $_.Name -like "*web*" -or 
                    $_.Name -like "*extension*" -or $_.Name -like "*disable*"
                }
                
                if ($restrictiveSettings) {
                    $foundRestrictive = $true
                    Write-Status "  ⚠️  Found restrictive settings in $($policy.Name):" "Yellow"
                    foreach ($setting in $restrictiveSettings) {
                        Write-Status "    - $($setting.Name) = $($setting.Value)" "Red"
                    }
                }
            } catch {
                Write-Status "  ❌ Access denied to $($policy.Name)" "Red"
            }
        }
    }
    
    if (-not $foundRestrictive) {
        Write-Status "  ✅ No restrictive Office policies found" "Green"
    }

    Write-Status ""
    Write-Status "💡 Additional Manual Checks:" "Yellow"
    Write-Status "• File → Options → Trust Center → Add-ins → Uncheck 'Require signed add-ins'" "Cyan"
    Write-Status "• File → Options → Trust Center → Add-ins → Uncheck 'Disable all add-ins'" "Cyan"
    Write-Status "• Verify internet connectivity to: appsforoffice.microsoft.com" "Cyan"
}

function Check-AddinRegistryEntries {
    Write-Status ""
    Write-Status "🔍 Add-in Registry Analysis" "Cyan"
    Write-Status "===========================" "Cyan"
    
    $registryPaths = @(
        "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer",
        "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs",
        "HKLM:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
    )
    
    foreach ($path in $registryPaths) {
        Write-Status ""
        Write-Status "📂 Checking: $path" "White"
        if (Test-Path $path) {
            try {
                $keys = Get-ChildItem $path -ErrorAction SilentlyContinue
                if ($keys) {
                    Write-Status "  ✅ Found entries:" "Green"
                    foreach ($key in $keys) {
                        Write-Status "    - $($key.PSChildName)" "Gray"
                        
                        # Get properties of each entry
                        $props = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
                        if ($props.Id) {
                            Write-Status "      ID: $($props.Id)" "Gray"
                        }
                        if ($props.Url) {
                            Write-Status "      URL: $($props.Url)" "Gray"
                        }
                    }
                } else {
                    Write-Status "  ℹ️  Path exists but no entries found" "Yellow"
                }
            } catch {
                Write-Status "  ❌ Access denied: $($_.Exception.Message)" "Red"
            }
        } else {
            Write-Status "  ❌ Path does not exist" "Red"
        }
    }
}

function Check-FileSystemCache {
    Write-Status ""
    Write-Status "🗂️  Add-in File System Cache Analysis" "Cyan"
    Write-Status "=====================================" "Cyan"
    
    $cachePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
        "$env:LOCALAPPDATA\Microsoft\Office\Wef",
        "$env:APPDATA\Microsoft\Office\16.0\Wef"
    )
    
    foreach ($path in $cachePaths) {
        Write-Status ""
        Write-Status "📂 Checking: $path" "White"
        if (Test-Path $path) {
            try {
                $files = Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue
                Write-Status "  ✅ Found $($files.Count) cached files" "Green"
                
                if ($files.Count -gt 0) {
                    $manifests = $files | Where-Object { $_.Name -like "*manifest*" -or $_.Extension -eq ".xml" }
                    if ($manifests) {
                        Write-Status "  📄 Manifest-related files:" "Yellow"
                        foreach ($manifest in $manifests) {
                            Write-Status "    - $($manifest.Name) ($($manifest.Length) bytes)" "Gray"
                        }
                    }
                    
                    # Check log files
                    $logFiles = $files | Where-Object { $_.Extension -eq ".log" }
                    if ($logFiles) {
                        Write-Status "  📊 Log files found: $($logFiles.Count)" "Yellow"
                    }
                }
            } catch {
                Write-Status "  ❌ Access denied: $($_.Exception.Message)" "Red"
            }
        } else {
            Write-Status "  ❌ Path does not exist" "Red"
        }
    }
}

function Test-OutlookCOMAccess {
    Write-Status ""
    Write-Status "🔌 Outlook COM Object Testing" "Cyan"
    Write-Status "==============================" "Cyan"
    
    $outlookProcess = Get-Process "OUTLOOK" -ErrorAction SilentlyContinue
    if (-not $outlookProcess) {
        Write-Status "  ❌ Outlook is not running - COM testing skipped" "Red"
        return
    }

    try {
        Write-Status "  🔄 Creating Outlook COM Object..." "White"
        $outlook = New-Object -ComObject Outlook.Application
        Write-Status "  ✅ COM Object created successfully" "Green"
        
        # Test add-ins collection access
        try {
            $addins = $outlook.COMAddIns
            Write-Status "  ✅ COMAddIns collection accessible ($($addins.Count) add-ins)" "Green"
            
            if ($addins.Count -gt 0) {
                Write-Status "  📦 Registered COM Add-ins:" "Yellow"
                foreach ($addin in $addins) {
                    $status = if ($addin.Connect) { "Connected" } else { "Disconnected" }
                    Write-Status "    - $($addin.Description) [$status]" "White"
                }
            }
        } catch {
            Write-Status "  ❌ Cannot access COMAddIns: $($_.Exception.Message)" "Red"
        }
        
        # Release COM object properly
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        
    } catch {
        Write-Status "  ❌ Cannot create COM Object: $($_.Exception.Message)" "Red"
        Write-Status "  💡 Try running PowerShell as Administrator" "Yellow"
    }
}

function Test-NetworkConnectivity {
    Write-Status ""
    Write-Status "🌐 Network Connectivity Testing" "Cyan"
    Write-Status "================================" "Cyan"
    
    # Test Office.js CDN
    Write-Status ""
    Write-Status "📡 Testing Office.js CDN:" "Yellow"
    $officejsUrl = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
    try {
        $response = Invoke-WebRequest -Uri $officejsUrl -Method Head -TimeoutSec 10 -ErrorAction Stop
        Write-Status "  ✅ Office.js CDN accessible [Status: $($response.StatusCode)]" "Green"
    } catch {
        Write-Status "  ❌ Office.js CDN failed: $($_.Exception.Message)" "Red"
    }
    
    # Test your S3 assets
    Write-Status ""
    Write-Status "📦 Testing Add-in Assets:" "Yellow"
    $assetUrls = @(
        "https://293354421824-outlook-email-assistant-prd.s3.us-east-1.amazonaws.com/manifest.xml",
        "https://293354421824-outlook-email-assistant-prd.s3.us-east-1.amazonaws.com/taskpane.html",
        "https://293354421824-outlook-email-assistant-prd.s3.us-east-1.amazonaws.com/icons/icon-32.png"
    )
    
    foreach ($url in $assetUrls) {
        $fileName = [System.IO.Path]::GetFileName($url)
        try {
            $response = Invoke-WebRequest -Uri $url -Method Head -TimeoutSec 10 -ErrorAction Stop
            Write-Status "  ✅ $fileName [Status: $($response.StatusCode)]" "Green"
        } catch {
            Write-Status "  ❌ $fileName [Error: $($_.Exception.Message)]" "Red"
        }
    }
}

function Test-ManifestFiles {
    Write-Status ""
    Write-Status "📋 Manifest File Validation" "Cyan"
    Write-Status "============================" "Cyan"
    
    $manifestPaths = @(
        ".\public\manifest.xml",
        ".\src\manifest.xml",
        ".\manifest.xml"
    )
    
    foreach ($manifestPath in $manifestPaths) {
        Write-Status ""
        Write-Status "📄 Checking: $manifestPath" "White"
        if (Test-Path $manifestPath) {
            try {
                [xml]$xml = Get-Content $manifestPath
                
                # Basic validation
                $id = $xml.OfficeApp.Id
                $version = $xml.OfficeApp.Version
                $displayName = $xml.OfficeApp.DisplayName.DefaultValue
                
                Write-Status "  ✅ Valid XML structure" "Green"
                Write-Status "    ID: $id" "Gray"
                Write-Status "    Version: $version" "Gray"
                Write-Status "    Display Name: $displayName" "Gray"
                
                # Check for common issues
                $issues = @()
                
                # Check for HTTPS URLs
                $urls = $xml.SelectNodes("//*[@DefaultValue]") | Where-Object { $_.DefaultValue -like "http://*" }
                if ($urls) {
                    $issues += "Contains HTTP URLs (should be HTTPS)"
                }
                
                # Check for missing required elements
                if (-not $xml.OfficeApp.Requirements) {
                    $issues += "Missing Requirements section"
                }
                
                if ($issues.Count -eq 0) {
                    Write-Status "  ✅ No obvious issues found" "Green"
                } else {
                    Write-Status "  ⚠️  Potential issues:" "Yellow"
                    foreach ($issue in $issues) {
                        Write-Status "    - $issue" "Red"
                    }
                }
                
            } catch {
                Write-Status "  ❌ XML parsing failed: $($_.Exception.Message)" "Red"
            }
        } else {
            Write-Status "  ❌ File not found" "Red"
        }
    }
}

# === DEBUGGING CONTROL FUNCTIONS ===

function Enable-OfficeAddinDebugging {
    Write-Status ""
    Write-Status "🔧 Enabling Office Add-in Debugging" "Cyan"
    Write-Status "====================================" "Cyan"
    
    # Enable runtime logging in registry
    $regPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    
    try {
        # Enable logging
        Set-ItemProperty -Path $regPath -Name "EnableLogging" -Value 1 -Type DWord
        Set-ItemProperty -Path $regPath -Name "LogLevel" -Value 0 -Type DWord  # 0 = Verbose
        
        # Enable runtime logging
        $runtimeLogPath = "$regPath\RuntimeLogging"
        if (-not (Test-Path $runtimeLogPath)) {
            New-Item -Path $runtimeLogPath -Force | Out-Null
        }
        Set-ItemProperty -Path $runtimeLogPath -Name "EnableLogging" -Value 1 -Type DWord
        
        Write-Status "  ✅ Debugging enabled in registry" "Green"
        Write-Status "  📁 Logs will be written to: $env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\Logs" "Gray"
    } catch {
        Write-Status "  ❌ Error enabling debugging: $($_.Exception.Message)" "Red"
    }
}

function Disable-OfficeAddinDebugging {
    Write-Status ""
    Write-Status "🔧 Disabling Office Add-in Debugging" "Cyan"
    Write-Status "=====================================" "Cyan"
    
    $regPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer"
    if (Test-Path $regPath) {
        try {
            # Disable logging
            Set-ItemProperty -Path $regPath -Name "EnableLogging" -Value 0 -Type DWord
            Set-ItemProperty -Path $regPath -Name "LogLevel" -Value 1 -Type DWord  # 1 = Error only
            
            # Disable runtime logging
            $runtimeLogPath = "$regPath\RuntimeLogging"
            if (Test-Path $runtimeLogPath) {
                Set-ItemProperty -Path $runtimeLogPath -Name "EnableLogging" -Value 0 -Type DWord
            }
            
            Write-Status "  ✅ Debugging disabled in registry" "Green"
        } catch {
            Write-Status "  ❌ Error disabling debugging: $($_.Exception.Message)" "Red"
        }
    } else {
        Write-Status "  ℹ️  No debugging registry entries found" "Yellow"
    }
}

function Show-OfficeAddinLogs {
    Write-Status ""
    Write-Status "📊 Office Add-in Logs" "Cyan"
    Write-Status "======================" "Cyan"
    
    $logDir = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\Logs"
    
    if (-not (Test-Path $logDir)) {
        Write-Status "  ❌ Log directory not found: $logDir" "Red"
        Write-Status "  💡 Try enabling debugging first" "Yellow"
        return
    }
    
    $logFiles = Get-ChildItem $logDir -Filter "*.log" -ErrorAction SilentlyContinue
    if ($logFiles.Count -eq 0) {
        Write-Status "  ⚠️  No log files found" "Yellow"
        Write-Status "  💡 Try loading an add-in to generate logs" "Gray"
        return
    }
    
    Write-Status "  📁 Found $($logFiles.Count) log files:" "White"
    foreach ($logFile in $logFiles | Sort-Object LastWriteTime -Descending) {
        $age = (Get-Date) - $logFile.LastWriteTime
        $ageText = if ($age.TotalHours -lt 1) { "$($age.Minutes)m ago" } else { "$([math]::Round($age.TotalHours, 1))h ago" }
        Write-Status "    📄 $($logFile.Name) ($('{0:N0}' -f $logFile.Length) bytes, $ageText)" "Gray"
    }
    
    # Show recent entries from the most recent log
    $mostRecentLog = $logFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Write-Status ""
    Write-Status "  📖 Recent entries from $($mostRecentLog.Name):" "White"
    try {
        $recentEntries = Get-Content $mostRecentLog.FullName -Tail 10 -ErrorAction SilentlyContinue
        if ($recentEntries) {
            foreach ($entry in $recentEntries) {
                Write-Status "    $entry" "Gray"
            }
        }
    } catch {
        Write-Status "    ❌ Could not read log file (may be locked)" "Red"
    }
}

function Start-AddinLoadMonitoring {
    Write-Status ""
    Write-Status "🔄 Starting Real-time Add-in Load Monitoring" "Cyan"
    Write-Status "=============================================" "Cyan"
    Write-Status "Press Ctrl+C to stop monitoring" "Yellow"
    Write-Status ""
    
    # Create log directory if it doesn't exist
    $logDir = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\Logs"
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        Write-Status "📁 Created log directory: $logDir" "Green"
    }
    
    # Monitor file system changes in log directory
    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $logDir
    $watcher.Filter = "*.log"
    $watcher.IncludeSubdirectories = $false
    $watcher.EnableRaisingEvents = $true
    
    # Event handler for new log entries
    $action = {
        $path = $Event.SourceEventArgs.FullPath
        $changeType = $Event.SourceEventArgs.ChangeType
        $timestamp = Get-Date -Format 'HH:mm:ss'
        
        Write-Host "[$timestamp] Log activity: $changeType - $([System.IO.Path]::GetFileName($path))" -ForegroundColor Yellow
        
        if ($changeType -eq 'Changed') {
            # Try to read new content
            Start-Sleep -Milliseconds 100
            try {
                $newContent = Get-Content $path -Tail 1 -ErrorAction SilentlyContinue
                if ($newContent) {
                    Write-Host "  Content: $newContent" -ForegroundColor White
                }
            } catch {
                Write-Host "  Could not read new content (file locked)" -ForegroundColor Gray
            }
        }
    }
    
    # Register event handlers
    Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action $action | Out-Null
    Register-ObjectEvent -InputObject $watcher -EventName "Changed" -Action $action | Out-Null
    Register-ObjectEvent -InputObject $watcher -EventName "Deleted" -Action $action | Out-Null
    
    Write-Status "👀 Monitoring active... Now try to load/reload your add-in" "Green"
    
    try {
        # Keep running until Ctrl+C
        while ($true) {
            Start-Sleep -Seconds 1
        }
    } finally {
        # Clean up
        $watcher.EnableRaisingEvents = $false
        $watcher.Dispose()
        Get-EventSubscriber | Unregister-Event
        Write-Status ""
        Write-Status "🛑 Monitoring stopped" "Yellow"
    }
}

function Clean-DebugArtifacts {
    Write-Status ""
    Write-Status "🧹 Cleaning Debug Artifacts" "Cyan"
    Write-Status "============================" "Cyan"
    
    # Clear log files
    $logDir = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\Logs"
    if (Test-Path $logDir) {
        try {
            $logFiles = Get-ChildItem -Path $logDir -Filter "*.log"
            $clearedCount = 0
            foreach ($logFile in $logFiles) {
                try {
                    Remove-Item $logFile.FullName -Force -ErrorAction SilentlyContinue
                    $clearedCount++
                } catch {
                    Write-Status "  ⚠️  Could not remove $($logFile.Name) (may be in use)" "Yellow"
                }
            }
            Write-Status "  ✅ Cleared $clearedCount log files" "Green"
        } catch {
            Write-Status "  ❌ Error clearing log files: $($_.Exception.Message)" "Red"
        }
    }
    
    # Clear cache
    $cacheDir = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
    if (Test-Path $cacheDir) {
        try {
            $cacheDirs = Get-ChildItem -Path $cacheDir -Directory | Where-Object { $_.Name -ne "Logs" }
            $clearedDirs = 0
            foreach ($dir in $cacheDirs) {
                try {
                    Remove-Item $dir.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    $clearedDirs++
                } catch {
                    Write-Status "  ⚠️  Could not remove $($dir.Name) (may be in use)" "Yellow"
                }
            }
            if ($clearedDirs -gt 0) {
                Write-Status "  ✅ Cleared $clearedDirs cache directories" "Green"
            }
        } catch {
            Write-Status "  ❌ Error clearing cache: $($_.Exception.Message)" "Red"
        }
    }
    
    Write-Status "  ℹ️  Restart Outlook to ensure all changes take effect" "Yellow"
}

# === MAIN MENU AND EXECUTION ===

if ($Help) {
    Show-Help
    exit 0
}

# Main execution
Write-Status "🚀 Office Add-in Comprehensive Diagnostics Toolkit" "Blue"
Write-Status "===================================================" "Blue"
Write-Status ""

# Show main menu
Write-Status "Select diagnostic action:" "Yellow"
Write-Status ""
Write-Status "📊 Analysis Options:" "Green"
Write-Status "1. Complete environment analysis" "White"
Write-Status "2. Check critical Office settings (Optional Connected Experiences)" "White"
Write-Status "3. Analyze add-in registry entries" "White"
Write-Status "4. Check file system cache" "White"
Write-Status "5. Test Outlook COM access" "White"
Write-Status "6. Test network connectivity" "White"
Write-Status "7. Validate manifest files" "White"
Write-Status ""
Write-Status "🔧 Debugging Controls:" "Green"
Write-Status "8. Enable Office add-in debugging" "White"
Write-Status "9. Disable Office add-in debugging" "White"
Write-Status "10. Show existing debug logs" "White"
Write-Status "11. Start real-time monitoring" "White"
Write-Status ""
Write-Status "🧹 Maintenance:" "Green"
Write-Status "12. Clean debug artifacts & cache" "White"
Write-Status "13. Run comprehensive analysis (all checks)" "White"
Write-Status ""

$choice = Read-Host "Enter choice (1-13)"

switch ($choice) {
    "1" { Get-OfficeEnvironmentInfo }
    "2" { Check-CriticalOfficeSettings }
    "3" { Check-AddinRegistryEntries }
    "4" { Check-FileSystemCache }
    "5" { Test-OutlookCOMAccess }
    "6" { Test-NetworkConnectivity }
    "7" { Test-ManifestFiles }
    "8" { Enable-OfficeAddinDebugging }
    "9" { Disable-OfficeAddinDebugging }
    "10" { Show-OfficeAddinLogs }
    "11" { Start-AddinLoadMonitoring }
    "12" { Clean-DebugArtifacts }
    "13" {
        Write-Status "🔍 Running Comprehensive Analysis..." "Cyan"
        Write-Status ""
        Get-OfficeEnvironmentInfo
        Check-CriticalOfficeSettings  
        Check-AddinRegistryEntries
        Check-FileSystemCache
        Test-OutlookCOMAccess
        Test-NetworkConnectivity
        Test-ManifestFiles
        
        Write-Status ""
        Write-Status "✅ Comprehensive analysis complete!" "Green"
        Write-Status ""
        Write-Status "💡 Quick Troubleshooting Recommendations:" "Yellow"
        Write-Status "1. If no add-in button: Check 'Optional Connected Experiences' setting first!" "White"
        Write-Status "2. If registry empty: Try sideloading the manifest directly in Outlook" "White"
        Write-Status "3. If network fails: Check firewall/proxy settings" "White"
        Write-Status "4. If COM errors: Try running PowerShell as Administrator" "White"
        Write-Status "5. Use option 8 + 11 to enable debugging and monitor real-time" "White"
    }
    default { Write-Status "Invalid choice. Please enter a number between 1-13." "Red" }
}

Write-Status ""
