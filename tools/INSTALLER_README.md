# Outlook Email Assistant - Windows Installer

This directory contains installer scripts for deploying the Outlook Email Assistant add-in on Windows systems.

## Files

- **`outlook_installer.ps1`** - Main PowerShell installer script
- **`install-outlook-assistant.bat`** - Simple batch file wrapper
- **`outlook_cache_clear.ps1`** - Original cache clearing script (for reference)

## Quick Installation

### Option 1: Batch File (Easiest)
```cmd
# Install production version
install-outlook-assistant.bat

# Install development version
install-outlook-assistant.bat --dev

# Silent installation (no prompts)
install-outlook-assistant.bat --prod --silent

# Uninstall
install-outlook-assistant.bat --uninstall
```

### Option 2: PowerShell Script (More Options)
```powershell
# Install production version
.\outlook_installer.ps1 -Environment Prd

# Install with custom manifest URL
.\outlook_installer.ps1 -ManifestUrl "https://custom-url.com/manifest.xml"

# Silent installation to custom path
.\outlook_installer.ps1 -Environment Prd -InstallPath "C:\MyApps\OutlookAssistant" -Silent

# Uninstall only
.\outlook_installer.ps1 -UninstallOnly
```

## What the Installer Does

1. **Downloads Manifest**: Retrieves the manifest.xml file from the configured S3 bucket
2. **Stops Outlook**: Safely terminates all Outlook processes
3. **Clears Cache**: Removes Office add-in cache files that might cause conflicts
4. **Registry Configuration**: Adds the necessary registry keys for sideloading the add-in
5. **Verification**: Confirms the installation was successful

## Installation Process

### Automatic Steps:
1. Downloads manifest from S3 (based on environment or custom URL)
2. Validates the manifest file format
3. Stops any running Outlook processes
4. Clears Office add-in cache directories:
   - `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
   - `%LOCALAPPDATA%\Microsoft\Office\15.0\Wef` 
   - `%LOCALAPPDATA%\Microsoft\Office\14.0\Wef`
5. Adds registry entries for sideloading:
   - `HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\OutlookEmailAssistant`
   - `HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer\OutlookEmailAssistant`
6. Verifies installation success

### Manual Steps (After Installation):
1. Start Outlook
2. Look for the add-in buttons in the ribbon
3. If buttons don't appear, check Office settings:
   - File → Options → Trust Center → Trust Center Settings → Add-ins
   - Ensure "Require Application Add-ins to be signed by Trusted Publisher" is UNCHECKED
   - Ensure "Optional connected experiences" is ENABLED (most common issue)

## Environment Configuration

The installer uses the `deployment-environments.json` file to determine S3 URLs:

```json
{
  "environments": {
    "Dev": {
      "region": "us-east-1",
      "publicUri": {
        "protocol": "https",
        "host": "your-bucket-dev.s3.region.amazonaws.com"
      }
    },
    "Prd": {
      "region": "us-east-1", 
      "publicUri": {
        "protocol": "https",
        "host": "your-bucket-prod.s3.region.amazonaws.com"
      }
    }
  }
}
```

## Parameters Reference

### PowerShell Script Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `Environment` | String | "Prd" | Environment to install from (Dev, Prd) |
| `ManifestUrl` | String | "" | Custom manifest URL (overrides environment) |
| `InstallPath` | String | `%LOCALAPPDATA%\OutlookEmailAssistant` | Installation directory |
| `Silent` | Switch | false | Run without user prompts |
| `UninstallOnly` | Switch | false | Only remove the add-in |
| `Help` | Switch | false | Show help message |

### Batch File Options

| Option | Description |
|--------|-------------|
| `--dev` | Install from development environment |
| `--prod` | Install from production environment (default) |
| `--silent` | Install without user prompts |
| `--uninstall` | Uninstall the add-in |
| `--help`, `-h`, `/?` | Show help message |

## Use Cases

### End User Installation
```cmd
install-outlook-assistant.bat --prod --silent
```

### Development/Testing
```powershell
.\outlook_installer.ps1 -Environment Dev
```

### Custom Deployment
```powershell
.\outlook_installer.ps1 -ManifestUrl "https://custom.company.com/manifest.xml" -InstallPath "C:\CompanyApps\EmailAssistant"
```

### MSI Package Integration
The PowerShell script can be embedded in MSI packages or other installers:
```cmd
powershell.exe -ExecutionPolicy Bypass -File "outlook_installer.ps1" -Environment Prd -Silent
```

### Group Policy Deployment
Deploy via logon script:
```cmd
if not exist "%LOCALAPPDATA%\OutlookEmailAssistant\manifest.xml" (
    "\\server\share\install-outlook-assistant.bat" --prod --silent
)
```

## Troubleshooting

### Common Issues:

1. **"Execution of scripts is disabled"**
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

2. **"Access denied" or permission errors**
   - Run as Administrator
   - Close all Office applications before installing

3. **Add-in doesn't appear in Outlook**
   - Check Office settings: File → Options → Trust Center → Trust Center Settings → Privacy Options
   - Ensure "Optional connected experiences" is ENABLED
   - Restart Outlook completely

4. **Download fails**
   - Check internet connectivity
   - Verify S3 bucket is publicly accessible
   - Try custom manifest URL to test

5. **Registry keys not added**
   - Run as Administrator
   - Check if antivirus software is blocking registry modifications

### Debug Mode:
Run with verbose output:
```powershell
.\outlook_installer.ps1 -Environment Prd -Verbose
```

### Manual Verification:
Check if registry keys exist:
```powershell
Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer" -Name "OutlookEmailAssistant" -ErrorAction SilentlyContinue
```

### Manual Cleanup:
```powershell
.\outlook_installer.ps1 -UninstallOnly
```

## Security Considerations

- Scripts download from HTTPS URLs only
- Manifest files are validated before installation
- Registry modifications are limited to current user (HKCU)
- No elevation required for basic installation
- Installation path defaults to user's local app data

## MSI Integration Example

For enterprise deployment, the PowerShell script can be called from MSI custom actions:

```xml
<CustomAction Id="InstallOutlookAddin" 
              BinaryKey="WixCA" 
              DllEntry="WixQuietExec" 
              Execute="deferred" 
              Return="check" 
              Impersonate="yes" />

<SetProperty Id="InstallOutlookAddin" 
             Value="&quot;[System64Folder]WindowsPowerShell\v1.0\powershell.exe&quot; -ExecutionPolicy Bypass -File &quot;[INSTALLFOLDER]outlook_installer.ps1&quot; -Environment Prd -Silent" 
             Before="InstallOutlookAddin" 
             Sequence="InstallExecuteSequence" />
```

This provides a comprehensive Windows installer solution that can be used standalone or integrated into enterprise deployment systems.
