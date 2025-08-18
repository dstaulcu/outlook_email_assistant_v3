# Office Add-in Troubleshooting Checklist

## ❗ CRITICAL FIRST STEP: Optional Connected Experiences

**This is the #1 cause of add-in buttons not appearing!**

### Steps to Check:
1. Open Outlook
2. Go to **File → Options → Trust Center → Trust Center Settings → Privacy Options**
3. Verify **"Optional connected experiences"** is **CHECKED/ENABLED**
4. If disabled, enable it and **restart Outlook completely**

### Why This Matters:
- When disabled, ALL web-based add-ins fail to load
- No error messages are shown
- Add-in appears to be installed but button never appears
- Affects both development and production add-ins

---

## Secondary Checks (if add-in still doesn't appear)

### 1. Add-in Trust Settings
**Location**: File → Options → Trust Center → Trust Center Settings → Add-ins

**Required Settings**:
- ✅ **UNCHECK** "Require Application Add-ins to be signed by Trusted Publisher" 
- ✅ **UNCHECK** "Disable all Application Add-ins"

### 2. Clear Office Add-in Cache
```powershell
# Run this command in PowerShell
Remove-Item -Recurse -Force "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
```
Then restart Outlook.

### 3. Verify Office.js Dependencies
- Office.js modules are hosted locally in S3 bucket (air-gapped network design)
- Test access to your configured S3 endpoint: Check `tools\deployment-environments.json`
- Verify corporate firewall allows S3 bucket access
- No external internet connectivity required for Office.js dependencies

### 4. Office Version Compatibility
- Minimum required: Office 2016 or Microsoft 365
- Web add-ins require modern Office versions
- Check Office update status

---

## Advanced Debugging

### Use the Debug Toolkit
```powershell
.\tools\outlook_addin_diagnostics.ps1
```
Select option **6** to check all critical Office settings automatically.

### Manual Registry Check
Check if add-in is registered:
```
HKCU\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs
```

### Windows Event Logs
Look for Office add-in errors:
1. Open Event Viewer
2. Navigate to: Applications and Services Logs → Microsoft → Office → Outlook
3. Look for add-in loading errors

---

## Quick Test Process

1. ✅ **Enable "Optional Connected Experiences"** (most important!)
2. ✅ Clear Office cache and restart Outlook
3. ✅ Verify add-in trust settings
4. ✅ Test with a known-good add-in from Office Store
5. ✅ Run debug toolkit if issues persist

---

## Still Not Working?

If the add-in still doesn't appear after all checks:

1. **Test in different Office environment**:
   - Try Outlook Web Access (OWA) in browser
   - Test on different computer
   - Try different Office profile

2. **Check manifest file**:
   - Validate XML syntax
   - Verify all URLs are accessible
   - Ensure HTTPS is used (except localhost)

3. **Contact Support**:
   - Include Office version details
   - Provide screenshots of Trust Center settings
   - Share any error messages from Event Viewer

**Remember**: "Optional Connected Experiences" causes 90% of add-in loading issues!
