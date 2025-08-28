# Tools Directory Cleanup & Organization

## 📁 Core Production Tools (Keep These)
These are the essential tools for normal deployment and operation:

### **Primary Installer**
- `outlook_installer.ps1` - Main installer with Process Monitor-based registry sideloading
- `deployment-environments.json` - Environment configuration
- `deploy_web_assets.ps1` - S3 deployment script

### **Diagnostics & Troubleshooting**  
- `outlook_troubleshooter.ps1` - Systematic diagnostics
- `outlook_addin_diagnostics.ps1` - Basic diagnostics
- `outlook_cache_clear.ps1` - Cache clearing utility

### **Manual Installation Fallbacks**
- `outlook_manual_install.ps1` - Manual installation guide
- `install-outlook-assistant.bat` - Simple batch installer

### **Configuration**
- `set-environment-dev.reg` - Registry files for environment setting
- `set-environment-test.reg`
- `set-environment-prod.reg`

## 🧪 Development/Research Tools (Can Remove)
These were created during our troubleshooting research:

### **Research Tools (DELETE THESE)**
- `outlook_cloud_detective.ps1` - Cloud source investigation (one-time use)
- `outlook_nuclear_uninstall.ps1` - Extreme cleanup (Microsoft 365 specific issue)
- `outlook_enterprise_uninstall.ps1` - Enterprise-focused uninstall (redundant with main installer)
- `outlook_admin_cleanup.ps1` - Auto-generated helper (not needed)

### **Legacy Tools (DEPRECATED)**
- `outlook_addin_sideload.ps1` - Old sideloading method (replaced by Process Monitor method)

## 📋 Documentation (Keep & Update)
- `README.md` - Main tools documentation
- `INSTALLER_README.md` - Installation guide

## 🎯 Recommended Cleanup Actions

### Files to DELETE:
```powershell
Remove-Item "outlook_cloud_detective.ps1"
Remove-Item "outlook_nuclear_uninstall.ps1" 
Remove-Item "outlook_enterprise_uninstall.ps1"
Remove-Item "outlook_admin_cleanup.ps1"
Remove-Item "outlook_addin_sideload.ps1"
```

### Files to KEEP:
- All core production tools
- All documentation
- All configuration files

## 📝 Documentation Updates Needed

### Update `README.md` with:
1. **Process Monitor Discovery**: Document the OutlookSideloadManifestPath registry key discovery
2. **Enterprise Deployment**: Add section about Office 365 vs on-premises differences  
3. **Troubleshooting**: Add cloud sync interference notes
4. **Installation Options**: Document the three installation methods (registry, manual, admin)

### Update `INSTALLER_README.md` with:
1. **Known Issues**: Microsoft 365 personal account sync behavior
2. **Environment Differences**: Home vs work deployment considerations
3. **Uninstall Process**: Comprehensive uninstall steps including cloud considerations

## 🚀 Final Tool Organization

After cleanup, you'll have a clean, professional toolset:

```
tools/
├── outlook_installer.ps1           # Main installer (Process Monitor method)
├── outlook_troubleshooter.ps1      # Comprehensive diagnostics  
├── outlook_manual_install.ps1      # Manual installation fallback
├── outlook_cache_clear.ps1         # Utility script
├── deploy_web_assets.ps1          # S3 deployment
├── deployment-environments.json    # Configuration
├── install-outlook-assistant.bat   # Simple batch installer
├── set-environment-*.reg          # Registry configuration files
├── README.md                      # Main documentation
└── INSTALLER_README.md            # Installation guide
```

This gives you a professional, maintainable toolkit ready for enterprise deployment! 🎯
