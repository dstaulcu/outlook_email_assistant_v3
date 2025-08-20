# Tools Directory

This directory contains various PowerShell scripts for building, deploying, and debugging the PromptEmail Outlook Add-in.

## 📁 Script Organization

### 🚀 Build & Deployment
- **`deploy_web_assets.ps1`** - Complete build and deployment pipeline for web assets (recommended for production)
- **`deploy_s3_config.ps1`** - S3 bucket creation and configuration only

### 🛠️ Development & Debugging  
- **`outlook_addin_diagnostics.ps1`** - Comprehensive Office add-in diagnostics and debugging toolkit
- **`outlook_cache_clear.ps1`** - Clear Office/Outlook add-in cache
- **`outlook_addin_sideload.ps1`** - Production-focused sideloading helper for S3-hosted manifests

### ⚙️ Configuration
- **`deployment-environments.json`** - Environment-specific configuration (buckets, regions, URLs)

## 🎯 Recommended Workflows

### Initial Setup (One-time)
```powershell
# Create and configure S3 buckets for all environments
.\deploy_s3_config.ps1 -AllEnvironments

# Or create specific environment
.\deploy_s3_config.ps1 -Environment Dev
```

### Development Deployment
```powershell
# Build and deploy to development environment
.\deploy_web_assets.ps1 -Environment Dev

# Dry run to see what would be deployed
.\deploy_web_assets.ps1 -Environment Dev -DryRun
```

### Production Deployment  
```powershell
# Build and deploy to production environment
.\deploy_web_assets.ps1 -Environment Prd
```

### Debugging Add-in Issues
```powershell
# Comprehensive diagnostics toolkit
.\outlook_addin_diagnostics.ps1

# Quick cache clear
.\outlook_cache_clear.ps1
```

### Production Sideloading
```powershell
# Sideload production manifest from S3
.\outlook_addin_sideload.ps1 -ManifestUrl "https://your-bucket.s3.region.amazonaws.com/manifest.xml"

# Sideload development manifest
.\outlook_addin_sideload.ps1 -ManifestUrl "https://dev-bucket.s3.region.amazonaws.com/manifest.xml" -Environment Dev

# Use local manifest file
.\outlook_addin_sideload.ps1 -ManifestUrl ".\public\manifest.xml" -UseLocalManifest
```

## 🔧 Script Dependencies

- **AWS CLI** - Required for S3 deployment operations
- **Node.js & npm** - Required for build operations  
- **PowerShell 5.1+** - All scripts require PowerShell
- **Office/Outlook** - Required for add-in debugging scripts

## 📋 Configuration Files

### deployment-environments.json
Defines environment-specific settings:
- S3 bucket names and regions
- Public URL endpoints
- Regional configurations

Example structure:
```json
{
  "environments": {
    "Dev": {
      "region": "us-east-1",
      "s3Uri": { "host": "bucket-name-dev" },
      "publicUri": { "host": "bucket-name-dev.s3.region.amazonaws.com" }
    }
  }
}
```

## 🚨 Troubleshooting

If add-in buttons don't appear, check in this order:

1. **Enable "Optional Connected Experiences"** in Outlook settings (most common issue)
2. Run `.\outlook_addin_diagnostics.ps1` and select option 2 for critical Office settings check
3. Clear cache with `.\outlook_cache_clear.ps1`
4. Verify S3 bucket accessibility
