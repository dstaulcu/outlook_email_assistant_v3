# Deployment Guide

This guide covers the complete deployment process for the PromptEmail Outlook Add-in, including environment setup, automated deployment, and production considerations.

## Prerequisites

### Required Tools
1. **AWS CLI**: Configured with appropriate S3
2. **PowerShell**: Version 5.1+ or PowerShell Core 7+
3. **Node.js**: Version 16+ with npm
4. **Git**: For version control and deployment tracking

### AWS Permissions Required
Your AWS user/role needs the following permissions:
- `s3:CreateBucket`, `s3:ListBucket`, `s3:GetObject`, `s3:PutObject`, `s3:DeleteObject`
- `s3:PutBucketWebsite`, `s3:PutBucketPolicy` 

## Environment Configuration

### 1. Deployment Environments Setup

The project uses `tools\deployment-environments.json` for environment management:

```json
{
  "Dev": {
    "bucketName": "your-company-promptemail-dev",
    "region": "us-east-1", 
    "description": "Development environment for testing"
  },
  "Prd": {
    "bucketName": "your-company-promptemail-prod",
    "region": "us-east-1",
    "description": "Production environment"
  }
}
```

### 2. AWS S3 Bucket Creation and Configuration

#### Create S3 Buckets
```bash
# Create development bucket
aws s3 mb s3://your-company-promptemail-dev --region us-east-1

# Create production bucket  
aws s3 mb s3://your-company-promptemail-prod --region us-east-1
```

#### Configure Static Website Hosting
```bash
# Enable static website hosting
aws s3 website s3://your-company-promptemail-dev --index-document index.html --error-document index.html

aws s3 website s3://your-company-promptemail-prod --index-document index.html --error-document index.html
```

#### Set Bucket Policy for Public Read Access

Create `bucket-policy.json`:
```json
{
    "Version": "2012-10-17",
    "Statement": [
        {
            "Sid": "PublicReadGetObject",
            "Effect": "Allow",
            "Principal": "*",
            "Action": "s3:GetObject",
            "Resource": "arn:aws:s3:::your-company-promptemail-prod/*"
        }
    ]
}
```

Apply the policy:
```bash
aws s3api put-bucket-policy --bucket your-company-promptemail-prod --policy file://bucket-policy.json
```

### 3. Manifest Configuration

Update `src/manifest.xml` with your deployment URLs:

```xml
<!-- Update the GUID to be unique for your organization -->
<Id>12345678-1234-1234-1234-123456789012</Id>

<!-- Production URLs -->
<IconUrl DefaultValue="https://your-company-promptemail-prod.s3-website-us-east-1.amazonaws.com/icons/icon-32.png"/>
<HighResolutionIconUrl DefaultValue="https://your-company-promptemail-prod.s3-website-us-east-1.amazonaws.com/icons/icon-128.png"/>

<!-- App Domain -->
<AppDomain>https://your-company-promptemail-prod.s3-website-us-east-1.amazonaws.com</AppDomain>
```

## Automated Deployment Process

### 1. Primary Deployment Script

The main deployment script `tools\deploy_web_assets.ps1` provides comprehensive build and deployment automation:

```powershell
# Deploy to development environment
.\tools\deploy_web_assets.ps1 -Environment Dev

# Deploy to production with safety checks
.\tools\deploy_web_assets.ps1 -Environment Prd

# Preview deployment without making changes  
.\tools\deploy_web_assets.ps1 -Environment Prd -DryRun

# Force deployment (skip validation prompts)
.\tools\deploy_web_assets.ps1 -Environment Prd -Force
```

### 2. Deployment Script Features

#### Automated Build Process
- Runs `npm run build` to create production assets
- Validates all required files are present
- Performs manifest validation
- Optimizes assets for production

#### Smart URL Management  
- Automatically updates embedded URLs in all files
- Replaces localhost references with production URLs
- Handles both absolute and relative path references
- Preserves URL schemes (http/https/s3://)

#### File Synchronization
- Uploads only changed files for faster deployments
- Maintains file structure and permissions
- Handles icon assets, HTML, CSS, and JavaScript bundles
- Preserves source maps for debugging

#### Safety Features
- Dry-run mode for previewing changes
- Environment validation to prevent wrong deployments
- Backup recommendations before major updates
- Rollback procedures documentation

### 3. Legacy Deployment Support

The `tools\deploy.ps1` script provides basic deployment functionality:

```powershell
# Basic deployment (legacy)
.\tools\deploy.ps1 -BucketName your-bucket-name -Region us-east-1

# Preview mode
.\tools\deploy.ps1 -BucketName your-bucket-name -DryRun
```

## Build Process Details

### 1. Webpack Build Configuration

The build process uses Webpack 5 with the following optimizations:

#### Production Optimizations
- **Minification**: JavaScript and CSS minification  
- **Bundle splitting**: Separate bundles for taskpane and commands
- **Asset optimization**: Icon compression and optimization
- **Source maps**: Generated for production debugging

#### Build Outputs
```
public/
├── index.html              # Commands entry point
├── taskpane.html          # Main application interface  
├── taskpane.bundle.js     # Main application logic (minified)
├── commands.bundle.js     # Ribbon command handlers
├── taskpane.css          # Styles 
├── default-providers.json # AI provider configurations
├── default-models.json   # Default model mappings
├── icons/                # Optimized icon assets
│   ├── icon-16.png
│   ├── icon-32.png  
│   ├── icon-80.png
│   └── icon-128.png
└── *.js.map              # Source maps for debugging
```

### 2. Asset Processing Pipeline

#### CSS Processing
- CSS bundling and minification
- Custom property support
- Cross-browser compatibility

#### Icon Processing  
- PNG optimization
- Multiple resolution support (16px, 32px, 80px, 128px)
- Proper aspect ratio validation

#### JavaScript Processing
- ES6 module bundling
- Class property support
- Async/await transformation for older browsers

## Production Deployment Workflow

### 1. Pre-Deployment Checklist

#### Code Quality Verification
- [ ] All code changes committed and pushed to main branch
- [ ] No console.log or debugging code in production
- [ ] All TODO comments resolved or documented
- [ ] Version numbers updated in `package.json` and manifest

#### Build Verification  
- [ ] Clean build completes without errors: `npm run build`
- [ ] Manifest validation passes: `npm run validate-manifest`  
- [ ] Bundle size is acceptable (check webpack-bundle-analyzer)
- [ ] All required assets present in `public/` directory

#### Environment Configuration
- [ ] Production environment configured in `deployment-environments.json`
- [ ] S3 bucket exists and has proper permissions
- [ ] AWS CLI configured with correct credentials
- [ ] Manifest URLs point to production environment

#### Security and Privacy
- [ ] No API keys or secrets in client-side code
- [ ] Sensitive data sanitization verified in Logger service
- [ ] Classification detection patterns tested
- [ ] HTTPS URLs used for all production resources

### 2. Deployment Execution

#### Standard Production Deployment
```powershell
# 1. Perform dry run to preview changes
.\tools\deploy_web_assets.ps1 -Environment Prd -DryRun

# 2. Review the proposed changes carefully

# 3. Execute actual deployment
.\tools\deploy_web_assets.ps1 -Environment Prd

# 4. Verify deployment success
# Check AWS S3 console for uploaded files
# Test website URL accessibility
```

#### Emergency/Hotfix Deployment
```powershell
# For critical fixes, use force mode to skip prompts
.\tools\deploy_web_assets.ps1 -Environment Prd -Force
```

### 3. Post-Deployment Verification

#### Functional Testing
1. **Access verification**: Open S3 website URL in browser
2. **Manifest testing**: Validate manifest XML at deployment URL
3. **Asset loading**: Verify all icons, CSS, and JS files load correctly
4. **Cross-browser testing**: Test in Edge, Chrome, Firefox

#### Outlook Integration Testing
1. **Sideload testing**: Install manifest in Outlook Desktop
2. **Ribbon integration**: Verify button appears in Message Read/Compose
3. **Taskpane functionality**: Test opening and basic operations
4. **AI service integration**: Verify API connections work
5. **Settings persistence**: Test settings save/load functionality

#### Performance Validation
1. **Load times**: Measure initial load and subsequent operations
2. **Memory usage**: Monitor memory consumption during use
3. **API response times**: Test AI provider response times
4. **Error handling**: Verify graceful degradation on failures

## Sideloading and Distribution

### 1. Manual Sideloading (Development/Testing)

#### Outlook Desktop Sideloading
1. Open Outlook Desktop
2. Go to **File** > **Manage Add-ins** > **My Add-ins**  
3. Click **Add a custom add-in** > **Add from file**
4. Select your `manifest.xml` file (ensure it points to production URLs)
5. Click **Install**

#### Outlook Web App Sideloading  
1. Open Outlook on the web (office.com)
2. Click **Settings** gear > **View all Outlook settings**
3. Navigate to **General** > **Manage add-ins** 
4. Click **Add a custom add-in** > **Add from file**
5. Upload your `manifest.xml` file

### 2. Enterprise Distribution

#### Microsoft 365 Admin Center Deployment
For organization-wide deployment:

1. **Admin Center Access**: Navigate to admin.microsoft.com
2. **Add-in Management**: Go to **Settings** > **Integrated apps**
3. **Upload Custom App**: Click **Upload custom apps**
4. **Manifest Upload**: Upload your production manifest.xml
5. **User Assignment**: Assign to specific users or groups
6. **Deployment Monitoring**: Monitor installation status

#### AppSource Publication (Future)
For public distribution:
1. Prepare for AppSource submission requirements
2. Complete certification process
3. Submit through Partner Center
4. Await Microsoft validation and approval

## Testing and Validation

### 1. Automated Testing Pipeline

#### Manifest Validation
```bash
# Validate Office Add-in manifest schema
npm run validate-manifest

# Custom validation for production URLs
node -e "
const fs = require('fs');
const manifest = fs.readFileSync('manifest.xml', 'utf8');
const hasLocalhost = manifest.includes('localhost');
if (hasLocalhost) {
  console.error('ERROR: Manifest contains localhost URLs');
  process.exit(1);
}
console.log('✓ Manifest validation passed');
"
```

#### Build Validation  
```powershell
# Complete build and validation pipeline
$ErrorActionPreference = 'Stop'

Write-Host "Running build validation pipeline..." -ForegroundColor Blue

# Clean build
npm run build

# Validate all required files exist  
$requiredFiles = @(
    'public/index.html',
    'public/taskpane.html', 
    'public/taskpane.bundle.js',
    'public/commands.bundle.js',
    'public/taskpane.css'
)

foreach ($file in $requiredFiles) {
    if (!(Test-Path $file)) {
        throw "Required file missing: $file"
    }
}

Write-Host "✓ All required build files present" -ForegroundColor Green
```

### 2. Functional Testing Procedures

#### Email Analysis Testing
Create test emails with various content types:

1. **Plain text emails**: Basic functionality testing
2. **HTML-rich emails**: Complex formatting handling  
3. **Classified content**: Test classification detection
4. **Multiple recipients**: Test recipient parsing
5. **Attachments**: Test attachment detection
6. **Various languages**: Internationalization testing

#### AI Provider Testing
Test each supported AI provider:

```javascript
// Test script for AI provider validation
const testProviders = ['openai', 'ollama', 'custom'];
const testMessage = "Analyze this test email content";

testProviders.forEach(async (provider) => {
  try {
    const config = {
      provider,
      model: 'default-model',
      apiKey: 'test-key'
    };
    
    console.log(`Testing ${provider}...`);
    // Note: This would require actual API keys for real testing
    
  } catch (error) {
    console.error(`${provider} test failed:`, error.message);
  }
});
```

### 3. Accessibility and Compliance Testing

#### WCAG 2.1 AA Compliance Checklist
- [ ] Keyboard navigation for all interactive elements
- [ ] Screen reader compatibility (test with NVDA/JAWS)
- [ ] Color contrast ratios meet AA standards (4.5:1 normal text)
- [ ] ARIA labels and live regions properly implemented
- [ ] Focus management and tab order logical
- [ ] Error messages announced to assistive technology

#### Security Classification Testing
- [ ] Test detection of UNCLASSIFIED markings
- [ ] Test detection of CONFIDENTIAL/SECRET markings  
- [ ] Test user override warnings and logging
- [ ] Verify sensitive data sanitization in logs
- [ ] Test audit trail completeness

## Troubleshooting and Maintenance

### 1. Common Deployment Issues

#### Build Failures
```powershell
# Common build issue resolution
# 1. Clear node_modules and package-lock
Remove-Item -Recurse -Force node_modules, package-lock.json -ErrorAction SilentlyContinue
npm install

# 2. Clear webpack cache
Remove-Item -Recurse -Force public -ErrorAction SilentlyContinue  

# 3. Rebuild
npm run build
```

#### S3 Upload Failures
```bash
# Debug S3 permissions
aws s3api head-bucket --bucket your-bucket-name
aws s3api get-bucket-location --bucket your-bucket-name

# Test file upload manually
aws s3 cp public/index.html s3://your-bucket-name/ --dry-run
```

#### Office Add-in Loading Issues

**❗ CRITICAL: Check "Optional Connected Experiences" First**

The most common cause of add-in buttons not appearing is disabled "Optional Connected Experiences":

1. **Location**: File → Options → Trust Center → Trust Center Settings → Privacy Options
2. **Setting**: "Optional connected experiences" must be **ENABLED** 
3. **Impact**: When disabled, ALL web-based add-ins fail to load with no error messages
4. **Solution**: Enable the setting and restart Outlook

**Other Common Add-in Issues:**

1. **Clear Office cache**:
   ```powershell
   Remove-Item -Recurse -Force "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
   ```

2. **Reset Office settings**:
   ```powershell
   # Reset Office add-in trust center settings
   reg delete "HKCU\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs" /f
   ```

3. **Verify add-in trust settings**:
   - File → Options → Trust Center → Trust Center Settings → Add-ins
   - Uncheck "Require Application Add-ins to be signed by Trusted Publisher"
   - Uncheck "Disable all Application Add-ins"

4. **Use comprehensive debugging tool**:
   ```powershell
   .\tools\outlook_addin_diagnostics.ps1
   # Select option 6 to check all critical Office settings
   ```

### 2. Performance Monitoring

#### Bundle Size Monitoring
```bash
# Analyze bundle size trends
npm run build -- --json > bundle-stats.json
npx webpack-bundle-analyzer bundle-stats.json public/
```

#### Runtime Performance
```javascript
// Performance monitoring in production
class PerformanceMonitor {
  static measureOperation(name, operation) {
    const start = performance.now();
    const result = await operation();
    const duration = performance.now() - start;
    
    // Log to application telemetry
    console.log(`${name}: ${duration.toFixed(2)}ms`);
    return result;
  }
}

// Usage example
const analysisResult = await PerformanceMonitor.measureOperation(
  'Email Analysis',
  () => aiService.analyzeEmail(emailData, config)
);
```

### 3. Rollback and Recovery Procedures

#### Emergency Rollback
```powershell
# Quick rollback to previous version
# 1. Identify previous deployment tag/commit
git tag --list --sort=-version:refname | Select-Object -First 5

# 2. Checkout previous version
git checkout v1.2.3  # Replace with actual version tag

# 3. Rebuild and redeploy
npm run build
.\tools\deploy_web_assets.ps1 -Environment Prd -Force

# 4. Verify rollback success
# Test key functionality in Outlook
```

#### Data Recovery
- Settings are stored in Office.js roaming settings (backed up by Microsoft)
- Local fallback in browser localStorage
- No server-side data to recover
- User-specific API keys may need to be re-entered

### 4. Monitoring and Observability

#### Windows Event Log Monitoring
```powershell
# Monitor PromptEmail application events
Get-WinEvent -FilterHashtable @{LogName='Application'; ProviderName='PromptEmail'} -MaxEvents 50

# Set up event forwarding for centralized logging
wevtutil sl Application /rt:false
```

#### AWS CloudWatch Integration (Optional)
For advanced monitoring, consider S3 access logging:
```json
{
  "Rules": [{
    "ID": "AccessLogRule",
    "Status": "Enabled", 
    "Filter": {"Prefix": ""},
    "Destination": {
      "BucketName": "your-logs-bucket",
      "Prefix": "access-logs/"
    }
  }]
}
```

## Security and Compliance

### 1. Production Security Checklist

#### Data Protection
- [ ] No sensitive data (API keys, emails) logged or transmitted
- [ ] All external API calls use HTTPS
- [ ] Sensitive fields sanitized in SettingsManager and Logger
- [ ] Classification detection warns before processing classified content
- [ ] User consent required for classified content processing

#### API Security
- [ ] API keys stored only in Office.js roaming settings (encrypted by Microsoft)
- [ ] No API keys in client-side code or build artifacts
- [ ] API endpoints validate against allowed origins
- [ ] Rate limiting implemented where applicable

#### Infrastructure Security  
- [ ] S3 bucket configured with minimal required permissions
- [ ] AWS IAM roles follow principle of least privilege
- [ ] Regular security credential rotation

### 2. Compliance Considerations

#### Email Classification Handling
The add-in includes built-in detection for:
- UNCLASSIFIED
- CONFIDENTIAL  
- SECRET
- TOP SECRET
- COSMIC TOP SECRET

**Warning System**: Users receive explicit warnings before AI processing of classified content, with audit logging of override decisions.

#### Audit and Logging
- All user actions logged to Windows Application Log
- Sensitive content filtered from logs
- User consent and override decisions recorded
- Session tracking for compliance reporting

#### Data Retention
- No persistent storage of email content
- Settings stored in Microsoft-managed Office.js roaming settings
- Temporary processing data cleared after each operation
- No server-side data retention

### 3. Privacy Protection

#### Minimal Data Collection
- Only essential metadata collected (timestamp, user ID, session info)
- Email content processed locally and not stored
- AI API calls contain only necessary content for analysis
- No personal identifiable information in telemetry

#### User Control
- Users control AI provider selection and API keys
- Opt-out available for all logging and telemetry
- Settings export/import for user data portability
- Clear data deletion procedures

## Production Checklist

### Pre-Release Validation
- [ ] Code review completed by senior developer
- [ ] Security review passed
- [ ] Accessibility testing (WCAG 2.1 AA) completed
- [ ] Cross-browser compatibility verified
- [ ] Performance benchmarks met
- [ ] Documentation updated

### Build and Deployment
- [ ] Clean build without warnings: `npm run build`
- [ ] Manifest validation passed: `npm run validate-manifest`
- [ ] All assets uploaded to production S3 bucket
- [ ] URLs in manifest point to correct production paths
- [ ] SSL certificates valid and not expiring soon

### Functional Verification
- [ ] Ribbon button appears in Outlook Message Read/Compose
- [ ] Taskpane loads without errors
- [ ] All AI providers connect successfully (with valid API keys)
- [ ] Email analysis functionality works end-to-end
- [ ] Response generation and insertion works
- [ ] Settings persistence works across sessions
- [ ] Classification detection triggers appropriate warnings

### Accessibility and Compliance
- [ ] Keyboard navigation works throughout application
- [ ] Screen reader announcements work correctly
- [ ] High contrast mode displays properly
- [ ] Classification warnings display and log correctly
- [ ] Sensitive data sanitization verified
- [ ] Windows Application Log events generating properly

### Performance and Reliability  
- [ ] Initial load time under 3 seconds
- [ ] AI operations complete within reasonable time (< 30 seconds)
- [ ] Memory usage remains stable during extended use
- [ ] Error handling provides meaningful user feedback
- [ ] Graceful degradation when AI services unavailable

### Post-Deployment Monitoring
- [ ] Windows Event Log monitoring configured
- [ ] S3 access logs enabled (if required)
- [ ] User feedback collection mechanism in place
- [ ] Performance metrics baseline established
- [ ] Support documentation distributed to end users

## Maintenance and Updates

### Regular Maintenance Schedule

#### Monthly Reviews
- Review Windows Application Log for errors or issues
- Check S3 storage usage and costs
- Verify SSL certificate expiration dates
- Update dependencies with security patches

#### Quarterly Updates
- Review and update AI provider configurations
- Performance optimization and bundle size analysis
- Accessibility compliance re-verification
- Security vulnerability assessment

#### Annual Reviews
- Comprehensive security audit
- Compliance documentation update
- User feedback analysis and feature planning
- Infrastructure cost optimization review

### Update Deployment Process
1. **Feature Branches**: All updates developed in feature branches
2. **Testing**: Comprehensive testing on development environment  
3. **Staging**: Deploy to staging environment for user acceptance testing
4. **Production**: Controlled production deployment with rollback plan
5. **Monitoring**: Enhanced monitoring during and after deployment

### Emergency Response Procedures
1. **Incident Detection**: Monitoring alerts or user reports
2. **Assessment**: Rapid assessment of impact and severity  
3. **Response**: 
   - Critical: Immediate rollback to previous version
   - High: Hotfix deployment within 24 hours
   - Medium: Include in next scheduled release
4. **Communication**: User notification via appropriate channels
5. **Post-Incident**: Root cause analysis and prevention measures
