# Development Setup Guide

This guide will help you set up the PromptEmail Outlook Add-in for development and testing.

## Prerequisites

### Required Software
- **Node.js 16+**: [Download from nodejs.org](https://nodejs.org/)
- **npm**: Comes with Node.js
- **Outlook Desktop**: Microsoft 365 subscription
- **AWS CLI**: [Download from AWS](https://aws.amazon.com/cli/) (for deployment)
- **PowerShell**: Windows PowerShell 5.1+ or PowerShell Core 7+
- **Git**: For version control

### Optional Tools
- **Visual Studio Code**: Recommended editor with Office Add-in extensions
- **Office Add-in Debugger**: VS Code extension for debugging

## Initial Setup

### 1. Clone and Install Dependencies

```bash
# Clone the repository
git clone <your-repo-url>
cd outlook_email_assistant_v3

# Install dependencies
npm install
```

### 2. Create Real Icon Files

The project includes placeholder icon files. You need to create actual PNG files:

```
src/assets/icons/
├── icon-16.png   (16x16 pixels)
├── icon-32.png   (32x32 pixels) 
├── icon-80.png   (80x80 pixels)
└── icon-128.png  (128x128 pixels)
```

### 3. Configure AWS S3 (for deployment)

1. Create an S3 bucket for hosting:
```bash
aws s3 mb s3://your-promptemail-bucket-name
```

2. Configure static website hosting:
```bash
aws s3 website s3://your-promptemail-bucket-name --index-document index.html
```

3. Update `deploy.ps1` with your bucket name:
```powershell
$bucketName = "your-promptemail-bucket-name"
```

4. Update `manifest.xml` with your S3 URLs:
```xml
<IconUrl DefaultValue="https://your-promptemail-bucket-name.s3.amazonaws.com/icons/icon-32.png"/>
```

## Development Workflow

### 1. Local Development

Start the development build watcher:
```bash
npm run dev
```

This will:
- Watch for file changes
- Rebuild automatically
- Output files to the `public/` directory

### 2. Testing Locally

Start the local development server:
```bash
npm start
```

The server will run on `http://localhost:3000` and serve your add-in files.

### 3. Building for Production

Create a production build:
```bash
npm run build
```

This creates optimized files in the `public/` directory.

### 4. Manifest Validation

Validate your manifest file:
```bash
npm run validate-manifest
```

Fix any validation errors before deployment.

## Sideloading for Testing

### Method 1: Outlook Desktop

1. Open Outlook Desktop
2. Go to **File** > **Manage Add-ins** > **My Add-ins**
3. Click **Add a custom add-in** > **Add from file**
4. Select your `manifest.xml` file
5. Click **Install**

### Method 2: Outlook Web App

1. Open Outlook on the web
2. Click **Settings** > **View all Outlook settings**
3. Go to **General** > **Manage add-ins**
4. Click **Add a custom add-in** > **Add from file**
5. Upload your `manifest.xml` file

## Deployment

### 1. Deploy to S3

```bash
# Build and deploy in one command
npm run build && npm run deploy

# Or deploy with dry run first
pwsh -ExecutionPolicy Bypass -File deploy.ps1 -DryRun
pwsh -ExecutionPolicy Bypass -File deploy.ps1
```

### 2. Update Manifest

After deployment, ensure your `manifest.xml` references the correct S3 URLs.

### 3. Test Deployed Version

1. Update your sideloaded add-in with the new manifest
2. Clear Outlook cache if needed
3. Test all functionality

## Debugging

### Browser Developer Tools

1. Open the taskpane in Outlook
2. Right-click in the taskpane
3. Select **Inspect Element** or **Developer Tools**
4. Use console, network, and debugger tabs

### Visual Studio Code

1. Install the Office Add-in Debugger extension
2. Set breakpoints in your JavaScript files
3. Attach the debugger to your running add-in

### Console Logging

The add-in includes comprehensive console logging:
- Check browser console for client-side logs
- Check Windows Application Log for system events

## Troubleshooting

### Common Issues

**Add-in not loading:**
- Check manifest validation
- Verify all URLs are accessible
- Clear Outlook cache: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`

**CSS/JS not updating:**
- Clear browser cache
- Restart Outlook
- Rebuild the project: `npm run build`

**API calls failing:**
- Check API keys and endpoints
- Verify CORS settings
- Check network connectivity

**Icons not displaying:**
- Ensure icon files exist and are accessible
- Check file paths in manifest
- Verify S3 bucket permissions

### Cache Clearing

```bash
# Clear Outlook cache (run as admin)
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\*" -Recurse -Force

# Clear browser cache for Office
# Go to browser settings and clear site data for office.com
```

## File Structure Reference

```
outlook_email_assistant_v3/
├── public/                 # Build output (generated)
├── src/
│   ├── assets/
│   │   ├── css/           # Stylesheets
│   │   └── icons/         # Icon files
│   ├── commands/          # Ribbon command handlers
│   ├── services/          # Business logic services
│   ├── taskpane/          # Main taskpane code
│   └── ui/               # UI components
├── manifest.xml          # Office Add-in manifest
├── deploy.ps1           # Deployment script
├── package.json         # Project configuration
└── webpack.config.js    # Build configuration
```

## API Configuration

### Supported AI Services

- **OpenAI**: Set service to 'openai', provide API key
- **Anthropic**: Set service to 'anthropic', provide API key  
- **Azure OpenAI**: Set service to 'azure', provide endpoint URL and API key
- **Custom**: Set service to 'custom', provide endpoint URL

### Environment Variables (Optional)

Create a `.env` file for default settings:
```
DEFAULT_AI_SERVICE=openai
DEFAULT_MODEL=gpt-4
LOG_LEVEL=info
```

## Security Notes

- Never commit API keys to version control
- Use environment variables for sensitive configuration
- Test classification detection with sample classified emails
- Verify logging excludes sensitive content

## Next Steps

1. Complete icon design and creation
2. Set up AWS S3 bucket and deployment
3. Configure AI service credentials
4. Test with real emails in Outlook
5. Deploy to production S3 bucket
6. Share manifest with users for installation
