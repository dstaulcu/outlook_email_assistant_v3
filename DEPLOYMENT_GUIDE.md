# Deployment Guide

## Prerequisites

1. **AWS CLI**: Configured with appropriate S3 permissions
2. **S3 Bucket**: Created and configured for static website hosting
3. **Node.js**: Version 16 or higher
4. **PowerShell**: For deployment script execution

## S3 Bucket Configuration

### 1. Create S3 Bucket

```bash
aws s3 mb s3://your-promptemail-bucket-name
```

### 2. Configure Static Website Hosting

```bash
aws s3 website s3://your-promptemail-bucket-name --index-document index.html
```

### 3. Set Bucket Policy for Public Read Access

Create a bucket policy JSON file:

```json
{
    "Version": "2012-10-17",
    "Statement": [
        {
            "Sid": "PublicReadGetObject",
            "Effect": "Allow",
            "Principal": "*",
            "Action": "s3:GetObject",
            "Resource": "arn:aws:s3:::your-promptemail-bucket-name/*"
        }
    ]
}
```

Apply the policy:
```bash
aws s3api put-bucket-policy --bucket your-promptemail-bucket-name --policy file://bucket-policy.json
```

## Deployment Steps

### 1. Configure Deployment Script

Edit `deploy.ps1` and set your bucket name:

```powershell
$bucketName = "your-promptemail-bucket-name"
```

### 2. Update Manifest URLs

In `manifest.xml`, replace all URLs with your S3 bucket URL:

```xml
https://your-promptemail-bucket-name.s3.amazonaws.com/
```

### 3. Build and Deploy

```bash
# Build the project
npm run build

# Deploy to S3
npm run deploy
```

### 4. Verify Deployment

1. Check that all files are uploaded to S3
2. Verify the website is accessible via S3 URL
3. Test the manifest validation

## Sideloading in Outlook

### Method 1: File Explorer

1. Open Outlook Desktop
2. Go to **File** > **Manage Add-ins** > **My Add-ins**
3. Click **Add a custom add-in** > **Add from file**
4. Select your `manifest.xml` file
5. Click **Install**

### Method 2: Developer Mode

1. Enable Developer Mode in Outlook
2. Use **Insert** > **Get Add-ins** > **My Add-ins** > **Custom Add-ins**
3. Upload manifest file

## Testing Deployment

### 1. Manifest Validation

```bash
npm run validate-manifest
```

### 2. Functional Testing

1. Open Outlook and verify the ribbon button appears
2. Click the button to open the taskpane
3. Test AI features with sample emails
4. Verify classification detection works
5. Check logging functionality

### 3. Accessibility Testing

- Test keyboard navigation
- Verify screen reader compatibility
- Check color contrast ratios

## Troubleshooting

### Common Issues

1. **Add-in not loading**: Check manifest URLs are correct and accessible
2. **Assets not found**: Verify S3 bucket permissions and file paths
3. **CORS errors**: Ensure proper S3 CORS configuration
4. **Cache issues**: Clear Outlook cache and restart

### Cache Clearing

```bash
# Clear Office cache
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

### Logs and Debugging

- Check Windows Application Log for PromptEmail events
- Use browser developer tools in taskpane
- Enable Office Add-in logging

## Rollback Procedure

1. Keep previous version of assets in S3
2. Update manifest URLs to previous version
3. Re-sideload manifest if needed

## Security Considerations

- Never commit AWS credentials to version control
- Use environment variables for sensitive configuration
- Regularly rotate access keys
- Monitor S3 access logs

## Production Checklist

- [ ] Manifest validated
- [ ] All assets uploaded to S3
- [ ] URLs in manifest point to correct S3 paths
- [ ] Functional testing completed
- [ ] Accessibility testing passed
- [ ] Security review completed
- [ ] Documentation updated
