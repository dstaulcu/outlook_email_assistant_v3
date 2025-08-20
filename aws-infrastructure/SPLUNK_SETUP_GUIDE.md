# Splunk on AWS Setup Guide

This guide will walk you through setting up a Splunk Enterprise instance on AWS EC2.

## Prerequisites

1. **AWS CLI configured** with appropriate permissions
2. **EC2 Key Pair** for SSH access to the instance
3. **PowerShell** (Windows PowerShell or PowerShell Core)

## Step 1: Create EC2 Key Pair (if you don't have one)

```powershell
# Create a new key pair
aws ec2 create-key-pair --key-name splunk-keypair --region us-east-1 --query 'KeyMaterial' --output text > splunk-keypair.pem

# Set appropriate permissions (Linux/Mac)
chmod 400 splunk-keypair.pem
```

## Step 2: Deploy Splunk Instance

```powershell
.\deploy-splunk-ec2.ps1 `
    -KeyPairName "splunk-keypair" `
    -SplunkAdminPassword "YourStrongPassword123!" `
    -StackName "splunk-enterprise" `
    -InstanceType "t3.medium"
```

### Parameters Explained

- `KeyPairName`: Your EC2 key pair name for SSH access
- `SplunkAdminPassword`: Admin password for Splunk (minimum 8 characters)
- `StackName`: CloudFormation stack name (optional)
- `InstanceType`: EC2 instance type (t3.medium recommended for testing)
- `AllowedCidr`: IP range allowed to access Splunk (default: 0.0.0.0/0)

## Step 3: Wait for Deployment

The deployment takes about 10-15 minutes:
- CloudFormation creates the infrastructure (2-3 minutes)
- EC2 instance starts and installs Docker (2-3 minutes)  
- Splunk container downloads and starts (5-10 minutes)

## Step 4: Access Splunk

Once deployment completes, you'll see:

```
=== SPLUNK INSTANCE DETAILS ===
PublicIP: 54.123.456.789
SplunkWebUrl: http://54.123.456.789:8000
SplunkHecUrl: http://54.123.456.789:8088
SSHCommand: ssh -i splunk-keypair.pem ec2-user@54.123.456.789
```

### Access Splunk Web UI

1. Open: `http://your-public-ip:8000`
2. Login:
   - Username: `admin`
   - Password: `YourStrongPassword123!`

### Verify HEC is Working

Test the HTTP Event Collector:

```powershell
$splunkHecUrl = "http://your-public-ip:8088/services/collector/event"
$headers = @{
    "Authorization" = "Splunk 520fe85b-68f1-4a82-9131-33d9e5a5cddd"
    "Content-Type" = "application/json"
}
$body = @{
    event = @{
        message = "Test from PowerShell"
        timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    }
    sourcetype = "json:test"
} | ConvertTo-Json -Depth 3

Invoke-RestMethod -Uri $splunkHecUrl -Method POST -Headers $headers -Body $body
```

## Step 5: Update API Gateway

Update your existing API Gateway to use the new Splunk instance:

```powershell
.\deploy-splunk-gateway.ps1 `
    -StackName "outlook-assistant-splunk-gateway-prod" `
    -SplunkHecToken "520fe85b-68f1-4a82-9131-33d9e5a5cddd" `
    -SplunkHecUrl "http://your-public-ip:8088"
```

## Step 6: Test End-to-End

Run the API Gateway test:

```powershell
.\test-api-gateway.ps1
```

All tests should now pass! 🎉

## Security Considerations

### For Production:

1. **Restrict access**: Change `AllowedCidr` from `0.0.0.0/0` to your specific IP ranges
2. **Use HTTPS**: Configure SSL certificates
3. **VPC Peering**: Place Splunk in private subnet with VPC peering to Lambda
4. **Backup**: Configure regular backups of Splunk data
5. **Monitoring**: Set up CloudWatch alarms

### Example secure deployment:

```powershell
.\deploy-splunk-ec2.ps1 `
    -KeyPairName "splunk-keypair" `
    -SplunkAdminPassword "YourStrongPassword123!" `
    -AllowedCidr "203.0.113.0/24" `  # Your office IP range
    -InstanceType "t3.large"
```

## Troubleshooting

### Splunk won't start
```bash
# SSH to instance and check logs
ssh -i splunk-keypair.pem ec2-user@your-ip
sudo docker logs splunk-enterprise
```

### Can't access web UI
- Check security group allows port 8000 from your IP
- Verify instance is running: `aws ec2 describe-instances`

### HEC not working
```bash
# Check HEC status
docker exec splunk-enterprise /opt/splunk/bin/splunk http-event-collector list
```

## Cost Estimation

- **t3.medium**: ~$30/month (24/7)
- **t3.large**: ~$60/month (24/7)
- **Storage**: ~$8/month per 80GB
- **Data transfer**: Variable based on usage

For development/testing, consider stopping the instance when not in use to save costs.

## Cleanup

To remove all resources:

```powershell
aws cloudformation delete-stack --stack-name splunk-enterprise --region us-east-1
```
