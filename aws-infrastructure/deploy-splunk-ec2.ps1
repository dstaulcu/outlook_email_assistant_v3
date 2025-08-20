# Deploy Splunk Enterprise on EC2
param(
    [Parameter(Mandatory=$true)]
    [string]$KeyPairName,
    
    [Parameter(Mandatory=$true)]
    [string]$SplunkAdminPassword,
    
    [Parameter(Mandatory=$false)]
    [string]$StackName = "splunk-enterprise-ec2",
    
    [Parameter(Mandatory=$false)]
    [string]$InstanceType = "t3.medium",
    
    [Parameter(Mandatory=$false)]
    [string]$Region = "us-east-1",
    
    [Parameter(Mandatory=$false)]
    [string]$AllowedCidr = "0.0.0.0/0"
)

Write-Host "Deploying Splunk Enterprise on EC2..." -ForegroundColor Green
Write-Host "Stack Name: $StackName" -ForegroundColor Yellow
Write-Host "Instance Type: $InstanceType" -ForegroundColor Yellow
Write-Host "Key Pair: $KeyPairName" -ForegroundColor Yellow
Write-Host "Region: $Region" -ForegroundColor Yellow
Write-Host "WARNING: Allowed CIDR is set to $AllowedCidr - restrict this for production!" -ForegroundColor Red

# Check if key pair exists
$keyPairCheck = aws ec2 describe-key-pairs --key-names $KeyPairName --region $Region 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Error "Key pair '$KeyPairName' not found in region $Region. Please create it first."
    Write-Host "You can create a key pair with:" -ForegroundColor Yellow
    Write-Host "aws ec2 create-key-pair --key-name $KeyPairName --region $Region --query 'KeyMaterial' --output text > $KeyPairName.pem" -ForegroundColor Yellow
    exit 1
}

Write-Host "✓ Key pair found" -ForegroundColor Green

$templatePath = Join-Path $PSScriptRoot "splunk-ec2-template.yaml"

try {
    Write-Host "Deploying CloudFormation stack..." -ForegroundColor Yellow
    
    aws cloudformation deploy `
        --template-file $templatePath `
        --stack-name $StackName `
        --parameter-overrides `
            KeyPairName=$KeyPairName `
            SplunkAdminPassword=$SplunkAdminPassword `
            InstanceType=$InstanceType `
            AllowedCidr=$AllowedCidr `
        --capabilities CAPABILITY_IAM `
        --region $Region
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Stack deployed successfully!" -ForegroundColor Green
        
        # Get stack outputs
        Write-Host "`nRetrieving stack outputs..." -ForegroundColor Yellow
        $outputs = aws cloudformation describe-stacks `
            --stack-name $StackName `
            --region $Region `
            --query 'Stacks[0].Outputs' `
            --output json | ConvertFrom-Json
        
        Write-Host "`n=== SPLUNK INSTANCE DETAILS ===" -ForegroundColor Green
        $publicIp = ""
        $webUrl = ""
        $hecUrl = ""
        $sshCommand = ""
        
        foreach ($output in $outputs) {
            Write-Host "$($output.OutputKey): $($output.OutputValue)" -ForegroundColor Yellow
            
            switch ($output.OutputKey) {
                "PublicIP" { $publicIp = $output.OutputValue }
                "SplunkWebUrl" { $webUrl = $output.OutputValue }
                "SplunkHecUrl" { $hecUrl = $output.OutputValue }
                "SSHCommand" { $sshCommand = $output.OutputValue }
            }
        }
        
        Write-Host "`n=== NEXT STEPS ===" -ForegroundColor Green
        Write-Host "1. Wait 5-10 minutes for Splunk to fully start up" -ForegroundColor White
        Write-Host "2. Access Splunk Web UI: $webUrl" -ForegroundColor White
        Write-Host "   Username: admin" -ForegroundColor White
        Write-Host "   Password: $SplunkAdminPassword" -ForegroundColor White
        Write-Host "3. SSH to instance: $sshCommand" -ForegroundColor White
        Write-Host "4. Update your API Gateway with HEC URL: $hecUrl" -ForegroundColor White
        
        Write-Host "`n=== UPDATE API GATEWAY ===" -ForegroundColor Green
        Write-Host "Run this command to update your API Gateway with the new Splunk URL:" -ForegroundColor White
        Write-Host ".\deploy-splunk-gateway.ps1 ``" -ForegroundColor Cyan
        Write-Host "    -StackName `"outlook-assistant-splunk-gateway-prod`" ``" -ForegroundColor Cyan
        Write-Host "    -SplunkHecToken `"520fe85b-68f1-4a82-9131-33d9e5a5cddd`" ``" -ForegroundColor Cyan
        Write-Host "    -SplunkHecUrl `"$hecUrl`"" -ForegroundColor Cyan
        
        # Save outputs to file
        $outputsFile = Join-Path $PSScriptRoot "splunk-deployment-outputs.json"
        $outputs | ConvertTo-Json -Depth 3 | Out-File -FilePath $outputsFile -Encoding UTF8
        Write-Host "`nOutputs saved to: $outputsFile" -ForegroundColor Green
        
    } else {
        Write-Error "Stack deployment failed"
        exit 1
    }
    
} catch {
    Write-Error "Deployment failed: $($_.Exception.Message)"
    exit 1
}

Write-Host "`n✓ Splunk Enterprise deployment completed!" -ForegroundColor Green
