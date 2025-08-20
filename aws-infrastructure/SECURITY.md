# 🔒 AWS Infrastructure Security Guidelines

## ⚠️ **Never Commit These Files:**

- `*.pem` - SSH private keys
- `*.key` - Any private key files  
- `*-outputs.json` - Contains sensitive deployment info
- `response.json` - May contain API responses with sensitive data
- `*.zip` - Lambda deployment packages

## ✅ **Safe to Commit:**

- `*.ps1` - PowerShell deployment scripts (reviewed for secrets)
- `*.yaml` - CloudFormation templates (no hardcoded secrets)
- `*.md` - Documentation files
- `*.sh` - Shell scripts (reviewed for secrets)

## 🛡️ **Security Best Practices:**

1. **Always review files before committing**
2. **Use environment variables for secrets**  
3. **Rotate keys regularly**
4. **Use AWS IAM roles instead of hardcoded keys when possible**
5. **Enable AWS CloudTrail for audit logging**

## 🚨 **If You Accidentally Commit Secrets:**

1. **Immediately rotate/delete the compromised keys**
2. **Remove from git history**: `git filter-branch --force --index-filter 'git rm --cached --ignore-unmatch path/to/secret/file' --prune-empty --tag-name-filter cat -- --all`
3. **Force push to all remotes**
4. **Notify team members to rebase their branches**

## 📞 **Emergency Contacts:**
- AWS Security: Report immediately if credentials are compromised
- Team Lead: For internal security protocol escalation
