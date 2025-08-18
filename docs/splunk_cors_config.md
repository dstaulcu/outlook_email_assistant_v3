# Splunk CORS Configuration for Office Add-in

## For Splunk 6.4+ (HEC-specific CORS)

Add to `inputs.conf` under the `[http]` stanza:

```ini
[http]
crossOriginSharingPolicy = *
# OR for more security, specify the Office add-in domain:
# crossOriginSharingPolicy = https://293354421824-outlook-email-assistant-prd.s3.us-east-1.amazonaws.com
```

## For Splunk 6.3 (General REST API CORS)

Add to `server.conf` under the `[httpserver]` stanza:

```ini
[httpserver]
crossOriginSharingPolicy = *
```

## Location of Configuration Files

Typically found at:
- Windows: `%SPLUNK_HOME%\etc\system\local\inputs.conf`
- Linux/Mac: `$SPLUNK_HOME/etc/system/local/inputs.conf`

## After Making Changes

1. Restart Splunk service
2. Verify HEC endpoint is accessible
3. Test with the Office add-in

## Security Note

Using `*` allows all origins. For production, replace with specific domain:
`crossOriginSharingPolicy = https://your-office-addin-domain.com`
