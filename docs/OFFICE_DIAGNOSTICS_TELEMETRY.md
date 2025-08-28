# Office Diagnostic Values in Telemetry Events

## Overview
All telemetry events now automatically include Office diagnostic information to provide better context about the user's environment and improve debugging capabilities.

## Implementation Details

### What's Added
Every telemetry event now includes flattened diagnostic fields for easier querying:

**Office Diagnostic Fields:**
- **office_host**: The Office application (e.g., "Outlook", "Word", "Excel")
- **office_platform**: The operating system/platform ("Windows", "Mac", "Web", "iOS", "Android")
- **office_version**: The Office version number (e.g., "16.0.14332.20130")
- **office_owa_view**: For Outlook Web Access, the current view mode (optional)
- **office_error**: Any error encountered while gathering diagnostics (optional)

**User Profile Fields (flattened):**
- **userProfile_displayName**: User's display name
- **userProfile_emailAddress**: User's email address
- **userProfile_timeZone**: User's time zone
- **userProfile_accountType**: Type of email account (e.g., "exchange", "gmail")

**Environment Fields (flattened):**
- **environment_type**: Detected environment (Dev, Test, Prod, Local, unknown)
- **environment_host**: The hostname/domain where the add-in is running
- **environment_error**: Any error encountered while detecting environment (optional)

**Client Context Fields (flattened):**
- **client_browser_name**: Browser name (Chrome, Firefox, Safari, Edge, etc.)
- **client_browser_version**: Browser version number
- **client_platform**: Operating system platform (Win32, MacIntel, Linux, etc.)
- **client_language**: Primary browser language (e.g., "en-US")
- **client_timezone**: Client timezone (e.g., "America/New_York")
- **client_screen_resolution**: Screen resolution (e.g., "1920x1080")
- **client_viewport_size**: Browser viewport size (e.g., "1024x768")
- **client_connection_type**: Network connection type ("4g", "wifi", etc.)
- **client_cpu_cores**: Number of CPU cores available
- **client_device_memory_gb**: Device memory in gigabytes (if available)
- **client_js_heap_size_mb**: JavaScript heap usage in megabytes
- **client_connection_rtt_ms**: Network round-trip time in milliseconds
- **client_connection_downlink_mbps**: Download speed in Mbps
- **client_error**: Any error encountered while gathering client info (optional)

**Performance Metrics (flattened):**
- **analysis_duration_ms**: Time taken for email analysis in milliseconds
- **response_generation_duration_ms**: Time taken for response generation in milliseconds (when applicable)
- **total_duration_ms**: Total time for combined operations in milliseconds (when applicable)

### Code Location
The enhancement is implemented in `src/services/Logger.js`:
- `getOfficeDiagnostics()` method gathers Office context information
- `createLogEntry()` method automatically flattens and includes Office diagnostics in all events
- `logToApiGateway()` method flattens environment context
- Graceful fallback when Office context is unavailable

**Note:** The structure was flattened for better Splunk querying performance. Previous nested objects like `office.userProfile.emailAddress` are now top-level fields like `userProfile_emailAddress`.

### Example Telemetry Event (Flattened Structure)
```json
{
  "eventType": "email_analyzed",
  "timestamp": "2025-08-28T04:35:18.571Z",
  "source": "promptemail",
  "version": "1.2.3",
  "sessionId": "sess_1756355693822_ip835fpqb",
  
  // Office diagnostic fields (flattened)
  "office_host": "Outlook",
  "office_platform": "Windows", 
  "office_version": "16.0.14332.20130",
  "office_owa_view": "ReadingPane",
  
  // User profile fields (flattened)
  "userProfile_displayName": "John Doe",
  "userProfile_emailAddress": "john.doe@company.com",
  "userProfile_timeZone": "Pacific Standard Time",
  "userProfile_accountType": "exchange",
  
  // Environment fields (flattened)
  "environment_type": "Prod",
  "environment_host": "293354421824-outlook-email-assistant-prod.s3.us-east-1.amazonaws.com",
  
  // Client context fields (flattened)
  "client_browser_name": "Chrome",
  "client_browser_version": "116.0",
  "client_platform": "Win32",
  "client_language": "en-US",
  "client_timezone": "America/New_York",
  "client_screen_resolution": "1920x1080",
  "client_viewport_size": "1024x768",
  "client_connection_type": "4g",
  "client_cpu_cores": 8,
  "client_device_memory_gb": 16,
  "client_js_heap_size_mb": 45,
  "client_connection_rtt_ms": 50,
  "client_connection_downlink_mbps": 25.5,
  
  // Performance metrics (flattened)
  "analysis_duration_ms": 2400,
  
  // Server-side enrichment (added by API Gateway Lambda)
  "client_ip_address": "203.0.113.42",
  "request_id": "c6af9ac6-7b61-11e6-9a41-93e8deadbeef",
  "api_gateway_stage": "prod",
  "server_received_time": 1724798215571,
  "server_user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)...",
  "lambda_function_name": "outlook-telemetry-proxy",
  "lambda_function_version": "12",
  
  // Original event data
  "model_service": "ollama",
  "model_name": "llama3:latest",
  "email_length": 1260,
  "recipients_count": 1,
  "analysis_success": true,
  "refinement_count": 0,
  "clipboard_used": false
}
```

**Server-side Fields (added by API Gateway Lambda):**
- **client_ip_address**: Client IP address (captured server-side for security)
- **request_id**: API Gateway request ID for tracing
- **api_gateway_stage**: Deployment stage (prod, dev, test)
- **server_received_time**: Server timestamp when request was received
- **server_user_agent**: Server-side captured user agent (for verification)
- **lambda_function_name**: AWS Lambda function processing the request
- **lambda_function_version**: Lambda function version

For combined analysis + response generation (`auto_analysis_completed` event):
```json
{
  "eventType": "auto_analysis_completed",
  // ... other fields ...
  
  // Performance metrics for combined operations
  "analysis_duration_ms": 1200,
  "response_generation_duration_ms": 1800,
  "total_duration_ms": 3000
}
```

Note: The Splunk/API Gateway `host` field will use the environment host domain (e.g., `293354421824-outlook-email-assistant-prod.s3.us-east-1.amazonaws.com`) which provides clear environment identification in your telemetry system.

## Benefits

### 1. Enhanced Debugging
- Quickly identify platform-specific issues
- Understand which Office version/environment issues occur in
- Better context for reproducing user-reported problems

### 2. Improved Analytics & Querying
- **Flattened structure**: Easier Splunk queries like `office_host="Outlook"` instead of `office.host="Outlook"`
- **Better performance**: Faster searches without nested field parsing
- **Simpler filtering**: Direct field access for environment-specific analysis
- Track add-in usage across different Office versions
- Identify most common user environments
- Plan feature development based on platform usage

### 3. Support & Troubleshooting
- Customer support can immediately see user's Office environment
- Faster issue resolution with environmental context
- Proactive identification of compatibility issues

### 4. User Context & Support
- Understand email account types users are connecting (Exchange, Gmail, etc.)
- Time zone information for better understanding of usage patterns
- User identification for better customer support
- Account type analysis for feature compatibility

### 5. Environment & Deployment Tracking
- Automatic detection of deployment environment (Dev, Test, Prod)
- S3 host domain included in Splunk/API Gateway host field for easy filtering
- Clear separation of telemetry data by environment
- Helps with environment-specific issue tracking and performance analysis

### 6. Feature Planning
- Understand user environment distribution across different Office versions
- Make informed decisions about minimum system requirements
- Track adoption of newer Office versions
- Plan features based on account type distribution

## Fallback Behavior
When the add-in runs outside of Office (e.g., in a standalone browser):
```json
{
  "office": {
    "host": "Non-Office",
    "platform": "Web",
    "version": "N/A"
  }
}
```

## Impact on Existing Events
All existing telemetry events will now include the Office diagnostic information:
- `session_start`
- `session_summary` 
- `email_analyzed`
- `auto_analysis_completed`
- `email_context_detected`
- `initial_setup_prompted`
- `session_telemetry`
- `processing_metrics`
- `security_event`

## Privacy & Compliance
- Office diagnostic values are technical metadata about the application environment
- User profile information (displayName, emailAddress) should be handled according to your privacy policy
- Time zone and account type information helps with service optimization
- Version information helps with security patching and support
- Consider data anonymization or hashing for sensitive user profile fields
- Follows existing telemetry privacy practices

## Testing
Use the example script at `docs/office-diagnostics-telemetry-example.js` to see the expected telemetry event structure.

## Configuration
No additional configuration required - Office diagnostics are automatically included in all telemetry events when the Logger service is used.
