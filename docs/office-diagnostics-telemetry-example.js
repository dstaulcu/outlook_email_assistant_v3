/**
 * Test script for Office Diagnostic Values in Telemetry
 * This script demonstrates the flattened Office diagnostic information being captured
 * in telemetry events.
 */

// Example of what the new telemetry events will look like with flattened structure
const exampleTelemetryEvent = {
    "eventType": "email_analyzed",
    "timestamp": "2025-08-28T04:35:18.571Z",
    "source": "promptemail",
    "version": "1.2.3",
    "sessionId": "sess_1756355693822_ip835fpqb",
    
    // Office diagnostic fields (flattened from office.*)
    "office_host": "Outlook",
    "office_platform": "Windows",
    "office_version": "16.0.14332.20130",
    "office_owa_view": "ReadingPane", // if available
    
    // User profile fields (flattened from office.userProfile.*)
    "userProfile_displayName": "John Doe",
    "userProfile_emailAddress": "john.doe@company.com", 
    "userProfile_timeZone": "Pacific Standard Time",
    "userProfile_accountType": "exchange",
    
    // Environment fields (flattened from environment.*)
    "environment_type": "Prod",
    "environment_host": "293354421824-outlook-email-assistant-prod.s3.us-east-1.amazonaws.com",
    
    // Client context fields (flattened from client.*)
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
    
    // Performance metrics (flattened from performance_metrics.*)
    "analysis_duration_ms": 2400,
    
    // Server-side enrichment (added by API Gateway Lambda - not in client-side data)
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
};

console.info('[INFO] - Example telemetry event with flattened Office diagnostics:');
console.log(JSON.stringify(exampleTelemetryEvent, null, 2));

// Example of combined analysis + response generation event
const combinedOperationEvent = {
    "eventType": "auto_analysis_completed",
    "timestamp": "2025-08-28T04:35:21.571Z",
    "source": "promptemail",
    "version": "1.2.3",
    "sessionId": "sess_1756355693822_ip835fpqb",
    
    // Office and environment fields (same as above)
    "office_host": "Outlook",
    "office_platform": "Windows",
    "userProfile_emailAddress": "john.doe@company.com",
    "environment_type": "Dev",
    
    // Performance metrics for combined operations (flattened)
    "analysis_duration_ms": 1200,
    "response_generation_duration_ms": 1800,
    "total_duration_ms": 3000,
    
    // Event-specific data
    "model_service": "ollama",
    "model_name": "llama3:latest",
    "auto_response_generated": true,
    "email_context": "inbox",
    "generation_type": "standard_response"
};

console.info('[INFO] - \nExample combined operation event:');
console.log(JSON.stringify(combinedOperationEvent, null, 2));

/**
 * Office diagnostic values that will be captured (now flattened):
 * 
 * 1. office_*: Office application context (host, platform, version, owa view)
 * 2. userProfile_*: User profile information (displayName, emailAddress, timeZone, accountType)
 * 3. environment_*: Deployment environment context (type, host)
 * 4. client_*: Browser and client system information
 *    - Browser: client_browser_name, client_browser_version
 *    - System: client_platform, client_cpu_cores, client_device_memory_gb
 *    - Display: client_screen_resolution, client_viewport_size
 *    - Network: client_connection_type, client_connection_rtt_ms, client_connection_downlink_mbps
 *    - Locale: client_language, client_timezone
 *    - Performance: client_js_heap_size_mb
 * 5. Performance metrics: *_duration_ms fields
 * 6. Error fields: office_error, client_error, environment_error
 * 
 * Benefits:
 * - Better understanding of user environment for debugging
 * - Platform-specific issue identification
 * - Version compatibility tracking  
 * - Feature support analysis across different Office versions
 * - Better customer support with environmental context
 * - Network performance correlation with user experience
 * - Client-side performance monitoring
 * - Easier querying with flattened field names
 * - Better Splunk search performance
 * - Cleaner analytics and reporting
 * 
 * Note: Client IP address and other server-side fields are automatically
 * added by the API Gateway Lambda function for security and compliance tracking.
 */

// Fallback behavior when Office context is not available
const fallbackExample = {
    "eventType": "session_start",
    "timestamp": "2025-08-27T22:30:15.123Z",
    "source": "promptemail",
    "version": "1.0.0",
    "sessionId": "sess_1724798215123_xyz789",
    
    // Office diagnostics when running outside Office (flattened)
    "office_host": "Non-Office",
    "office_platform": "Web",
    "office_version": "N/A",
    "office_error": "Office context not available",
    
    // Environment context still captured
    "environment_type": "Local",
    "environment_host": "localhost"
};

console.info('[INFO] - \nFallback telemetry event when Office context unavailable:');
console.log(JSON.stringify(fallbackExample, null, 2));
