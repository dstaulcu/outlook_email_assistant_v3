#!/bin/bash
# Test the Splunk API Gateway endpoint using curl
# Replace with your actual API Gateway URL

API_ENDPOINT="https://23epm9o08b.execute-api.us-east-1.amazonaws.com/prod/telemetry"

echo "Testing API Gateway endpoint: $API_ENDPOINT"
echo "================================="

# Test 1: Simple POST request
echo "Test 1: Simple event"
curl -X POST $API_ENDPOINT \
  -H "Content-Type: application/json" \
  -d '{
    "event": {
      "eventType": "curl_test",
      "message": "Testing from curl",
      "timestamp": "'$(date -u +"%Y-%m-%dT%H:%M:%S.%3NZ")'"
    },
    "sourcetype": "json:outlook_email_assistant",
    "source": "curl_test",
    "index": "main"
  }' \
  -w "\nStatus: %{http_code}\nTime: %{time_total}s\n\n"

# Test 2: CORS preflight
echo "Test 2: CORS preflight (OPTIONS)"
curl -X OPTIONS $API_ENDPOINT \
  -H "Origin: https://your-app-domain.com" \
  -H "Access-Control-Request-Method: POST" \
  -H "Access-Control-Request-Headers: Content-Type" \
  -w "\nStatus: %{http_code}\nTime: %{time_total}s\n\n" \
  -i

echo "Testing completed!"
