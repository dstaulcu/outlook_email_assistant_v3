const https = require('https');
const http = require('http');
const url = require('url');

/**
 * AWS Lambda function to proxy Splunk HEC requests
 * Handles CORS and credential management
 */
exports.handler = async (event) => {
    console.log('Received event:', JSON.stringify(event, null, 2));
    
    const allowedOrigin = process.env.ALLOWED_ORIGIN || '*';
    const corsHeaders = {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': allowedOrigin,
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Requested-With, X-API-Key'
    };
    
    try {
        // Handle preflight CORS requests
        if (event.httpMethod === 'OPTIONS') {
            console.log('Handling CORS preflight request');
            return createResponse(200, { message: 'CORS preflight handled' }, corsHeaders);
        }
        
        // Only allow POST requests for telemetry
        if (event.httpMethod !== 'POST') {
            console.log('Method not allowed:', event.httpMethod);
            return createResponse(405, { error: 'Method not allowed' }, corsHeaders);
        }
        
        // Validate required environment variables
        if (!process.env.SPLUNK_HEC_TOKEN || !process.env.SPLUNK_HEC_URL) {
            console.error('Missing required environment variables');
            return createResponse(500, { error: 'Service configuration error' }, corsHeaders);
        }
        
        // Parse and validate request body
        let requestData;
        let parsedEvents = [];
        try {
            requestData = typeof event.body === 'string' ? event.body : JSON.stringify(event.body);
            
            // Parse each line as a separate JSON event (Splunk HEC format)
            const lines = requestData.trim().split('\n');
            for (const line of lines) {
                if (line.trim()) {
                    const eventData = JSON.parse(line);
                    
                    // Enrich with server-side information
                    if (eventData.event) {
                        // Add client IP address from API Gateway
                        if (event.requestContext && event.requestContext.identity) {
                            const identity = event.requestContext.identity;
                            if (identity.sourceIp) {
                                eventData.event.client_ip_address = identity.sourceIp;
                            }
                            if (identity.userAgent) {
                                eventData.event.server_user_agent = identity.userAgent; // Server-side UA for verification
                            }
                        }
                                               
                        // Add Lambda execution context
                        //if (process.env.AWS_LAMBDA_FUNCTION_NAME) {
                        //    eventData.event.lambda_function_name = process.env.AWS_LAMBDA_FUNCTION_NAME;
                        //}
                    }
                    
                    parsedEvents.push(eventData);
                }
            }
            
            // Reconstruct the data for Splunk
            requestData = parsedEvents.map(event => JSON.stringify(event)).join('\n');
            
        } catch (parseError) {
            console.error('Invalid JSON in request body:', parseError);
            return createResponse(400, { error: 'Invalid JSON in request body' }, corsHeaders);
        }
        
        // Forward to Splunk
        const splunkResponse = await forwardToSplunk(requestData);
        
        console.log('Splunk response status:', splunkResponse.statusCode);
        return createResponse(splunkResponse.statusCode, splunkResponse.body, corsHeaders);
        
    } catch (error) {
        console.error('Handler error:', error);
        return createResponse(500, { 
            error: 'Internal server error',
            message: error.message 
        }, corsHeaders);
    }
};

/**
 * Forward request to Splunk HEC endpoint
 * @param {string} data - JSON data to send
 * @returns {Promise<Object>} Response from Splunk
 */
async function forwardToSplunk(data) {
    return new Promise((resolve, reject) => {
        try {
            const splunkUrl = new URL(process.env.SPLUNK_HEC_URL);
            
            // Ensure we're hitting the collector/event endpoint
            if (!splunkUrl.pathname.includes('/services/collector')) {
                splunkUrl.pathname = '/services/collector/event';
            }
            
            console.log('Forwarding to Splunk:', splunkUrl.href);
            
            const options = {
                hostname: splunkUrl.hostname,
                port: splunkUrl.port || (splunkUrl.protocol === 'https:' ? 443 : 80),
                path: splunkUrl.pathname + splunkUrl.search,
                method: 'POST',
                headers: {
                    'Authorization': `Splunk ${process.env.SPLUNK_HEC_TOKEN}`,
                    'Content-Type': 'application/json',
                    'Content-Length': Buffer.byteLength(data),
                    'User-Agent': 'OutlookEmailAssistant-Gateway/1.0'
                },
                // Handle self-signed certificates
                rejectUnauthorized: false
            };
            
            const protocol = splunkUrl.protocol === 'https:' ? https : http;
            
            const req = protocol.request(options, (res) => {
                let responseBody = '';
                
                res.on('data', (chunk) => {
                    responseBody += chunk;
                });
                
                res.on('end', () => {
                    console.log(`Splunk response: ${res.statusCode} - ${responseBody}`);
                    
                    let parsedBody;
                    try {
                        parsedBody = JSON.parse(responseBody);
                    } catch (e) {
                        parsedBody = { message: responseBody };
                    }
                    
                    resolve({
                        statusCode: res.statusCode,
                        body: parsedBody
                    });
                });
            });
            
            req.on('error', (error) => {
                console.error('Request to Splunk failed:', error);
                reject(new Error(`Splunk request failed: ${error.message}`));
            });
            
            req.on('timeout', () => {
                console.error('Request to Splunk timed out');
                req.destroy();
                reject(new Error('Request to Splunk timed out'));
            });
            
            // Set timeout
            req.setTimeout(25000); // 25 seconds (Lambda has 30s timeout)
            
            // Send the data
            req.write(data);
            req.end();
            
        } catch (error) {
            console.error('Error setting up Splunk request:', error);
            reject(error);
        }
    });
}

/**
 * Create standardized API response
 * @param {number} statusCode - HTTP status code
 * @param {Object} body - Response body
 * @param {Object} headers - Additional headers
 * @returns {Object} API Gateway response format
 */
function createResponse(statusCode, body, headers = {}) {
    return {
        statusCode: statusCode,
        headers: {
            'Content-Type': 'application/json',
            ...headers
        },
        body: JSON.stringify(body, null, 2)
    };
}
