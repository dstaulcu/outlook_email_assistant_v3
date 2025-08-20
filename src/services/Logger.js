/**
 * Logger Service
 * Handles console logging (always enabled) and Splunk HEC telemetry (configurable)
 */

export class Logger {
    constructor() {
        this.eventSource = 'PromptEmail';
        this.isEnabled = true;
        this.telemetryConfig = null;
        this.splunkQueue = [];
        this.splunkFlushInterval = null;
        this.splunkRetryCount = 0;
        this.maxRetries = 3;
        this.splunkConnectionError = false;
        
        // Add unique instance ID for debugging
        this.instanceId = Date.now() + '-' + Math.random().toString(36).substr(2, 9);
        console.debug('Logger instance created:', this.instanceId);
        
        // Initialize telemetry configuration
        this.initializeTelemetryConfig();
    }

    /**
     * Initialize telemetry configuration
     */
    async initializeTelemetryConfig() {
        try {
            console.debug('Loading telemetry configuration... (instance:', this.instanceId, ')');
            const response = await fetch('/config/telemetry.json');
            if (response.ok) {
                this.telemetryConfig = await response.json();
                console.info('Telemetry configuration loaded (instance:', this.instanceId, '):', this.telemetryConfig);
            } else {
                console.warn('Could not load telemetry configuration, using defaults');
                this.telemetryConfig = this.getDefaultTelemetryConfig();
            }
        } catch (error) {
            console.error('Failed to load telemetry configuration:', error);
            this.telemetryConfig = this.getDefaultTelemetryConfig();
        }
    }

    /**
     * Get default telemetry configuration
     */
    getDefaultTelemetryConfig() {
        return {
            telemetry: {
                enabled: false,
                provider: "local"
            },
            environment: {
                name: "unknown",
                version: "1.0.0"
            }
        };
    }

    /**
     * Logs an event to console (always) and Splunk HEC (if enabled)
     * @param {string} eventType - Type of event (e.g., 'session_start', 'email_analyzed')
     * @param {Object} data - Event data object
     * @param {string} level - Log level ('Information', 'Warning', 'Error')
     * @param {string} contextEmail - Optional email address for better user context
     */
    async logEvent(eventType, data = {}, level = 'Information', contextEmail = null) {
        if (!this.isEnabled) {
            console.debug('Logging disabled:', eventType, data);
            return;
        }

        try {
            const logEntry = this.createLogEntry(eventType, data, contextEmail);
            
            // Always log to console for development/debugging
            console.debug(`[${level}] ${eventType}:`, logEntry);

            // Add to telemetry queue if enabled
            if (this.telemetryConfig?.telemetry?.enabled) {
                if (this.telemetryConfig.telemetry.provider === 'splunk_hec') {
                    await this.logToSplunk(eventType, logEntry, level);
                } else if (this.telemetryConfig.telemetry.provider === 'api_gateway') {
                    await this.logToApiGateway(eventType, logEntry, level);
                }
            }

        } catch (error) {
            console.error('Failed to log event:', error);
        }
    }

    /**
     * Creates a standardized log entry
     * @param {string} eventType - Event type
     * @param {Object} data - Event data
     * @param {string} contextEmail - Optional email context for user identification
     * @returns {Object} Log entry object
     */
    createLogEntry(eventType, data, contextEmail = null) {
        const baseEntry = {
            eventType: eventType,
            timestamp: new Date().toISOString(),
            source: this.eventSource,
            version: '1.0.0',
            sessionId: this.getSessionId(),
            userId: this.getUserId(contextEmail)
        };

        // Sanitize sensitive data
        const sanitizedData = this.sanitizeData(data);

        return {
            ...baseEntry,
            ...sanitizedData
        };
    }

    /**
     * Sanitizes data to remove sensitive information
     * @param {Object} data - Raw data object
     * @returns {Object} Sanitized data
     */
    sanitizeData(data) {
        const sanitized = { ...data };

        // Remove sensitive fields
        const sensitiveFields = [
            'apiKey', 'api_key', 'password', 'token', 'secret',
            'emailBody', 'email_body', 'content', 'body',
            'personalInfo', 'personal_info'
        ];

        sensitiveFields.forEach(field => {
            if (sanitized[field]) {
                delete sanitized[field];
            }
        });

        // Sanitize classification data
        if (sanitized.classification && sanitized.subject) {
            sanitized.subject = '[REDACTED]';
        }

        // Ensure no long text fields
        Object.keys(sanitized).forEach(key => {
            if (typeof sanitized[key] === 'string' && sanitized[key].length > 500) {
                sanitized[key] = sanitized[key].substring(0, 500) + '...[TRUNCATED]';
            }
        });

        return sanitized;
    }

    /**
     * Logs an event to Splunk HEC
     * @param {string} eventType - Event type
     * @param {Object} logEntry - Log entry data
     * @param {string} level - Log level
     */
    async logToSplunk(eventType, logEntry, level) {
        try {
            if (!this.telemetryConfig?.telemetry?.splunk) {
                console.warn('Splunk configuration not available');
                return;
            }

            const splunkConfig = this.telemetryConfig.telemetry.splunk;
            const splunkEvent = {
                time: Math.floor(Date.now() / 1000), // Unix timestamp
                host: window.location.hostname,
                index: splunkConfig.index,
                source: splunkConfig.source,
                sourcetype: splunkConfig.sourcetype,
                event: {
                    ...logEntry,
                    level: level,
                    environment: this.telemetryConfig.environment
                }
            };

            console.debug('Preparing Splunk event:', splunkEvent);

            // Add to Splunk queue for batch processing
            this.splunkQueue.push(splunkEvent);

            // Process Splunk queue if it's getting full or on timer
            if (this.splunkQueue.length >= (splunkConfig.batchSize || 10)) {
                await this.flushSplunkQueue();
            }

        } catch (error) {
            console.error('Failed to log to Splunk:', error);
        }
    }

    /**
     * Logs an event to API Gateway
     * @param {string} eventType - Event type
     * @param {Object} logEntry - Log entry data  
     * @param {string} level - Log level
     */
    async logToApiGateway(eventType, logEntry, level) {
        try {
            if (!this.telemetryConfig?.telemetry?.api_gateway) {
                console.warn('API Gateway configuration not available');
                return;
            }

            const apiGatewayConfig = this.telemetryConfig.telemetry.api_gateway;
            
            // Create event in Splunk-compatible format for the API Gateway to forward
            const splunkEvent = {
                time: Math.floor(Date.now() / 1000),
                host: window.location.hostname,
                source: "outlook_addon", 
                sourcetype: "json:outlook_email_assistant",
                event: {
                    ...logEntry,
                    level: level,
                    environment: this.telemetryConfig.environment
                }
            };

            console.debug('Preparing API Gateway event:', splunkEvent);

            // Add to Splunk queue for batch processing (reuse the same queue)
            this.splunkQueue.push(splunkEvent);

            // Process queue if it's getting full or on timer
            if (this.splunkQueue.length >= (apiGatewayConfig.batchSize || 10)) {
                await this.flushApiGatewayQueue();
            }

        } catch (error) {
            console.error('Failed to log to API Gateway:', error);
        }
    }

    /**
     * Flush Splunk queue to HEC endpoint
     */
    async flushSplunkQueue() {
        if (this.splunkQueue.length === 0) return;

        const events = [...this.splunkQueue];
        this.splunkQueue = [];

        try {
            const splunkConfig = this.telemetryConfig.telemetry.splunk;

            console.debug(`[Logger] Flushing ${events.length} events to Splunk HEC`);

            // Prepare fetch options
            const fetchOptions = {
                method: 'POST',
                headers: {
                    'Authorization': `Splunk ${splunkConfig.hecToken}`,
                    'Content-Type': 'application/json'
                },
                body: events.map(event => JSON.stringify(event)).join('\n')
            };

            // For development, if validateCertificate is false and endpoint is HTTPS,
            // warn user to use HTTP endpoint instead
            if (!splunkConfig.validateCertificate && splunkConfig.hecEndpoint.startsWith('https://')) {
                console.warn('Certificate validation disabled but HTTPS endpoint used. Consider using HTTP endpoint for development.');
            }

            console.debug(`[Logger] Attempting to send to: ${splunkConfig.hecEndpoint}`);
            console.debug(`[Logger] Request headers:`, fetchOptions.headers);
            
            const response = await fetch(splunkConfig.hecEndpoint, fetchOptions);

            console.debug(`[Logger] Response status: ${response.status} ${response.statusText}`);

            if (response.ok) {
                console.info('Successfully sent events to Splunk HEC');
                
                // Reset connection error state on successful connection
                this.splunkConnectionError = false;
                this.splunkRetryCount = 0;
                
                // Log successful telemetry transmission
                const result = await response.json();
                console.debug('Splunk HEC response:', result);
            } else {
                console.error('Failed to send events to Splunk HEC:', response.status, response.statusText);
                
                // Re-queue events for retry (with limit)
                if (events.length < 100) {
                    this.splunkQueue.unshift(...events);
                }
            }

        } catch (error) {
            console.debug(`[Logger] Fetch error details:`, {
                message: error.message,
                name: error.name,
                stack: error.stack?.substring(0, 200)
            });
            
            // Handle connection errors more gracefully
            if (error.message.includes('fetch') || 
                error.message.includes('ERR_CONNECTION_REFUSED') ||
                error.message.includes('ERR_CERT_AUTHORITY_INVALID') ||
                error.message.includes('ERR_CERT_COMMON_NAME_INVALID') ||
                error.message.includes('ERR_BLOCKED_BY_PRIVATE_NETWORK_ACCESS_CHECKS')) {
                
                // Special handling for private network access errors
                if (error.message.includes('ERR_BLOCKED_BY_PRIVATE_NETWORK_ACCESS_CHECKS')) {
                    if (!this.splunkConnectionError) {
                        console.warn('Private Network Access blocked connection to localhost. Office add-ins cannot access localhost directly. Consider using a proxy or cloud endpoint for telemetry.');
                        this.splunkConnectionError = true;
                    }
                }
                // Special handling for certificate errors
                else if (error.message.includes('ERR_CERT_')) {
                    if (!this.splunkConnectionError) {
                        const splunkConfig = this.telemetryConfig.telemetry.splunk;
                        if (!splunkConfig.validateCertificate && splunkConfig.hecEndpoint.startsWith('https://')) {
                            const httpEndpoint = splunkConfig.hecEndpoint.replace('https://', 'http://');
                            console.warn(`[Logger] SSL certificate error detected. Since validateCertificate=false, consider changing endpoint from ${splunkConfig.hecEndpoint} to ${httpEndpoint}`);
                        } else {
                            console.warn('Splunk HEC SSL certificate validation failed');
                        }
                        this.splunkConnectionError = true;
                    }
                } else if (!this.splunkConnectionError) {
                    console.warn('Splunk HEC connection unavailable, events will be queued');
                    this.splunkConnectionError = true;
                }
                
                // Re-queue events for retry (with limit)
                if (events.length < 100) { // Prevent infinite queue growth
                    this.splunkQueue.unshift(...events);
                    this.splunkRetryCount++;
                    
                    // If too many retries, clear the queue to prevent memory issues
                    if (this.splunkRetryCount > this.maxRetries) {
                        console.warn('Max Splunk retry attempts reached, clearing queue');
                        this.splunkQueue = [];
                        this.splunkRetryCount = 0;
                    }
                }
            } else {
                console.error('Error flushing Splunk queue:', error);
            }
        }
    }

    /**
     * Flush queue to API Gateway endpoint
     */
    async flushApiGatewayQueue() {
        if (this.splunkQueue.length === 0) return;

        const events = [...this.splunkQueue];
        this.splunkQueue = [];

        try {
            const apiGatewayConfig = this.telemetryConfig.telemetry.api_gateway;

            console.debug(`[Logger] Flushing ${events.length} events to API Gateway`);

            // Prepare fetch options - Send events one by one like Splunk HEC
            const fetchOptions = {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: events.map(event => JSON.stringify(event)).join('\n')
            };

            console.debug(`[Logger] Attempting to send to: ${apiGatewayConfig.endpoint}`);
            
            const response = await fetch(apiGatewayConfig.endpoint, fetchOptions);

            console.debug(`[Logger] Response status: ${response.status} ${response.statusText}`);

            if (response.ok) {
                console.info('Successfully sent events to API Gateway');
                
                // Reset connection error state on successful connection
                this.splunkConnectionError = false;
                this.splunkRetryCount = 0;
                
                // Log successful telemetry transmission
                const result = await response.json();
                console.debug('API Gateway response:', result);
            } else {
                console.error('Failed to send events to API Gateway:', response.status, response.statusText);
                
                // Re-queue events for retry (with limit)
                if (events.length < 100) {
                    this.splunkQueue.unshift(...events);
                }
            }

        } catch (error) {
            console.debug(`[Logger] API Gateway error details:`, {
                message: error.message,
                name: error.name,
                stack: error.stack?.substring(0, 200)
            });
            
            // Handle connection errors
            if (error.message.includes('fetch') || error.message.includes('NetworkError')) {
                if (!this.splunkConnectionError) {
                    console.warn('API Gateway connection unavailable, events will be queued');
                    this.splunkConnectionError = true;
                }
                
                // Re-queue events for retry (with limit)
                if (events.length < 100) {
                    this.splunkQueue.unshift(...events);
                    this.splunkRetryCount++;
                    
                    // If too many retries, clear the queue to prevent memory issues
                    if (this.splunkRetryCount > this.maxRetries) {
                        console.warn('Max API Gateway retry attempts reached, clearing queue');
                        this.splunkQueue = [];
                        this.splunkRetryCount = 0;
                    }
                }
            } else {
                console.error('Error flushing API Gateway queue:', error);
            }
        }
    }

    /**
     * Start automatic queue flushing for current provider
     */
    startSplunkAutoFlush() {
        if (this.splunkFlushInterval) return;

        const provider = this.telemetryConfig?.telemetry?.provider;
        let flushInterval;
        
        if (provider === 'api_gateway') {
            flushInterval = this.telemetryConfig?.telemetry?.api_gateway?.flushInterval || 60000;
        } else {
            flushInterval = this.telemetryConfig?.telemetry?.splunk?.flushInterval || 60000;
        }

        this.splunkFlushInterval = setInterval(async () => {
            if (this.splunkQueue.length > 0) {
                if (provider === 'api_gateway') {
                    await this.flushApiGatewayQueue();
                } else if (provider === 'splunk_hec') {
                    await this.flushSplunkQueue();
                }
            }
        }, flushInterval);

        console.log(`[Logger] Started ${provider} auto-flush with ${flushInterval}ms interval`);
    }

    /**
     * Stop automatic queue flushing
     */
    stopSplunkAutoFlush() {
        if (this.splunkFlushInterval) {
            clearInterval(this.splunkFlushInterval);
            this.splunkFlushInterval = null;
            console.log('Stopped telemetry auto-flush');
        }
    }

    /**
     * Logs classification warning override
     * @param {Object} details - Override details
     */
    async logClassificationOverride(details) {
        await this.logEvent('classification_override', {
            classification_level: details.level,
            classification_text: details.text,
            subject: '[REDACTED]', // Don't log actual subject for classified content
            override_timestamp: new Date().toISOString(),
            compliance_note: 'User proceeded despite classification warning'
        }, 'Warning');
    }

    /**
     * Logs session telemetry
     * @param {Object} sessionData - Session information
     */
    async logSessionTelemetry(sessionData) {
        const telemetryData = {
            session_duration: sessionData.duration,
            emails_processed: sessionData.emailsProcessed || 0,
            responses_generated: sessionData.responsesGenerated || 0,
            model_service: sessionData.modelService || 'unknown',
            feature_usage: sessionData.featureUsage || {},
            error_count: sessionData.errorCount || 0
        };

        await this.logEvent('session_telemetry', telemetryData);
    }

    /**
     * Logs email processing metrics
     * @param {Object} metrics - Processing metrics
     */
    async logProcessingMetrics(metrics) {
        const metricsData = {
            model_service: metrics.modelService,
            model_name: metrics.modelName,
            email_length: metrics.emailLength,
            processing_time: metrics.processingTime,
            participant_count: metrics.participantCount,
            success: metrics.success,
            error_type: metrics.errorType || null
        };

        await this.logEvent('email_processed', metricsData);
    }

    /**
     * Logs security events
     * @param {string} eventType - Security event type
     * @param {Object} details - Event details
     */
    async logSecurityEvent(eventType, details) {
        const securityData = {
            security_event: eventType,
            risk_level: details.riskLevel || 'medium',
            action_taken: details.actionTaken,
            additional_info: details.info || 'No additional information'
        };

        await this.logEvent('security_event', securityData, 'Warning');
    }

    /**
     * Gets a session identifier
     * @returns {string} Session ID
     */
    getSessionId() {
        if (!this.sessionId) {
            this.sessionId = 'sess_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
        }
        return this.sessionId;
    }

    /**
     * Gets user identifier
     * @param {string} contextEmail - Optional email context to use instead of current user
     * @returns {string} User ID
     */
    getUserId(contextEmail = null) {
        try {
            // Use provided context email first (e.g., current email recipient)
            if (contextEmail) {
                return contextEmail;
            }
            
            // Use Office context if available
            if (typeof Office !== 'undefined' && Office.context?.mailbox?.userProfile?.emailAddress) {
                const email = Office.context.mailbox.userProfile.emailAddress;
                return email;
            }
        } catch (error) {
            console.warn('Could not get user ID from Office context');
        }
        
        return 'user_unknown';
    }

    /**
     * Simple hash function for user ID anonymization
     * @param {string} str - String to hash
     * @returns {string} Hash value
     */
    simpleHash(str) {
        let hash = 0;
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32-bit integer
        }
        return Math.abs(hash).toString(36);
    }

    /**
     * Enables or disables logging
     * @param {boolean} enabled - Whether logging should be enabled
     */
    setLoggingEnabled(enabled) {
        this.isEnabled = enabled;
        
        if (!enabled) {
            this.logQueue = []; // Clear queue when disabling
        }
    }

    /**
     * Gets current logging status
     * @returns {Object} Logging status information
     */
    getStatus() {
        return {
            enabled: this.isEnabled,
            queueSize: this.logQueue.length,
            eventSource: this.eventSource,
            sessionId: this.getSessionId()
        };
    }

    /**
     * Forces immediate flush of telemetry queue
     */
    async forceFlush() {
        if (this.splunkQueue.length > 0) {
            const provider = this.telemetryConfig?.telemetry?.provider;
            if (provider === 'api_gateway') {
                await this.flushApiGatewayQueue();
            } else if (provider === 'splunk_hec') {
                await this.flushSplunkQueue();
            }
        }
    }
}

