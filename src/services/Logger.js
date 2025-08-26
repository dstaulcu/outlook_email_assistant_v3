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
    createLogEntry(eventType, data, contextEmail) {
        const timestamp = new Date().toISOString();
        const sessionId = this.getSessionId();

        // Always include consistent user context
        const userContext = this.getUserContext(contextEmail);

        return {
            eventType: eventType,
            timestamp: timestamp,
            source: this.eventSource,
            version: process.env.PACKAGE_VERSION || 'unknown',
            sessionId: sessionId,
            userContext: userContext,
            ...data
        };
    }

    /**
     * Get consistent user context for all events
     * @param {string} contextEmail - Optional email override
     * @returns {Object} User context object
     */
    getUserContext(contextEmail) {
        // If specific email provided, use it
        if (contextEmail) {
            return { email: contextEmail };
        }

        // Try to get user email from Office context
        try {
            if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox && Office.context.mailbox.userProfile) {
                const userProfile = Office.context.mailbox.userProfile;
                if (userProfile.emailAddress) {
                    return { email: userProfile.emailAddress };
                }
            }
        } catch (error) {
            console.debug('Could not retrieve Office user profile:', error);
        }

        // Fallback: Check if we have cached user context
        if (this.cachedUserContext) {
            return this.cachedUserContext;
        }

        // Return sanitized context for privacy
        return { email: 'unknown@domain.com' };
    }

    /**
     * Cache user context for consistent telemetry
     * @param {string} email - User email address
     */
    cacheUserContext(email) {
        if (email && email.includes('@')) {
            this.cachedUserContext = { email: email };
        }
    }

    /**
     * Log to Splunk HEC
     * @param {string} eventType - Event type
     * @param {Object} logEntry - Log entry
     * @param {string} level - Log level
     */
    async logToSplunk(eventType, logEntry, level) {
        const splunkData = {
            time: Math.floor(Date.now() / 1000),
            host: this.telemetryConfig?.environment?.host || 'localhost',
            source: this.eventSource.toLowerCase().replace(/\s+/g, '_'),
            sourcetype: 'json:outlook_email_assistant',
            event: logEntry
        };

        this.splunkQueue.push(splunkData);
        
        // Queue management - prevent unlimited growth
        if (this.splunkQueue.length > 1000) {
            console.warn('[Logger] Queue size exceeded 1000 events, removing oldest entries');
            this.splunkQueue = this.splunkQueue.slice(-500); // Keep most recent 500
        }
        
        console.debug('[Logger] Event queued for Splunk, queue size:', this.splunkQueue.length);
        
        // Auto-flush if queue is large or at regular intervals
        if (this.splunkQueue.length >= 10) {
            await this.flushSplunkQueue();
        }
    }

    /**
     * Log to API Gateway
     * @param {string} eventType - Event type
     * @param {Object} logEntry - Log entry
     * @param {string} level - Log level
     */
    async logToApiGateway(eventType, logEntry, level) {
        try {
            console.debug('Preparing API Gateway event:', logEntry);
            
            const apiGatewayData = {
                time: Math.floor(Date.now() / 1000),
                host: this.telemetryConfig?.environment?.host || 'localhost',
                source: this.eventSource.toLowerCase().replace(/\s+/g, '_'),
                sourcetype: 'json:outlook_email_assistant',
                event: logEntry
            };

            this.splunkQueue.push(apiGatewayData);
            
            // Queue management - prevent unlimited growth
            if (this.splunkQueue.length > 1000) {
                console.warn('[Logger] Queue size exceeded 1000 events, removing oldest entries');
                this.splunkQueue = this.splunkQueue.slice(-500);
            }
            
            console.debug('[Logger] Event queued for API Gateway, queue size:', this.splunkQueue.length);
            
            // Auto-flush if queue is large
            if (this.splunkQueue.length >= 10) {
                await this.flushApiGatewayQueue();
            }
        } catch (error) {
            console.error('Error adding event to API Gateway queue:', error);
        }
    }

    /**
     * Flush queue to Splunk HEC endpoint
     */
    async flushSplunkQueue() {
        if (this.splunkQueue.length === 0) return;

        const events = [...this.splunkQueue];
        this.splunkQueue = [];

        try {
            const splunkConfig = this.telemetryConfig.telemetry.splunk;

            console.debug(`[Logger] Flushing ${events.length} events to Splunk HEC`);

            const fetchOptions = {
                method: 'POST',
                headers: {
                    'Authorization': `Splunk ${splunkConfig.token}`,
                    'Content-Type': 'application/json'
                },
                body: events.map(event => JSON.stringify(event)).join('\n')
            };

            if (!splunkConfig.validateCertificate) {
                fetchOptions.rejectUnauthorized = false;
            }

            console.debug(`[Logger] Request headers:`, fetchOptions.headers);

            const response = await fetch(splunkConfig.hecEndpoint, fetchOptions);

            console.debug(`[Logger] Response status: ${response.status} ${response.statusText}`);

            if (response.ok) {
                console.info('Successfully sent events to Splunk HEC');
                this.splunkConnectionError = false;
                this.splunkRetryCount = 0;
            } else {
                console.error('Failed to send events to Splunk HEC:', response.status);
                
                if (events.length < 100) {
                    this.splunkQueue.unshift(...events);
                }
            }

        } catch (error) {
            console.debug(`[Logger] Splunk HEC error details:`, {
                message: error.message,
                name: error.name,
                stack: error.stack?.substring(0, 200)
            });
            
            if (error.message.includes('certificate') || error.message.includes('SSL') || error.message.includes('TLS')) {
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
            } else if (error.message.includes('fetch') || error.message.includes('NetworkError')) {
                if (!this.splunkConnectionError) {
                    console.warn('Splunk HEC connection unavailable, events will be queued');
                    this.splunkConnectionError = true;
                }
                
                if (events.length < 100) {
                    this.splunkQueue.unshift(...events);
                    this.splunkRetryCount++;
                    
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
     * Flush queue to API Gateway endpoint with exponential backoff
     */
    async flushApiGatewayQueue() {
        if (this.splunkQueue.length === 0) return;

        const events = [...this.splunkQueue];
        this.splunkQueue = [];

        try {
            const apiGatewayConfig = this.telemetryConfig.telemetry.api_gateway;

            console.debug(`[Logger] Flushing ${events.length} events to API Gateway`);

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
                this.splunkConnectionError = false;
                this.splunkRetryCount = 0;
                const result = await response.json();
                console.debug('API Gateway response:', result);
            } else {
                // Only retry on transient errors (5xx, 429), not on 400/401/403
                const status = response.status;
                const isTransient = (status >= 500 && status < 600) || status === 429;
                if (isTransient) {
                    this.splunkRetryCount++;
                    const backoff = Math.min(1000 * Math.pow(2, this.splunkRetryCount - 1), 8000);
                    if (this.splunkRetryCount <= this.maxRetries) {
                        console.warn(`Transient error (${status}). Retrying in ${backoff}ms (attempt ${this.splunkRetryCount}/${this.maxRetries})`);
                        setTimeout(() => {
                            this.splunkQueue.unshift(...events);
                            this.flushApiGatewayQueue();
                        }, backoff);
                    } else {
                        console.error('Max API Gateway retry attempts reached, dropping events');
                        this.splunkRetryCount = 0;
                    }
                } else {
                    console.error(`Permanent error from API Gateway (${status}). Dropping events and not retrying.`);
                    this.splunkRetryCount = 0;
                }
            }
        } catch (error) {
            console.debug(`[Logger] API Gateway error details:`, {
                message: error.message,
                name: error.name,
                stack: error.stack?.substring(0, 200)
            });
            
            if (error.message.includes('fetch') || error.message.includes('NetworkError')) {
                if (!this.splunkConnectionError) {
                    console.warn('API Gateway connection unavailable, events will be queued');
                    this.splunkConnectionError = true;
                }
                this.splunkRetryCount++;
                const backoff = Math.min(1000 * Math.pow(2, this.splunkRetryCount - 1), 8000);
                if (this.splunkRetryCount <= this.maxRetries) {
                    setTimeout(() => {
                        this.splunkQueue.unshift(...events);
                        this.flushApiGatewayQueue();
                    }, backoff);
                } else {
                    console.error('Max API Gateway retry attempts reached (network error), dropping events');
                    this.splunkRetryCount = 0;
                }
            } else {
                console.error('Error flushing API Gateway queue:', error);
                this.splunkRetryCount = 0;
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
            console.debug('[Logger] Stopped auto-flush');
        }
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
        const telemetryData = {
            processing_time_ms: metrics.processingTime,
            email_length: metrics.emailLength,
            response_length: metrics.responseLength,
            model_used: metrics.model,
            classification: metrics.classification,
            tokens_used: metrics.tokensUsed || 0,
            error_occurred: metrics.errorOccurred || false
        };

        await this.logEvent('processing_metrics', telemetryData);
    }

    /**
     * Logs security events
     * @param {string} eventType - Type of security event
     * @param {Object} details - Security event details
     */
    async logSecurityEvent(eventType, details) {
        const securityData = {
            security_event_type: eventType,
            classification_level: details.classificationLevel,
            action_taken: details.actionTaken,
            user_override: details.userOverride || false,
            compliance_note: 'User proceeded despite classification warning'
        };

        await this.logEvent('security_event', securityData, 'Warning');
    }

    /**
     * Gets session ID for tracking
     * @returns {string} Session ID
     */
    getSessionId() {
        if (!this.sessionId) {
            this.sessionId = 'sess_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
        }
        return this.sessionId;
    }
}