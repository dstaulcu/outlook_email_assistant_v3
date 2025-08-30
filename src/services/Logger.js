/**
 * Logger Service
 * Handles console logging (always enabled) and API Gateway telemetry (configurable)
 */

export class Logger {
    constructor() {
        this.eventSource = 'PromptEmail';
        this.isEnabled = true;
        this.telemetryConfig = null;
        this.apiGatewayQueue = [];
        this.apiGatewayFlushInterval = null;
        this.apiGatewayRetryCount = 0;
        this.maxRetries = 3;
        this.apiGatewayConnectionError = false;
        
        // Initialize telemetry configuration
        this.initializeTelemetryConfig();
    }

    /**
     * Initialize telemetry configuration
     */
    async initializeTelemetryConfig() {
        try {
            console.debug('[VERBOSE] - Loading telemetry configuration...');
            const response = await fetch('/config/telemetry.json');
            if (response.ok) {
                this.telemetryConfig = await response.json();
                console.info('[INFO] - Telemetry configuration loaded:', this.telemetryConfig);
            } else {
                console.warn('[WARN] - Could not load telemetry configuration, using defaults');
                this.telemetryConfig = this.getDefaultTelemetryConfig();
            }
        } catch (error) {
            console.error('[ERROR] - Failed to load telemetry configuration:', error);
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
     * Logs an event to console (always) and API Gateway (if enabled)
     * Automatically includes Office diagnostic values (host, platform, version) in all telemetry events
     * @param {string} eventType - Type of event (e.g., 'session_start', 'email_analyzed')
     * @param {Object} data - Event data object
     * @param {string} level - Log level ('Information', 'Warning', 'Error')
     * @param {string} contextEmail - Optional email address for better user context
     */
    async logEvent(eventType, data = {}, level = 'Information', contextEmail = null) {
        if (!this.isEnabled) {
            console.debug('[VERBOSE] - Logging disabled:', eventType, data);
            return;
        }

        try {
            const logEntry = this.createLogEntry(eventType, data, contextEmail);
            
            // Always log to console for development/debugging
            console.debug(`[VERBOSE] - ${level} ${eventType}:`, logEntry);

            // Add to telemetry queue if enabled
            if (this.telemetryConfig?.telemetry?.enabled) {
                if (this.telemetryConfig.telemetry.provider === 'api_gateway') {
                    await this.logToApiGateway(eventType, logEntry, level);
                }
            }

        } catch (error) {
            console.error('[ERROR] - Failed to log event:', error);
        }
    }

    /**
     * Creates a standardized log entry with flattened structure
     * @param {string} eventType - Event type
     * @param {Object} data - Event data
     * @param {string} contextEmail - Optional email context for user identification (deprecated - now in flattened userProfile fields)
     * @returns {Object} Log entry object
     */
    createLogEntry(eventType, data, contextEmail) {
        const timestamp = new Date().toISOString();
        const sessionId = this.getSessionId();
        
        // Get Office diagnostic values (includes user profile information)
        const officeDiagnostics = this.getOfficeDiagnostics();
        
        // Get client context information
        const clientContext = this.getClientContext();

        // Flatten Office diagnostics
        const flattenedEntry = {
            eventType: eventType,
            timestamp: timestamp,
            source: this.eventSource.toLowerCase().replace(/\s+/g, '_'),
            version: process.env.PACKAGE_VERSION || 'unknown',
            sessionId: sessionId,
            // Flattened Office diagnostic fields
            office_host: officeDiagnostics.host,
            office_platform: officeDiagnostics.platform,
            office_version: officeDiagnostics.version,
            ...data
        };

        // Add optional OWA view if available
        if (officeDiagnostics.owaView) {
            flattenedEntry.office_owa_view = officeDiagnostics.owaView;
        }

        // Flatten userProfile fields
        if (officeDiagnostics.userProfile) {
            if (officeDiagnostics.userProfile.displayName) {
                flattenedEntry.userProfile_displayName = officeDiagnostics.userProfile.displayName;
            }
            if (officeDiagnostics.userProfile.emailAddress) {
                flattenedEntry.userProfile_emailAddress = officeDiagnostics.userProfile.emailAddress;
            }
            if (officeDiagnostics.userProfile.timeZone) {
                flattenedEntry.userProfile_timeZone = officeDiagnostics.userProfile.timeZone;
            }
            if (officeDiagnostics.userProfile.accountType) {
                flattenedEntry.userProfile_accountType = officeDiagnostics.userProfile.accountType;
            }
        }

        // Add flattened client context fields
        //if (clientContext.browser_name) flattenedEntry.client_browser_name = clientContext.browser_name;
        //if (clientContext.browser_version) flattenedEntry.client_browser_version = clientContext.browser_version;
        //if (clientContext.navigator_platform) flattenedEntry.client_platform = clientContext.navigator_platform;
        //if (clientContext.language) flattenedEntry.client_language = clientContext.language;
        //if (clientContext.timezone) flattenedEntry.client_timezone = clientContext.timezone;
        if (clientContext.screen_resolution) flattenedEntry.client_screen_resolution = clientContext.screen_resolution;
        //if (clientContext.viewport_size) flattenedEntry.client_viewport_size = clientContext.viewport_size;
        //if (clientContext.connection_type) flattenedEntry.client_connection_type = clientContext.connection_type;
        //if (clientContext.cpu_cores) flattenedEntry.client_cpu_cores = clientContext.cpu_cores;
        //if (clientContext.device_memory_gb) flattenedEntry.client_device_memory_gb = clientContext.device_memory_gb;
        
        // Add performance metrics if available
        //if (clientContext.js_heap_size_mb) flattenedEntry.client_js_heap_size_mb = clientContext.js_heap_size_mb;
        if (clientContext.connection_rtt_ms) flattenedEntry.client_connection_rtt_ms = clientContext.connection_rtt_ms;
        if (clientContext.connection_downlink_mbps) flattenedEntry.client_connection_downlink_mbps = clientContext.connection_downlink_mbps;

        // Add error if present
        if (officeDiagnostics.error) {
            flattenedEntry.office_error = officeDiagnostics.error;
        }
        if (clientContext.error) {
            flattenedEntry.client_error = clientContext.error;
        }

        return flattenedEntry;
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
            console.debug('[VERBOSE] - Could not retrieve Office user profile:', error);
        }

        // Fallback: Check if we have cached user context
        if (this.cachedUserContext) {
            return this.cachedUserContext;
        }

        // Return sanitized context for privacy
        return { email: 'unknown@domain.com' };
    }

    /**
     * Get client identification information
     * @returns {Object} Client context including browser, network info
     */
    getClientContext() {
        const clientContext = {};

        try {
            // Browser/User Agent information
            if (typeof navigator !== 'undefined') {
                // Basic browser info
                if (navigator.userAgent) {
                    clientContext.user_agent = navigator.userAgent;
                    
                    // Parse browser name and version from user agent
                    const browserInfo = this.parseBrowserInfo(navigator.userAgent);
                    clientContext.browser_name = browserInfo.name;
                    clientContext.browser_version = browserInfo.version;
                }
                
                // Language preferences
                if (navigator.language) {
                    clientContext.language = navigator.language;
                }
                if (navigator.languages && navigator.languages.length > 0) {
                    clientContext.languages = navigator.languages.slice(0, 3).join(','); // First 3 languages
                }
                
                // Platform information
                if (navigator.platform) {
                    clientContext.navigator_platform = navigator.platform;
                }
                
                // Hardware concurrency (CPU cores)
                if (navigator.hardwareConcurrency) {
                    clientContext.cpu_cores = navigator.hardwareConcurrency;
                }
                
                // Memory information (if available)
                if (navigator.deviceMemory) {
                    clientContext.device_memory_gb = navigator.deviceMemory;
                }
                
                // Connection information (if available)
                if (navigator.connection) {
                    const conn = navigator.connection;
                    if (conn.effectiveType) clientContext.connection_type = conn.effectiveType;
                    if (conn.downlink) clientContext.connection_downlink_mbps = conn.downlink;
                    if (conn.rtt) clientContext.connection_rtt_ms = conn.rtt;
                    if (conn.saveData !== undefined) clientContext.connection_save_data = conn.saveData;
                }
            }
            
            // Screen information
            if (typeof screen !== 'undefined') {
                if (screen.width && screen.height) {
                    clientContext.screen_resolution = `${screen.width}x${screen.height}`;
                }
                if (screen.colorDepth) {
                    clientContext.screen_color_depth = screen.colorDepth;
                }
            }
            
            // Window/viewport information  
            if (typeof window !== 'undefined') {
                if (window.innerWidth && window.innerHeight) {
                    clientContext.viewport_size = `${window.innerWidth}x${window.innerHeight}`;
                }
                
                // Timezone information
                try {
                    if (Intl && Intl.DateTimeFormat) {
                        clientContext.timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
                    }
                } catch (e) {
                    // Ignore timezone detection errors
                }
                
                // Performance information
                if (window.performance && window.performance.memory) {
                    const memory = window.performance.memory;
                    clientContext.js_heap_size_mb = Math.round(memory.usedJSHeapSize / 1024 / 1024);
                    clientContext.js_heap_limit_mb = Math.round(memory.jsHeapSizeLimit / 1024 / 1024);
                }
            }

            // Note: Client IP address is NOT accessible from browser JavaScript for security reasons
            // IP address would need to be obtained server-side (e.g., from API Gateway logs)
            
        } catch (error) {
            console.debug('[VERBOSE] - Error getting client context:', error);
            clientContext.error = error.message;
        }

        return clientContext;
    }

    /**
     * Parse browser information from user agent string
     * @param {string} userAgent - User agent string
     * @returns {Object} Browser name and version
     */
    parseBrowserInfo(userAgent) {
        const browsers = [
            { name: 'Chrome', regex: /Chrome\/(\d+\.\d+)/ },
            { name: 'Firefox', regex: /Firefox\/(\d+\.\d+)/ },
            { name: 'Safari', regex: /Version\/(\d+\.\d+).*Safari/ },
            { name: 'Edge', regex: /Edg\/(\d+\.\d+)/ },
            { name: 'IE', regex: /Trident.*rv:(\d+\.\d+)/ }
        ];
        
        for (const browser of browsers) {
            const match = userAgent.match(browser.regex);
            if (match) {
                return { name: browser.name, version: match[1] };
            }
        }
        
        return { name: 'Unknown', version: 'Unknown' };
    }

    /**
     * Get deployment environment information 
     * @returns {Object} Environment context including host domain
     */
    getEnvironmentContext() {
        const envContext = {
            environment: 'unknown',
            host: 'localhost'
        };

        try {
            // Try to determine environment from current location/domain
            if (typeof window !== 'undefined' && window.location) {
                const currentHost = window.location.hostname.toLowerCase();
                envContext.host = currentHost;
                
                // Check if we're running from S3 and parse environment
                if (currentHost.includes('s3') && currentHost.includes('amazonaws.com')) {
                    if (currentHost.includes('-dev.s3')) {
                        envContext.environment = 'Dev';
                    } else if (currentHost.includes('-test.s3')) {
                        envContext.environment = 'Test';
                    } else if (currentHost.includes('-prod.s3')) {
                        envContext.environment = 'Prod';
                    }
                } else {
                    // Local development or other domain
                    if (currentHost === 'localhost' || currentHost.startsWith('127.0.0.1')) {
                        envContext.environment = 'Local';
                    }
                }
            }

        } catch (error) {
            console.debug('[VERBOSE] - Error getting environment context:', error);
            envContext.error = error.message;
        }

        return envContext;
    }

    /**
     * Get Office diagnostic information for telemetry
     * @returns {Object} Office diagnostic context (host, platform, version)
     */
    getOfficeDiagnostics() {
        const diagnostics = {
            host: 'unknown',
            platform: 'unknown',
            version: 'unknown'
        };

        try {
            // Check if Office context is available
            if (typeof Office !== 'undefined' && Office.context) {
                // Get Office host application
                if (Office.context.host) {
                    diagnostics.host = Office.context.host;
                } else {
                    // Fallback: determine host from context
                    if (Office.context.mailbox) {
                        diagnostics.host = 'Outlook';
                    } else if (Office.context.document) {
                        diagnostics.host = 'Word/Excel/PowerPoint';
                    }
                }

                // Get platform information
                if (Office.context.platform) {
                    switch (Office.context.platform) {
                        case Office.PlatformType.PC:
                            diagnostics.platform = 'Windows';
                            break;
                        case Office.PlatformType.Mac:
                            diagnostics.platform = 'Mac';
                            break;
                        case Office.PlatformType.OfficeOnline:
                            diagnostics.platform = 'Web';
                            break;
                        case Office.PlatformType.Universal:
                            diagnostics.platform = 'Universal';
                            break;
                        case Office.PlatformType.iOS:
                            diagnostics.platform = 'iOS';
                            break;
                        case Office.PlatformType.Android:
                            diagnostics.platform = 'Android';
                            break;
                        default:
                            diagnostics.platform = 'Other';
                    }
                }

                // Get Office version
                if (Office.context.diagnostics) {
                    diagnostics.version = Office.context.diagnostics.version;
                    
                    // Add additional diagnostic info if available
                    if (Office.context.diagnostics.platform) {
                        diagnostics.platform = Office.context.diagnostics.platform;
                    }
                    if (Office.context.diagnostics.host) {
                        diagnostics.host = Office.context.diagnostics.host;
                    }
                }

                // Try to get more specific Outlook version information
                if (Office.context.mailbox && Office.context.mailbox.diagnostics) {
                    const mailboxDiagnostics = Office.context.mailbox.diagnostics;
                    if (mailboxDiagnostics.hostVersion) {
                        diagnostics.version = mailboxDiagnostics.hostVersion;
                    }
                    if (mailboxDiagnostics.OWAView) {
                        diagnostics.owaView = mailboxDiagnostics.OWAView;
                    }
                }

                // Add user profile information from Office.context.mailbox.userProfile
                if (Office.context.mailbox && Office.context.mailbox.userProfile) {
                    const userProfile = Office.context.mailbox.userProfile;
                    diagnostics.userProfile = {};
                    
                    if (userProfile.displayName) {
                        diagnostics.userProfile.displayName = userProfile.displayName;
                    }
                    if (userProfile.emailAddress) {
                        diagnostics.userProfile.emailAddress = userProfile.emailAddress;
                    }
                    if (userProfile.timeZone) {
                        diagnostics.userProfile.timeZone = userProfile.timeZone;
                    }
                    if (userProfile.accountType) {
                        diagnostics.userProfile.accountType = userProfile.accountType;
                    }
                }

                // Add Office requirement set information (optional - can be enabled if needed)
                // if (Office.context.requirements) {
                //     diagnostics.requirementSets = {};
                //     
                //     // Check common requirement sets
                //     const commonSets = [
                //         'Mailbox', 'IdentityAPI', 'CustomFunctionsRuntime',
                //         'DialogAPI', 'DocumentAPI', 'ExcelAPI', 'WordAPI',
                //         'PowerPointAPI', 'OneNoteAPI', 'OutlookAPI',
                //         'Ribbon', 'SharedRuntime', 'Telemetry'
                //     ];
                //     
                //     commonSets.forEach(setName => {
                //         try {
                //             if (Office.context.requirements.isSetSupported(setName)) {
                //                 diagnostics.requirementSets[setName] = 'supported';
                //             }
                //         } catch (e) {
                //             // Requirement set not available
                //         }
                //     });
                // }

            } else {
                // Office context not available - might be running outside Office
                diagnostics.host = 'Non-Office';
                diagnostics.platform = this.getBrowserPlatform();
                diagnostics.version = 'N/A';
            }

        } catch (error) {
            console.debug('[VERBOSE] - Error getting Office diagnostics:', error);
            diagnostics.error = error.message;
        }

        return diagnostics;
    }

    /**
     * Get browser platform information as fallback
     * @returns {string} Platform string
     */
    getBrowserPlatform() {
        if (typeof navigator === 'undefined') return 'unknown';
        
        const userAgent = navigator.userAgent;
        if (userAgent.includes('Windows')) return 'Windows';
        if (userAgent.includes('Mac')) return 'Mac';
        if (userAgent.includes('Linux')) return 'Linux';
        if (userAgent.includes('iOS')) return 'iOS';
        if (userAgent.includes('Android')) return 'Android';
        
        return 'Web';
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
     * Log to API Gateway
     * @param {string} eventType - Event type
     * @param {Object} logEntry - Log entry
     * @param {string} level - Log level
     */
    async logToApiGateway(eventType, logEntry, level) {
        try {
            console.debug('[VERBOSE] - Preparing API Gateway event:', logEntry);
            
            const environmentContext = this.getEnvironmentContext();
            
            // Flatten environment context into the log entry
            const flattenedLogEntry = {
                ...logEntry,
                // Add flattened environment fields
                environment_type: environmentContext.environment,
                environment_host: environmentContext.host
            };

            // Add environment error if present
            if (environmentContext.error) {
                flattenedLogEntry.environment_error = environmentContext.error;
            }
            
            const apiGatewayData = {
                time: Math.floor(Date.now() / 1000),
                // Use environment host for better filtering
                host: environmentContext.host,
                source: this.eventSource.toLowerCase().replace(/\s+/g, '_'),
                sourcetype: 'json:outlook_email_assistant',
                event: flattenedLogEntry
            };

            this.apiGatewayQueue.push(apiGatewayData);
            
            // Queue management - prevent unlimited growth
            if (this.apiGatewayQueue.length > 1000) {
                console.warn('[WARN] - Logger queue size exceeded 1000 events, removing oldest entries');
                this.apiGatewayQueue = this.apiGatewayQueue.slice(-500);
            }
            
            console.debug('[VERBOSE] Event queued for API Gateway, queue size:', this.apiGatewayQueue.length);
            
            // Auto-flush if queue is large
            if (this.apiGatewayQueue.length >= 10) {
                await this.flushApiGatewayQueue();
            }
        } catch (error) {
            console.error('[ERROR] - Error adding event to API Gateway queue:', error);
        }
    }

    /**
     * Flush queue to API Gateway endpoint with exponential backoff
     */
    async flushApiGatewayQueue() {
        if (this.apiGatewayQueue.length === 0) return;

        const events = [...this.apiGatewayQueue];
        this.apiGatewayQueue = [];

        try {
            const apiGatewayConfig = this.telemetryConfig.telemetry.api_gateway;

            console.debug(`[VERBOSE] - Flushing ${events.length} events to API Gateway`);

            const fetchOptions = {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: events.map(event => JSON.stringify(event)).join('\n')
            };

            console.debug(`[VERBOSE] - Attempting to send to: ${apiGatewayConfig.endpoint}`);
            
            const response = await fetch(apiGatewayConfig.endpoint, fetchOptions);

            console.debug(`[VERBOSE] - Response status: ${response.status} ${response.statusText}`);

            if (response.ok) {
                console.info('[INFO] - Successfully sent events to API Gateway');
                this.apiGatewayConnectionError = false;
                this.apiGatewayRetryCount = 0;
                const result = await response.json();
                console.debug('[VERBOSE] - API Gateway response:', result);
            } else {
                // Only retry on transient errors (5xx, 429), not on 400/401/403
                const status = response.status;
                const isTransient = (status >= 500 && status < 600) || status === 429;
                if (isTransient) {
                    this.apiGatewayRetryCount++;
                    const backoff = Math.min(1000 * Math.pow(2, this.apiGatewayRetryCount - 1), 8000);
                    if (this.apiGatewayRetryCount <= this.maxRetries) {
                        console.warn(`[WARN] - Transient error (${status}). Retrying in ${backoff}ms (attempt ${this.apiGatewayRetryCount}/${this.maxRetries})`);
                        setTimeout(() => {
                            this.apiGatewayQueue.unshift(...events);
                            this.flushApiGatewayQueue();
                        }, backoff);
                    } else {
                        console.error('[ERROR] - Max API Gateway retry attempts reached, dropping events');
                        this.apiGatewayRetryCount = 0;
                    }
                } else {
                    console.error(`[ERROR] - Permanent error from API Gateway (${status}). Dropping events and not retrying.`);
                    this.apiGatewayRetryCount = 0;
                }
            }
        } catch (error) {
            console.debug(`[VERBOSE] - Logger API Gateway error details:`, {
                message: error.message,
                name: error.name,
                stack: error.stack?.substring(0, 200)
            });
            
            if (error.message.includes('fetch') || error.message.includes('NetworkError')) {
                if (!this.apiGatewayConnectionError) {
                    console.warn('[WARN] - API Gateway connection unavailable, events will be queued');
                    this.apiGatewayConnectionError = true;
                }
                this.apiGatewayRetryCount++;
                const backoff = Math.min(1000 * Math.pow(2, this.apiGatewayRetryCount - 1), 8000);
                if (this.apiGatewayRetryCount <= this.maxRetries) {
                    setTimeout(() => {
                        this.apiGatewayQueue.unshift(...events);
                        this.flushApiGatewayQueue();
                    }, backoff);
                } else {
                    console.error('[ERROR] - Max API Gateway retry attempts reached (network error), dropping events');
                    this.apiGatewayRetryCount = 0;
                }
            } else {
                console.error('[ERROR] - Error flushing API Gateway queue:', error);
                this.apiGatewayRetryCount = 0;
            }
        }
    }

    /**
     * Start automatic queue flushing for API Gateway
     */
    startApiGatewayAutoFlush() {
        if (this.apiGatewayFlushInterval) return;

        const flushInterval = this.telemetryConfig?.telemetry?.api_gateway?.flushInterval || 60000;

        this.apiGatewayFlushInterval = setInterval(async () => {
            if (this.apiGatewayQueue.length > 0) {
                await this.flushApiGatewayQueue();
            }
        }, flushInterval);

        console.info(`[INFO] - Started API Gateway logging auto-flush with ${flushInterval}ms interval`);
    }

    /**
     * Stop automatic queue flushing
     */
    stopApiGatewayAutoFlush() {
        if (this.apiGatewayFlushInterval) {
            clearInterval(this.apiGatewayFlushInterval);
            this.apiGatewayFlushInterval = null;
            console.debug('[VERBOSE] - Stopped logging auto-flush');
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
            tokens_used: metrics.tokensUsed || 0,
            error_occurred: metrics.errorOccurred || false
        };

        await this.logEvent('processing_metrics', telemetryData);
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