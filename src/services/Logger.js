/**
 * Logger Service
 * Handles logging events to Windows Application Log using PowerShell
 */

export class Logger {
    constructor() {
        this.eventSource = 'PromptEmail';
        this.logName = 'Application';
        this.isEnabled = true;
        this.logQueue = [];
        this.maxQueueSize = 100;
    }

    /**
     * Logs an event to the Windows Application Log
     * @param {string} eventType - Type of event (e.g., 'session_start', 'email_analyzed')
     * @param {Object} data - Event data object
     * @param {string} level - Log level ('Information', 'Warning', 'Error')
     */
    async logEvent(eventType, data = {}, level = 'Information') {
        if (!this.isEnabled) {
            console.log('Logging disabled:', eventType, data);
            return;
        }

        try {
            const logEntry = this.createLogEntry(eventType, data);
            
            // Add to queue for batch processing
            this.logQueue.push({
                entry: logEntry,
                level: level,
                timestamp: new Date().toISOString()
            });

            // Process queue if it's getting full
            if (this.logQueue.length >= this.maxQueueSize) {
                await this.flushQueue();
            }

            // For console development
            console.log(`[${level}] ${eventType}:`, logEntry);

        } catch (error) {
            console.error('Failed to log event:', error);
        }
    }

    /**
     * Creates a standardized log entry
     * @param {string} eventType - Event type
     * @param {Object} data - Event data
     * @returns {Object} Log entry object
     */
    createLogEntry(eventType, data) {
        const baseEntry = {
            eventType: eventType,
            timestamp: new Date().toISOString(),
            source: this.eventSource,
            version: '1.0.0',
            sessionId: this.getSessionId(),
            userId: this.getUserId()
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
     * Writes log entry to Windows Application Log via PowerShell
     * @param {Object} logEntry - Log entry to write
     * @param {string} level - Log level
     */
    async writeToWindowsLog(logEntry, level) {
        try {
            const jsonMessage = JSON.stringify(logEntry);
            
            // PowerShell command to write to Application Log
            const psCommand = `
                $eventSource = "${this.eventSource}"
                $logName = "${this.logName}"
                $eventLevel = "${level}"
                $message = @'
${jsonMessage}
'@

                # Check if event source exists, create if not
                if (-not [System.Diagnostics.EventLog]::SourceExists($eventSource)) {
                    try {
                        [System.Diagnostics.EventLog]::CreateEventSource($eventSource, $logName)
                        Start-Sleep -Seconds 1
                    } catch {
                        Write-Host "Warning: Could not create event source. May need administrator privileges."
                        return
                    }
                }

                # Write event
                try {
                    $eventType = switch($eventLevel) {
                        "Information" { [System.Diagnostics.EventLogEntryType]::Information }
                        "Warning" { [System.Diagnostics.EventLogEntryType]::Warning }
                        "Error" { [System.Diagnostics.EventLogEntryType]::Error }
                        default { [System.Diagnostics.EventLogEntryType]::Information }
                    }
                    
                    [System.Diagnostics.EventLog]::WriteEntry($eventSource, $message, $eventType, 1001)
                    Write-Host "Event logged successfully"
                } catch {
                    Write-Host "Error writing to event log: $_"
                }
            `;

            // Execute PowerShell command (this would typically be done via a bridge in a real implementation)
            if (typeof window !== 'undefined' && window.chrome?.webview) {
                // WebView2 bridge for PowerShell execution
                await window.chrome.webview.postMessage({
                    type: 'executePowerShell',
                    script: psCommand
                });
            } else {
                // Fallback: log to console in development
                console.log('PowerShell Log Entry:', jsonMessage);
            }

        } catch (error) {
            console.error('Failed to write to Windows log:', error);
        }
    }

    /**
     * Flushes the log queue by writing all entries
     */
    async flushQueue() {
        const entriesToProcess = [...this.logQueue];
        this.logQueue = [];

        for (const item of entriesToProcess) {
            await this.writeToWindowsLog(item.entry, item.level);
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
     * Gets user identifier (anonymized)
     * @returns {string} User ID
     */
    getUserId() {
        try {
            // Use Office context if available
            if (typeof Office !== 'undefined' && Office.context?.mailbox?.userProfile?.emailAddress) {
                const email = Office.context.mailbox.userProfile.emailAddress;
                // Hash the email for privacy
                return 'user_' + this.simpleHash(email);
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
     * Forces immediate flush of all queued log entries
     */
    async forceFlush() {
        if (this.logQueue.length > 0) {
            await this.flushQueue();
        }
    }
}
