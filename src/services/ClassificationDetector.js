/**
 * Classification Detector Service
 * Detects security classifications in email content
 */

export class ClassificationDetector {
    constructor() {
        // Classification levels and their numeric values for comparison
        this.classifications = {
            'UNCLASSIFIED': { level: 0, color: 'green' },
            'CONFIDENTIAL': { level: 1, color: 'yellow' },
            'SECRET': { level: 2, color: 'orange' },
            'TOP SECRET': { level: 3, color: 'red' },
            'TS': { level: 3, color: 'red' },
            'COSMIC TOP SECRET': { level: 4, color: 'red' },
            'CTS': { level: 4, color: 'red' }
        };

        // Common classification patterns
        this.patterns = [
            // Standard classification markings
            /^(UNCLASSIFIED|CONFIDENTIAL|SECRET|TOP SECRET|TS|COSMIC TOP SECRET|CTS)\s*$/gim,
            
            // Classification with additional markings
            /^(UNCLASSIFIED|CONFIDENTIAL|SECRET|TOP SECRET|TS)\/\/([A-Z\s\/]+)\s*$/gim,
            
            // Classification banners
            /^\s*(CLASSIFICATION:|CLASS:)\s*(UNCLASSIFIED|CONFIDENTIAL|SECRET|TOP SECRET|TS|COSMIC TOP SECRET|CTS)\s*$/gim,
            
            // Portion markings
            /\(([UCS]|CONFIDENTIAL|SECRET|TOP SECRET|TS)\)/gim
        ];

        // Lines to check (first N lines of email)
        this.linesToCheck = 5;
    }

    /**
     * Detects classification in email content
     * @param {string} emailBody - The email body text
     * @returns {Object} Classification detection result
     */
    detectClassification(emailBody) {
        if (!emailBody || typeof emailBody !== 'string') {
            return {
                detected: false,
                level: 0,
                text: 'UNCLASSIFIED',
                warning: false,
                details: 'No content to analyze'
            };
        }

        // Get first few lines for classification checking
        const lines = emailBody.split('\n').slice(0, this.linesToCheck);
        const headerText = lines.join('\n');

        let highestClassification = null;
        let detectedMarkings = [];

        // Check each pattern
        for (const pattern of this.patterns) {
            const matches = headerText.matchAll(pattern);
            
            for (const match of matches) {
                const classification = this.normalizeClassification(match[1] || match[2]);
                
                if (this.classifications[classification]) {
                    detectedMarkings.push({
                        text: match[0].trim(),
                        classification: classification,
                        level: this.classifications[classification].level,
                        line: this.findLineNumber(emailBody, match[0])
                    });

                    // Track highest classification found
                    if (!highestClassification || 
                        this.classifications[classification].level > highestClassification.level) {
                        highestClassification = {
                            text: classification,
                            level: this.classifications[classification].level,
                            color: this.classifications[classification].color
                        };
                    }
                }
            }
        }

        // If no classification found, assume unclassified
        if (!highestClassification) {
            return {
                detected: false,
                level: 0,
                text: 'UNCLASSIFIED',
                warning: false,
                details: 'No classification markings detected',
                markings: []
            };
        }

        // Determine if warning should be shown (SECRET and above)
        const shouldWarn = highestClassification.level >= 2;

        return {
            detected: true,
            level: highestClassification.level,
            text: highestClassification.text,
            warning: shouldWarn,
            color: highestClassification.color,
            details: this.getClassificationMessage(highestClassification),
            markings: detectedMarkings
        };
    }

    /**
     * Normalizes classification text to standard format
     * @param {string} text - Raw classification text
     * @returns {string} Normalized classification
     */
    normalizeClassification(text) {
        if (!text) return 'UNCLASSIFIED';
        
        const normalized = text.trim().toUpperCase();
        
        // Handle common abbreviations
        if (normalized === 'U') return 'UNCLASSIFIED';
        if (normalized === 'C') return 'CONFIDENTIAL';
        if (normalized === 'S') return 'SECRET';
        if (normalized === 'TS') return 'TOP SECRET';
        if (normalized === 'CTS') return 'COSMIC TOP SECRET';
        
        return normalized;
    }

    /**
     * Finds the line number where a match was found
     * @param {string} text - Full text
     * @param {string} match - Matched text
     * @returns {number} Line number (1-based)
     */
    findLineNumber(text, match) {
        const lines = text.split('\n');
        for (let i = 0; i < Math.min(lines.length, this.linesToCheck); i++) {
            if (lines[i].includes(match.trim())) {
                return i + 1;
            }
        }
        return 1;
    }

    /**
     * Gets appropriate message for classification level
     * @param {Object} classification - Classification object
     * @returns {string} User-friendly message
     */
    getClassificationMessage(classification) {
        switch (classification.level) {
            case 0:
                return 'This content is unclassified and safe to process.';
            case 1:
                return 'This content is marked CONFIDENTIAL. Exercise caution when sharing with external systems.';
            case 2:
                return 'This content is marked SECRET. Sharing with external AI services may violate security policies.';
            case 3:
                return 'This content is marked TOP SECRET. External processing is strictly prohibited by security policy.';
            case 4:
                return 'This content has the highest classification level. External processing is absolutely forbidden.';
            default:
                return 'Classification level could not be determined. Proceed with caution.';
        }
    }

    /**
     * Validates if processing should be allowed for a classification level
     * @param {number} level - Classification level
     * @param {boolean} userOverride - Whether user has overridden the warning
     * @returns {Object} Validation result
     */
    validateProcessing(level, userOverride = false) {
        if (level < 2) {
            return {
                allowed: true,
                requiresWarning: false,
                message: 'Processing allowed for this classification level.'
            };
        }

        if (level >= 2 && !userOverride) {
            return {
                allowed: false,
                requiresWarning: true,
                message: 'Processing blocked due to classification restrictions. User override required.'
            };
        }

        if (level >= 2 && userOverride) {
            return {
                allowed: true,
                requiresWarning: true,
                requiresLogging: true,
                message: 'Processing allowed by user override. This action will be logged.'
            };
        }

        return {
            allowed: false,
            requiresWarning: true,
            message: 'Processing not allowed for this classification level.'
        };
    }

    /**
     * Generates a detailed classification report
     * @param {string} emailBody - Email content
     * @returns {Object} Detailed report
     */
    generateReport(emailBody) {
        const detection = this.detectClassification(emailBody);
        const validation = this.validateProcessing(detection.level);

        return {
            ...detection,
            validation,
            timestamp: new Date().toISOString(),
            analyzer: 'PromptEmail Classification Detector v1.0'
        };
    }

    /**
     * Checks if email content should trigger automatic blocking
     * @param {string} emailBody - Email content
     * @returns {boolean} True if should be blocked
     */
    shouldBlock(emailBody) {
        const detection = this.detectClassification(emailBody);
        return detection.level >= 3; // TOP SECRET and above
    }

    /**
     * Gets CSS class for classification level styling
     * @param {number} level - Classification level
     * @returns {string} CSS class name
     */
    getClassificationStyle(level) {
        switch (level) {
            case 0: return 'classification-unclassified';
            case 1: return 'classification-confidential';
            case 2: return 'classification-secret';
            case 3: 
            case 4: return 'classification-top-secret';
            default: return 'classification-unknown';
        }
    }

    /**
     * Sanitizes content for logging (removes actual classified content)
     * @param {string} content - Original content
     * @param {Object} detection - Detection results
     * @returns {string} Sanitized content safe for logging
     */
    sanitizeForLogging(content, detection) {
        if (!detection.detected || detection.level === 0) {
            // For unclassified, return first 100 characters
            return content.substring(0, 100) + (content.length > 100 ? '...' : '');
        }

        // For classified content, return only metadata
        return `[CLASSIFIED CONTENT - ${detection.text} - ${content.length} chars]`;
    }
}
