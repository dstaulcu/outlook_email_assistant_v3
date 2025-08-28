/**
 * Classification Detector Service
 * Detects security classifications in email content
 */

export class ClassificationDetector {
    constructor() {
        // Classification levels - keeping simple string-based approach
        this.classifications = {
            'UNCLASSIFIED': { color: 'green', restricted: false },
            'CONFIDENTIAL': { color: 'yellow', restricted: false }, 
            'SECRET': { color: 'orange', restricted: true },
            'TOP SECRET': { color: 'red', restricted: true },
            'TS': { color: 'red', restricted: true },
            'COSMIC TOP SECRET': { color: 'red', restricted: true },
            'CTS': { color: 'red', restricted: true }
        };

        // Common classification patterns - flexible to handle real-world formatting
        this.patterns = [
            // Classification at start of line with optional leading whitespace and trailing content
            /^\s*(UNCLASSIFIED|CONFIDENTIAL|(?:TOP\s+)?SECRET|TS|COSMIC\s+TOP\s+SECRET|CTS)(?:\s|$|[:\-])/gim,
            
            // Classification banners with various formats
            /^\s*(CLASSIFICATION:|CLASS:)\s*(UNCLASSIFIED|CONFIDENTIAL|(?:TOP\s+)?SECRET|TS|COSMIC\s+TOP\s+SECRET|CTS)/gim,
            
            // Classification with additional markings (e.g., "SECRET//NOFORN")
            /^\s*(UNCLASSIFIED|CONFIDENTIAL|(?:TOP\s+)?SECRET|TS)\/\/([A-Z\s\/]+)/gim,
            
            // Portion markings within content (only in parentheses)
            /\(([UCS]|CONFIDENTIAL|(?:TOP\s+)?SECRET|TS)\)/gim
        ];

        // Lines to check - classification should be in first 2 lines for structured blocks
        this.linesToCheck = 2;
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
                text: 'UNCLASSIFIED',
                restricted: false,
                warning: false,
                details: 'No content to analyze'
            };
        }

        // Debug: Log first few lines of email to understand what's being classified
        const firstLines = emailBody.split('\n').slice(0, 5).join('\n');
        // Only log if we detect something to avoid spam
        
        // First, try to parse the structured classification block
        const structuredResult = this.parseClassificationBlock(emailBody);
        if (structuredResult.detected) {
            console.debug('[ClassificationDetector] Classification found:', structuredResult.text);
            return structuredResult;
        }

        // Fallback to pattern-based detection
        const patternResult = this.detectClassificationByPatterns(emailBody);
        if (patternResult.detected) {
            console.debug('[ClassificationDetector] Pattern detection found:', patternResult.text);
        }
        return patternResult;
    }

    /**
     * Parses structured classification block from internal system
     * Classification should ONLY appear on the first readable line of the email
     * @param {string} emailBody - The email body text
     * @returns {Object} Classification detection result
     */
    parseClassificationBlock(emailBody) {
        // Look for classification ONLY in the first readable line
        const lines = emailBody.split('\n');
        
        // Find the first non-empty, non-whitespace line
        let firstReadableLine = null;
        let lineNumber = -1;
        
        for (let i = 0; i < lines.length && i < 5; i++) { // Don't look beyond first 5 lines
            const line = lines[i].trim();
            if (line) {
                firstReadableLine = line;
                lineNumber = i + 1;
                break;
            }
        }

        if (!firstReadableLine) {
            return { detected: false };
        }

        // Only log in debug mode to reduce noise
        // console.debug('[ClassificationDetector] First readable line:', firstReadableLine);

        // Check if the first readable line contains a classification level
        // Allow for leading/trailing whitespace and additional text after classification
        const upper = firstReadableLine.toUpperCase().trim();
        let detectedClassification = null;
        
        // Check for classification at the beginning of the line (with optional leading whitespace)
        for (const classLevel of Object.keys(this.classifications)) {
            // Create a regex that allows leading whitespace and optional text after classification
            const classPattern = new RegExp(`^\\s*${classLevel.replace(/\s+/g, '\\s+')}(?:\\s|$|[:\\-])`, 'i');
            if (classPattern.test(upper)) {
                detectedClassification = classLevel;
                break;
            }
        }

        // Also check for exact match after trimming (most common case)
        if (!detectedClassification) {
            for (const classLevel of Object.keys(this.classifications)) {
                if (upper === classLevel) {
                    detectedClassification = classLevel;
                    break;
                }
            }
        }

        if (!detectedClassification) {
            // Only log in debug scenarios - commented to reduce noise
            // console.debug('[ClassificationDetector] No classification found in first line');
            return { detected: false };
        }

        const classification = this.normalizeClassification(detectedClassification);
        const classificationInfo = this.classifications[classification];

        // Reduced logging - only log classification level without full details
        // console.debug('[ClassificationDetector] Classification detected:', classification, 'from line:', firstReadableLine);

        return {
            detected: true,
            text: classification,
            restricted: classificationInfo.restricted,
            warning: classificationInfo.restricted,
            color: classificationInfo.color,
            details: `Classification found on first line: ${this.getClassificationMessage(classification)}`,
            markings: [{
                text: `First Line Classification: ${classification}`,
                classification: classification,
                restricted: classificationInfo.restricted,
                line: lineNumber
            }],
            source: 'first_line'
        };
    }

    /**
     * Fallback pattern-based detection for emails without structured blocks
     * Only checks the first readable line for classification patterns
     * @param {string} emailBody - The email body text
     * @returns {Object} Classification detection result
     */
    detectClassificationByPatterns(emailBody) {

        // Only check the first readable line for classification patterns
        const lines = emailBody.split('\n');
        let firstReadableLine = null;
        
        for (let i = 0; i < lines.length && i < 5; i++) { // Don't look beyond first 5 lines
            const line = lines[i].trim();
            if (line) {
                firstReadableLine = line;
                break;
            }
        }

        if (!firstReadableLine) {
            return {
                detected: false,
                text: 'UNCLASSIFIED',
                restricted: false,
                warning: false,
                details: 'No readable content found'
            };
        }

        // Reduced logging for pattern detection
        // console.debug('[ClassificationDetector] Pattern detection on first line:', firstReadableLine);

        let highestClassification = null;
        let detectedMarkings = [];

        // Check each pattern against the first line only
        for (const pattern of this.patterns) {
            const matches = firstReadableLine.matchAll(pattern);
            
            for (const match of matches) {
                let classification = null;
                
                // Extract classification based on pattern structure
                if (match[0].includes('CLASSIFICATION:') || match[0].includes('CLASS:')) {
                    // Pattern with banner - classification is in match[2]
                    classification = this.normalizeClassification(match[2]);
                } else {
                    // Other patterns - classification is in match[1]
                    classification = this.normalizeClassification(match[1]);
                }
                
                if (this.classifications[classification]) {
                    detectedMarkings.push({
                        text: match[0].trim(),
                        classification: classification,
                        restricted: this.classifications[classification].restricted,
                        line: 1
                    });

                    // Determine highest classification using priority order
                    if (!highestClassification || this.isHigherClassification(classification, highestClassification.text)) {
                        highestClassification = {
                            text: classification,
                            color: this.classifications[classification].color,
                            restricted: this.classifications[classification].restricted
                        };
                    }
                }
            }
        }

        // If no classification found, assume unclassified
        if (!highestClassification) {
            // Reduced logging to avoid console noise
            // console.debug('[ClassificationDetector] No classification patterns found in first line');
            return {
                detected: false,
                text: 'UNCLASSIFIED',
                restricted: false,
                warning: false,
                details: 'No classification markings detected',
                markings: []
            };
        }

        // Determine if warning should be shown (restricted classifications)
        const shouldWarn = highestClassification.restricted;

        return {
            detected: true,
            text: highestClassification.text,
            restricted: highestClassification.restricted,
            warning: shouldWarn,
            color: highestClassification.color,
            details: this.getClassificationMessage(highestClassification.text),
            markings: detectedMarkings
        };
    }

    /**
     * Determines if one classification is higher than another
     * @param {string} classification1 - First classification
     * @param {string} classification2 - Second classification  
     * @returns {boolean} True if classification1 is higher
     */
    isHigherClassification(classification1, classification2) {
        const priority = ['UNCLASSIFIED', 'CONFIDENTIAL', 'SECRET', 'TOP SECRET', 'COSMIC TOP SECRET'];
        
        // Normalize TS and CTS
        const norm1 = classification1 === 'TS' ? 'TOP SECRET' : 
                     classification1 === 'CTS' ? 'COSMIC TOP SECRET' : classification1;
        const norm2 = classification2 === 'TS' ? 'TOP SECRET' : 
                     classification2 === 'CTS' ? 'COSMIC TOP SECRET' : classification2;
        
        const index1 = priority.indexOf(norm1);
        const index2 = priority.indexOf(norm2);
        
        return index1 > index2;
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
     * Finds the line number where a classification block was found
     * @param {string} text - Full text
     * @param {string} classification - Classification found
     * @returns {number} Line number (1-based)
     */
    findClassificationBlockLine(text, classification) {
        const lines = text.split('\n');
        for (let i = 0; i < Math.min(lines.length, 15); i++) {
            if (lines[i].toUpperCase().includes(classification)) {
                return i + 1;
            }
        }
        return 1;
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
     * @param {string} classification - Classification string
     * @returns {string} User-friendly message
     */
    getClassificationMessage(classification) {
        switch (classification) {
            case 'UNCLASSIFIED':
                return 'This content is unclassified and safe to process.';
            case 'CONFIDENTIAL':
                return 'This content is marked CONFIDENTIAL. Exercise caution when sharing with external systems.';
            case 'SECRET':
                return 'This content is marked SECRET. Sharing with external AI services may violate security policies.';
            case 'TOP SECRET':
            case 'TS':
                return 'This content is marked TOP SECRET. External processing is strictly prohibited by security policy.';
            case 'COSMIC TOP SECRET':
            case 'CTS':
                return 'This content has the highest classification level. External processing is absolutely forbidden.';
            default:
                return 'Classification level could not be determined. Proceed with caution.';
        }
    }

    /**
     * Validates if processing should be allowed for a classification
     * @param {string} classification - Classification string
     * @param {boolean} userOverride - Whether user has overridden the warning
     * @returns {Object} Validation result
     */
    validateProcessing(classification, userOverride = false) {
        const classificationInfo = this.classifications[classification];
        
        if (!classificationInfo || !classificationInfo.restricted) {
            return {
                allowed: true,
                requiresWarning: false,
                message: 'Processing allowed for this classification level.'
            };
        }

        if (classificationInfo.restricted && !userOverride) {
            return {
                allowed: false,
                requiresWarning: true,
                message: 'Processing blocked due to classification restrictions. User override required.'
            };
        }

        if (classificationInfo.restricted && userOverride) {
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
        const validation = this.validateProcessing(detection.text);

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
        // Block TOP SECRET and above
        return detection.text === 'TOP SECRET' || detection.text === 'TS' || 
               detection.text === 'COSMIC TOP SECRET' || detection.text === 'CTS';
    }

    /**
     * Gets CSS class for classification styling
     * @param {string} classification - Classification string
     * @returns {string} CSS class name
     */
    getClassificationStyle(classification) {
        switch (classification) {
            case 'UNCLASSIFIED': return 'classification-unclassified';
            case 'CONFIDENTIAL': return 'classification-confidential';
            case 'SECRET': return 'classification-secret';
            case 'TOP SECRET':
            case 'TS':
            case 'COSMIC TOP SECRET': 
            case 'CTS': return 'classification-top-secret';
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
        if (!detection.detected || detection.text === 'UNCLASSIFIED') {
            // For unclassified, return first 100 characters
            return content.substring(0, 100) + (content.length > 100 ? '...' : '');
        }

        // For classified content, return only metadata
        return `[CLASSIFIED CONTENT - ${detection.text} - ${content.length} chars]`;
    }
}

// For Node.js testing compatibility
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ClassificationDetector;
}
