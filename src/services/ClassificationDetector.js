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

        // Common classification patterns
        this.patterns = [
            // Internal system classification block (primary pattern)
            // Matches: CLASSIFICATION_BLOCK at start of email
            /^(UNCLASSIFIED|CONFIDENTIAL|SECRET|TOP SECRET|TS|COSMIC TOP SECRET|CTS)\s*$/gim,
            
            // Standard classification markings (fallback)
            /^(UNCLASSIFIED|CONFIDENTIAL|SECRET|TOP SECRET|TS|COSMIC TOP SECRET|CTS)\s*$/gim,
            
            // Classification with additional markings
            /^(UNCLASSIFIED|CONFIDENTIAL|SECRET|TOP SECRET|TS)\/\/([A-Z\s\/]+)\s*$/gim,
            
            // Classification banners
            /^\s*(CLASSIFICATION:|CLASS:)\s*(UNCLASSIFIED|CONFIDENTIAL|SECRET|TOP SECRET|TS|COSMIC TOP SECRET|CTS)\s*$/gim,
            
            // Portion markings within content
            /\(([UCS]|CONFIDENTIAL|SECRET|TOP SECRET|TS)\)/gim
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

        // First, try to parse the structured classification block
        const structuredResult = this.parseClassificationBlock(emailBody);
        if (structuredResult.detected) {
            return structuredResult;
        }

        // Fallback to pattern-based detection
        return this.detectClassificationByPatterns(emailBody);
    }

    /**
     * Parses structured classification block from internal system
     * Expected format:
     * ====================================
     * CLASSIFICATION_BLOCK
     * ...
     * ...
     * ====================================
     * @param {string} emailBody - The email body text
     * @returns {Object} Classification detection result
     */
    parseClassificationBlock(emailBody) {
        // Look for the classification block structure in first few lines
        const lines = emailBody.split('\n').slice(0, 6); // Check more lines to handle empty lines
        
        let classificationLine = null;
        let blockFound = false;
        let classificationLineNumber = -1;
        
        // Check if first line looks like a separator (equals signs)
        if (lines.length >= 2) {
            const firstLine = lines[0].trim();
            if (firstLine.match(/^={10,}$/)) {
                // Look for classification in the next few lines (skip empty lines)
                for (let i = 1; i < Math.min(lines.length, 5); i++) {
                    const potentialClassification = lines[i].trim().toUpperCase();
                    
                    // Skip empty lines
                    if (!potentialClassification) {
                        continue;
                    }
                    
                    if (this.classifications[potentialClassification]) {
                        classificationLine = potentialClassification;
                        blockFound = true;
                        classificationLineNumber = i + 1; // Convert to 1-based line number
                        break;
                    } else {
                        // Also check if classification is part of a longer line
                        for (const classLevel of Object.keys(this.classifications)) {
                            if (potentialClassification.includes(classLevel)) {
                                classificationLine = classLevel;
                                blockFound = true;
                                classificationLineNumber = i + 1; // Convert to 1-based line number
                                break;
                            }
                        }
                        if (blockFound) break;
                    }
                }
            }
        }

        if (!blockFound || !classificationLine) {
            return { detected: false };
        }

        const classification = this.normalizeClassification(classificationLine);
        const classificationInfo = this.classifications[classification];

        return {
            detected: true,
            text: classification,
            restricted: classificationInfo.restricted,
            warning: classificationInfo.restricted,
            color: classificationInfo.color,
            details: `Structured classification block detected: ${this.getClassificationMessage(classification)}`,
            markings: [{
                text: `Classification Block: ${classification}`,
                classification: classification,
                restricted: classificationInfo.restricted,
                line: classificationLineNumber  // Use actual line number where classification was found
            }],
            source: 'structured_block'
        };
    }

    /**
     * Fallback pattern-based detection for emails without structured blocks
     * @param {string} emailBody - The email body text
     * @returns {Object} Classification detection result
     */
    detectClassificationByPatterns(emailBody) {

        // Get first few lines for classification checking
        const lines = emailBody.split('\n').slice(0, this.linesToCheck);
        const headerText = lines.join('\n');

        let highestClassification = null;
        let detectedMarkings = [];

        // Check each pattern
        for (const pattern of this.patterns) {
            const matches = headerText.matchAll(pattern);
            
            for (const match of matches) {
                let classification = null;
                
                // Extract classification based on pattern structure
                if (match[0].includes('CLASSIFICATION:') || match[0].includes('CLASS:')) {
                    // Pattern 3: classification banner - classification is in match[2]
                    classification = this.normalizeClassification(match[2]);
                } else {
                    // Patterns 1, 2, 4: classification is in match[1]
                    classification = this.normalizeClassification(match[1]);
                }
                
                if (this.classifications[classification]) {
                    detectedMarkings.push({
                        text: match[0].trim(),
                        classification: classification,
                        restricted: this.classifications[classification].restricted,
                        line: this.findLineNumber(emailBody, match[0])
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
