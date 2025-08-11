/**
 * EmailAnalyzer Service
 * Handles extraction and analysis of email data from Outlook
 */

export class EmailAnalyzer {
    constructor() {
        this.currentItem = null;
    }

    /**
     * Gets the currently selected email in Outlook
     * @returns {Promise<Object>} Email data object
     */
    async getCurrentEmail() {
        return new Promise((resolve, reject) => {
            if (!Office.context.mailbox.item) {
                reject(new Error('No email item selected'));
                return;
            }

            const item = Office.context.mailbox.item;
            this.currentItem = item;

            // Get email body
            item.body.getAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    reject(new Error('Failed to get email body: ' + result.error.message));
                    return;
                }

                const userProfile = Office.context.mailbox.userProfile;
                const emailData = {
                    subject: this.getSubjectString(item),
                    from: this.getFromAddress(item),
                    recipients: this.getRecipients(item),
                    body: result.value || '',
                    bodyLength: (result.value || '').length,
                    date: item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date(),
                    hasAttachments: (item.attachments && item.attachments.length > 0),
                    itemType: item.itemType,
                    conversationId: item.conversationId,
                    sender: userProfile ? `${userProfile.displayName || 'Unknown'} <${userProfile.emailAddress || 'unknown@domain.com'}>` : 'Unknown Sender'
                };

                resolve(emailData);
            });
        });
    }

    /**
     * Gets the subject as a string
     * @param {Office.Item} item - The Outlook item
     * @returns {string} Subject string
     */
    getSubjectString(item) {
        if (!item.subject) return '';
        
        // Handle case where subject might be an object
        if (typeof item.subject === 'object') {
            // If it has a value property, use that
            if (item.subject.value) return String(item.subject.value);
            // If it has a toString method, use that
            if (typeof item.subject.toString === 'function') return item.subject.toString();
            // Otherwise return empty string
            return '';
        }
        
        return String(item.subject);
    }

    /**
     * Gets the sender's email address
     * @param {Office.Item} item - The Outlook item
     * @returns {string} Sender email address
     */
    getFromAddress(item) {
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            // For received messages
            if (item.from) {
                const displayName = item.from.displayName || 'Unknown';
                const emailAddress = item.from.emailAddress || 'unknown@domain.com';
                return `${displayName} <${emailAddress}>`;
            } else {
                return 'Unknown Sender';
            }
        } else {
            // For compose items, return current user
            const userProfile = Office.context.mailbox.userProfile;
            if (userProfile) {
                const displayName = userProfile.displayName || 'Unknown';
                const emailAddress = userProfile.emailAddress || 'unknown@domain.com';
                return `${displayName} <${emailAddress}>`;
            } else {
                return 'Current User';
            }
        }
    }

    /**
     * Gets all recipients (To, CC, BCC)
     * @param {Office.Item} item - The Outlook item
     * @returns {string} Formatted recipients string
     */
    getRecipients(item) {
        const recipients = [];

        // Helper function to safely format recipient
        const formatRecipient = (r) => {
            const displayName = r.displayName || 'Unknown';
            const emailAddress = r.emailAddress || 'unknown@domain.com';
            return `${displayName} <${emailAddress}>`;
        };

        // Get To recipients
        if (item.to && item.to.length > 0) {
            const toRecipients = item.to.map(formatRecipient);
            recipients.push('To: ' + toRecipients.join(', '));
        }

        // Get CC recipients
        if (item.cc && item.cc.length > 0) {
            const ccRecipients = item.cc.map(formatRecipient);
            recipients.push('CC: ' + ccRecipients.join(', '));
        }

        // Get BCC recipients (if available)
        if (item.bcc && item.bcc.length > 0) {
            const bccRecipients = item.bcc.map(formatRecipient);
            recipients.push('BCC: ' + bccRecipients.join(', '));
        }

        return recipients.join('; ') || 'No recipients';
    }

    /**
     * Extracts metadata about the email for analysis
     * @param {Object} emailData - The email data object
     * @returns {Object} Email metadata
     */
    extractMetadata(emailData) {
        return {
            wordCount: this.countWords(emailData.body),
            hasQuestions: this.hasQuestionMarks(emailData.body),
            hasActionItems: this.hasActionWords(emailData.body),
            hasDeadlines: this.hasDateMentions(emailData.body),
            formality: this.assessFormality(emailData.body),
            urgencyKeywords: this.findUrgencyKeywords(emailData.body),
            participantCount: this.countParticipants(emailData.recipients)
        };
    }

    /**
     * Counts words in text
     * @param {string} text - Text to analyze
     * @returns {number} Word count
     */
    countWords(text) {
        return text.trim().split(/\s+/).filter(word => word.length > 0).length;
    }

    /**
     * Checks if text contains question marks
     * @param {string} text - Text to analyze
     * @returns {boolean} True if questions found
     */
    hasQuestionMarks(text) {
        return text.includes('?');
    }

    /**
     * Checks for action-oriented words
     * @param {string} text - Text to analyze
     * @returns {boolean} True if action words found
     */
    hasActionWords(text) {
        const actionWords = [
            'please', 'need', 'require', 'should', 'must', 'action', 'task',
            'complete', 'finish', 'deliver', 'send', 'provide', 'review'
        ];
        
        const lowerText = text.toLowerCase();
        return actionWords.some(word => lowerText.includes(word));
    }

    /**
     * Checks for date/deadline mentions
     * @param {string} text - Text to analyze
     * @returns {boolean} True if dates found
     */
    hasDateMentions(text) {
        const datePattern = /\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b|\btoday\b|\btomorrow\b|\bdeadline\b|\bdue\b/i;
        return datePattern.test(text);
    }

    /**
     * Assesses formality level of text
     * @param {string} text - Text to analyze
     * @returns {string} Formality level
     */
    assessFormality(text) {
        const formalWords = ['please', 'kindly', 'respectfully', 'sincerely', 'regards'];
        const informalWords = ['hey', 'hi', 'thanks', 'cool', 'awesome'];
        
        const lowerText = text.toLowerCase();
        const formalCount = formalWords.filter(word => lowerText.includes(word)).length;
        const informalCount = informalWords.filter(word => lowerText.includes(word)).length;
        
        if (formalCount > informalCount) return 'formal';
        if (informalCount > formalCount) return 'informal';
        return 'neutral';
    }

    /**
     * Finds urgency keywords in text
     * @param {string} text - Text to analyze
     * @returns {Array} Array of found urgency keywords
     */
    findUrgencyKeywords(text) {
        const urgencyWords = [
            'urgent', 'asap', 'immediate', 'emergency', 'critical', 'priority',
            'rush', 'quickly', 'fast', 'soon', 'deadline'
        ];
        
        const lowerText = text.toLowerCase();
        return urgencyWords.filter(word => lowerText.includes(word));
    }

    /**
     * Counts unique participants in recipients string
     * @param {string} recipients - Recipients string
     * @returns {number} Number of participants
     */
    countParticipants(recipients) {
        if (!recipients || recipients === 'No recipients') return 0;
        
        // Extract email addresses using regex
        const emailPattern = /<([^>]+)>/g;
        const emails = new Set();
        let match;
        
        while ((match = emailPattern.exec(recipients)) !== null) {
            emails.add(match[1].toLowerCase());
        }
        
        return emails.size;
    }

    /**
     * Prepares email data for AI analysis
     * @param {Object} emailData - Raw email data
     * @returns {Object} Processed email data for AI
     */
    prepareForAI(emailData) {
        const metadata = this.extractMetadata(emailData);
        
        return {
            ...emailData,
            metadata,
            processedAt: new Date().toISOString(),
            cleanBody: this.cleanEmailBody(emailData.body)
        };
    }

    /**
     * Cleans email body by removing signatures, forwarded content, etc.
     * @param {string} body - Raw email body
     * @returns {string} Cleaned email body
     */
    cleanEmailBody(body) {
        let cleaned = body;
        
        // Remove common signature separators
        cleaned = cleaned.split(/(?:--\s*$|--- Original Message ---|From:.*Sent:)/m)[0];
        
        // Remove excessive whitespace
        cleaned = cleaned.replace(/\n{3,}/g, '\n\n');
        cleaned = cleaned.trim();
        
        return cleaned;
    }
}
