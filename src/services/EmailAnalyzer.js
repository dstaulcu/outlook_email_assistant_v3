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

            // Create promises for async property access
            const getFromAsync = () => {
                return new Promise((resolveFrom) => {
                    if (item.from && typeof item.from.getAsync === 'function') {
                        item.from.getAsync((result) => {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                resolveFrom(result.value);
                            } else {
                                console.warn('Failed to get from async:', result.error);
                                resolveFrom(null);
                            }
                        });
                    } else {
                        resolveFrom(item.from);
                    }
                });
            };

            const getRecipientsAsync = () => {
                return new Promise((resolveRecipients) => {
                    if (item.to && typeof item.to.getAsync === 'function') {
                        item.to.getAsync((result) => {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                resolveRecipients({ to: result.value, cc: item.cc, bcc: item.bcc });
                            } else {
                                console.warn('Failed to get recipients async:', result.error);
                                resolveRecipients({ to: null, cc: item.cc, bcc: item.bcc });
                            }
                        });
                    } else {
                        resolveRecipients({ to: item.to, cc: item.cc, bcc: item.bcc });
                    }
                });
            };

            const getSubjectAsync = () => {
                return new Promise((resolveSubject) => {
                    if (item.subject && typeof item.subject.getAsync === 'function') {
                        item.subject.getAsync((result) => {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                resolveSubject(result.value);
                            } else {
                                console.warn('Failed to get subject async:', result.error);
                                resolveSubject(item.subject);
                            }
                        });
                    } else {
                        resolveSubject(item.subject);
                    }
                });
            };

            const getDateAsync = () => {
                return new Promise((resolveDate) => {
                    
                    
                    // For compose mode, especially replies, we don't want to show a misleading sent date
                    // We'll determine this based on the item properties and context
                    // More accurate compose mode detection:
                    const isComposeMode = item.itemType === Office.MailboxEnums.ItemType.Message && 
                                         !item.internetMessageId && 
                                         !item.dateTimeCreated;
                    
                    console.debug('Compose mode detection details:', {
                        itemType: item.itemType,
                        itemClass: item.itemClass, 
                        internetMessageId: !!item.internetMessageId,
                        dateTimeCreated: !!item.dateTimeCreated,
                        hasInternetMessageId: !!item.internetMessageId,
                        hasDateTimeCreated: !!item.dateTimeCreated,
                        isComposeMode: isComposeMode
                    });
                    
                    if (isComposeMode) {
                        // In compose mode (including replies), don't show a sent date
                        console.debug('In compose mode, using null date');
                        resolveDate(null);
                    } else if (item.itemType === Office.MailboxEnums.ItemType.Message) {
                        // This should be a received email - use dateTimeCreated
                        const emailDate = item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date();
                        console.debug('Using email date for received message:', emailDate);
                        resolveDate(emailDate);
                    } else {
                        // For other item types
                        console.debug('Other item type, using null date');
                        resolveDate(null);
                    }
                });
            };

            // Get email body and other properties
            Promise.all([
                new Promise((resolveBody, rejectBody) => {
                    item.body.getAsync(Office.CoercionType.Text, (result) => {
                        if (result.status === Office.AsyncResultStatus.Failed) {
                            rejectBody(new Error('Failed to get email body: ' + result.error.message));
                            return;
                        }
                        resolveBody(result.value || '');
                    });
                }),
                getFromAsync(),
                getRecipientsAsync(),
                getSubjectAsync(),
                getDateAsync()
            ]).then(([bodyText, fromValue, recipientsValue, subjectValue, dateValue]) => {
                console.log('Async property results:', {
                    subject: subjectValue,
                    from: fromValue,
                    recipients: recipientsValue,
                    date: dateValue,
                    itemType: item.itemType
                });

                const userProfile = Office.context.mailbox.userProfile;
                
                // Check if this is a reply after we have the subject
                const subjectStr = this.getSubjectString({ subject: subjectValue });
                const isReply = subjectStr && (subjectStr.startsWith('RE:') || subjectStr.startsWith('Re:') || subjectStr.startsWith('FW:') || subjectStr.startsWith('Fw:'));
                
                const emailData = {
                    subject: subjectStr,
                    from: this.getFromAddressFromValue(fromValue, item),
                    recipients: this.getRecipientsFromValue(recipientsValue),
                    body: bodyText,
                    bodyLength: bodyText.length,
                    date: dateValue,
                    isReply: isReply, // Add this flag
                    hasAttachments: (item.attachments && item.attachments.length > 0),
                    itemType: item.itemType,
                    conversationId: item.conversationId,
                    sender: userProfile ? `${userProfile.displayName || 'Unknown'} <${userProfile.emailAddress || 'unknown@domain.com'}>` : 'Unknown Sender'
                };

                console.log('Final processed email data:', emailData);
                resolve(emailData);
            }).catch(reject);
        });
    }

    /**
     * Gets the sender's email address from async value
     * @param {Object} fromValue - The async from value
     * @param {Office.Item} item - The Outlook item
     * @returns {string} Sender email address
     */
    getFromAddressFromValue(fromValue, item) {
        console.log('Processing sender email address from async value:', fromValue);
        
        if (fromValue) {
            const displayName = (fromValue.displayName !== undefined) ? String(fromValue.displayName) : 'Unknown';
            const emailAddress = (fromValue.emailAddress !== undefined) ? String(fromValue.emailAddress) : 'unknown@domain.com';
            return `${displayName} <${emailAddress}>`;
        }
        
        // Fallback to synchronous access
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            return 'Unknown Sender';
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
     * Gets recipients from async values
     * @param {Object} recipientsValue - Object with to, cc, bcc arrays
     * @returns {string} Formatted recipients string
     */
    getRecipientsFromValue(recipientsValue) {
        console.log('Processing recipients from async values:', recipientsValue);
        
        const recipients = [];
        
        // Helper function to safely format recipient
        const formatRecipient = (r) => {
            if (!r) return 'Unknown <unknown@domain.com>';
            const displayName = (r.displayName !== undefined) ? String(r.displayName) : 'Unknown';
            const emailAddress = (r.emailAddress !== undefined) ? String(r.emailAddress) : 'unknown@domain.com';
            return `${displayName} <${emailAddress}>`;
        };

        try {
            // Get To recipients
            if (recipientsValue.to && Array.isArray(recipientsValue.to) && recipientsValue.to.length > 0) {
                const toRecipients = recipientsValue.to.map(formatRecipient);
                recipients.push('To: ' + toRecipients.join(', '));
            }

            // Get CC recipients
            if (recipientsValue.cc && Array.isArray(recipientsValue.cc) && recipientsValue.cc.length > 0) {
                const ccRecipients = recipientsValue.cc.map(formatRecipient);
                recipients.push('CC: ' + ccRecipients.join(', '));
            }

            // Get BCC recipients (if available)
            if (recipientsValue.bcc && Array.isArray(recipientsValue.bcc) && recipientsValue.bcc.length > 0) {
                const bccRecipients = recipientsValue.bcc.map(formatRecipient);
                recipients.push('BCC: ' + bccRecipients.join(', '));
            }
        } catch (error) {
            console.error('Error processing recipients value:', error);
        }

        const result = recipients.join('; ') || 'No recipients';
        console.log('Final recipients string:', result);
        return result;
    }

    /**
     * Gets the subject as a string
     * @param {Office.Item} item - The Outlook item
     * @returns {string} Subject string
     */
    getSubjectString(item) {
        console.log('Processing subject:', item.subject, typeof item.subject);
        
        if (!item.subject) return 'No Subject';
        
        // Handle case where subject might be an object
        if (typeof item.subject === 'object') {
            console.log('Subject is object:', item.subject);
            // If it has a value property, use that
            if (item.subject && item.subject.value !== undefined) return String(item.subject.value);
            // If it has a text property, use that
            if (item.subject && item.subject.text !== undefined) return String(item.subject.text);
            // If it has a toString method, use that
            if (item.subject && typeof item.subject.toString === 'function') {
                const stringified = item.subject.toString();
                if (stringified !== '[object Object]') return stringified;
            }
            // Try JSON.stringify as last resort
            try {
                const jsonString = JSON.stringify(item.subject);
                if (jsonString && jsonString !== '{}') return jsonString;
            } catch (e) {
                console.warn('Failed to stringify subject:', e);
            }
            // Otherwise return empty string
            return 'No Subject';
        }
        
        return String(item.subject);
    }

    /**
     * Gets the sender's email address with multiple fallback strategies
     * @param {Office.Item} item - The Outlook item
     * @returns {string} Sender email address
     */
    getFromAddress(item) {
        console.log('Processing from address:', {
            itemType: item.itemType,
            from: item.from,
            messageType: Office.MailboxEnums.ItemType.Message,
            isReadMode: Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message
        });
        
        // Try different approaches based on the item type and mode
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            // For received messages in read mode
            if (item.from) {
                console.log('From object:', item.from);
                const displayName = (item.from.displayName !== undefined) ? String(item.from.displayName) : 'Unknown';
                const emailAddress = (item.from.emailAddress !== undefined) ? String(item.from.emailAddress) : 'unknown@domain.com';
                return `${displayName} <${emailAddress}>`;
            } 
            
            // Fallback: try to get sender from internetMessageId or other properties
            if (item.sender) {
                console.log('Using sender property:', item.sender);
                const displayName = (item.sender.displayName !== undefined) ? String(item.sender.displayName) : 'Unknown';
                const emailAddress = (item.sender.emailAddress !== undefined) ? String(item.sender.emailAddress) : 'unknown@domain.com';
                return `${displayName} <${emailAddress}>`;
            }
            
            return 'Unknown Sender';
        } else {
            // For compose items, return current user
            const userProfile = Office.context.mailbox.userProfile;
            console.log('User profile:', userProfile);
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
     * Gets all recipients (To, CC, BCC) with enhanced error handling
     * @param {Office.Item} item - The Outlook item
     * @returns {string} Formatted recipients string
     */
    getRecipients(item) {
        console.log('Processing recipients:', {
            to: item.to,
            cc: item.cc,
            bcc: item.bcc
        });
        
        const recipients = [];

        // Helper function to safely format recipient
        const formatRecipient = (r) => {
            if (!r) return 'Unknown <unknown@domain.com>';
            const displayName = (r.displayName !== undefined) ? String(r.displayName) : 'Unknown';
            const emailAddress = (r.emailAddress !== undefined) ? String(r.emailAddress) : 'unknown@domain.com';
            return `${displayName} <${emailAddress}>`;
        };

        try {
            // Get To recipients
            if (item.to && Array.isArray(item.to) && item.to.length > 0) {
                const toRecipients = item.to.map(formatRecipient);
                recipients.push('To: ' + toRecipients.join(', '));
            }

            // Get CC recipients
            if (item.cc && Array.isArray(item.cc) && item.cc.length > 0) {
                const ccRecipients = item.cc.map(formatRecipient);
                recipients.push('CC: ' + ccRecipients.join(', '));
            }

            // Get BCC recipients (if available)
            if (item.bcc && Array.isArray(item.bcc) && item.bcc.length > 0) {
                const bccRecipients = item.bcc.map(formatRecipient);
                recipients.push('BCC: ' + bccRecipients.join(', '));
            }
        } catch (error) {
            console.error('Error processing recipients:', error);
        }

        const result = recipients.join('; ') || 'No recipients';
        console.log('Processed recipients:', result);
        return result;
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

