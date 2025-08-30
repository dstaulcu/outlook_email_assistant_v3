# Email Identifiers for Analysis and Review

This document outlines the non-content-revealing identifiers available for email tracking and analysis review in the Outlook Email Assistant.

## Available Email Identifiers

### Primary Identifiers
These are unique identifiers provided by Outlook/Office.js that can be used to locate specific emails for analysis review:

1. **conversationId** - Groups related emails in a conversation thread
2. **itemId** - Unique identifier for the specific email item in Outlook
3. **itemClass** - Outlook item class (e.g., "IPM.Note" for normal emails)

### Content-Safe Metadata
These provide email context without revealing content:

4. **subjectHash** - One-way hash of the subject line for correlation
5. **normalizedSubject** - Outlook's normalized subject (RE/FW prefixes removed)
6. **bodyLength** - Character count of email body
7. **hasAttachments** - Boolean indicating presence of attachments
8. **hasInternetMessageId** - Boolean indicating if email has internet message ID
9. **itemType** - Office.js item type (Message, Appointment, etc.)
10. **isReply** - Boolean indicating if this is a reply/forward
11. **date** - Email timestamp

## Usage in Analysis and Logging

### Email Analysis Tracking
When emails are analyzed, the system logs:
- All email identifiers above for correlation
- AI provider used for analysis
- User ID and timestamp
- Analysis results and success status

## Analysis Review Process

To review a specific email analysis:

1. **Locate by conversationId** - Find all emails in the thread
2. **Locate by itemId** - Find the exact email item  
3. **Use subjectHash** - Correlate with other systems without revealing content
4. **Check bodyLength/hasAttachments** - Verify email characteristics
5. **Review date/isReply** - Understand email context

## Professional Use Context

The telemetry system provides factual analysis data for:

- **Analysis tracking** for usage pattern analysis and system improvement
- **Performance metrics** without content disclosure
- **System monitoring** through aggregate analysis
- **Quality assurance** while maintaining privacy

## Privacy Protection

All identifiers are designed to:
- Enable email location for analysis review
- Avoid revealing email content or subject text
- Provide sufficient metadata for system analysis
- Comply with data protection requirements

## Example Telemetry Output

```json
{
  "conversationId": "AAMkADNh...conversation-guid...",
  "itemId": "AAMkADNh...item-guid...",
  "itemClass": "IPM.Note", 
  "subjectHash": "a1b2c3d4",
  "bodyLength": 1247,
  "hasAttachments": false,
  "hasInternetMessageId": true,
  "itemType": "Message",
  "isReply": true,
  "date": "2025-08-18T15:30:00.000Z",
  "provider_used": "ollama-local",
  "userId": "user@company.com",
  "timestamp": "2025-08-18T15:31:23.456Z",
  "eventType": "email_analyzed",
  "sessionId": "sess_123456789_abc",
  "source": "PromptEmail",
  "version": "1.0.0"
}
```

This structure allows analysis teams to:
- Locate the specific email for review if needed
- Understand the operational context 
- Track usage patterns over time
- Support system improvement and monitoring
