# Email Identifiers for Compliance and Review

This document outlines the non-content-revealing identifiers available for email tracking and compliance review in the Outlook Email Assistant.

## Available Email Identifiers

### Primary Identifiers
These are unique identifiers provided by Outlook/Office.js that can be used to locate specific emails for compliance review:

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

## Usage in Security Compliance

### Classification Override Tracking
When users override security classification warnings, the system logs:
- All email identifiers above
- Classification detected and markings count
- Provider used and supported classifications  
- User ID and timestamp
- Type of override action taken

### Provider Incompatibility Logging
When emails contain classifications not supported by the current provider:
- Email identifiers for review lookup
- Provider and classification details
- List of classifications supported by the provider

## Compliance Review Process

To review a specific email incident:

1. **Locate by conversationId** - Find all emails in the thread
2. **Locate by itemId** - Find the exact email item  
3. **Use subjectHash** - Correlate with other systems without revealing content
4. **Check bodyLength/hasAttachments** - Verify email characteristics
5. **Review date/isReply** - Understand email context

## Professional Use Context

The telemetry system recognizes that national security professionals may need to override 
classification restrictions when operational requirements demand it. The system:

- **Records override events** for pattern analysis and compliance documentation
- **Provides factual data** without prejudgment about appropriateness of actions
- **Enables oversight** through aggregate analysis rather than individual incident review
- **Supports professional discretion** while maintaining audit capabilities

## Privacy Protection

All identifiers are designed to:
- Enable email location for compliance review
- Avoid revealing email content or subject text
- Provide sufficient metadata for incident analysis
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
  "classification_detected": "SECRET",
  "provider_used": "ollama-local",
  "provider_supported_classifications": ["UNCLASSIFIED", "CONFIDENTIAL"],
  "userId": "user@company.com",
  "warning_type": "user_override",
  "timestamp": "2025-08-18T15:31:23.456Z",
  "eventType": "classification_warning_overridden",
  "sessionId": "sess_123456789_abc",
  "source": "PromptEmail",
  "version": "1.0.0"
}
```

This structure allows security teams to:
- Locate the specific email for manual review
- Understand the operational context 
- Track usage patterns over time
- Support compliance documentation and oversight
