# Testing Guide

This guide provides comprehensive testing procedures for the PromptEmail Outlook Add-in.

## Testing Overview

### Testing Scope
- ✅ Manifest validation and sideloading
- ✅ UI functionality and accessibility
- ✅ Email analysis and AI integration
- ✅ Classification detection and warnings
- ✅ Response generation and refinement
- ✅ Settings persistence
- ✅ Error handling and user feedback
- ✅ Performance and reliability

## Pre-Testing Setup

### 1. Prepare Test Environment

```bash
# Install dependencies
npm install

# Build the project
npm run build

# Validate manifest
npm run validate-manifest
```

### 2. Create Test Email Samples

Prepare these test emails in Outlook:

**Unclassified Email:**
```
Subject: Team Meeting Tomorrow
From: colleague@company.com

Hi everyone,

Just a reminder about our team meeting tomorrow at 2 PM in Conference Room A.

Best regards,
John
```

**Classified Email (for testing warnings):**
```
Subject: Quarterly Report
From: manager@company.com

CONFIDENTIAL

This quarterly report contains sensitive financial information...
```

**Action-Required Email:**
```
Subject: URGENT: Budget Approval Needed
From: finance@company.com

Hi,

Please review and approve the attached budget by EOD today. This is urgent for our Q4 planning.

Thanks,
Finance Team
```

## Functional Testing

### 1. Manifest and Sideloading

**Test Steps:**
1. Validate manifest: `npm run validate-manifest`
2. Sideload manifest in Outlook
3. Verify ribbon button appears
4. Click button to open taskpane

**Expected Results:**
- ✅ Manifest validation passes
- ✅ Add-in appears in Outlook ribbon
- ✅ Taskpane opens without errors
- ✅ PromptEmail interface loads

### 2. Email Loading and Summary

**Test Steps:**
1. Select an email in Outlook
2. Open PromptEmail taskpane
3. Verify email summary displays

**Expected Results:**
- ✅ Email from, subject, recipients populate
- ✅ Email length shows character count
- ✅ Loading state shows then disappears

### 3. AI Configuration

**Test Steps:**
1. Set AI service (OpenAI, Anthropic, Azure, Custom)
2. Enter valid API key
3. Test custom endpoint configuration
4. Save settings

**Expected Results:**
- ✅ Service selector works
- ✅ Custom endpoint field shows/hides correctly
- ✅ Settings persist across sessions
- ✅ API key is masked in UI

### 4. Classification Detection

**Test with Unclassified Email:**
- ✅ No warning appears
- ✅ Analysis proceeds normally

**Test with Classified Email:**
- ✅ Classification warning displays
- ✅ Warning message explains risk
- ✅ User can cancel or override
- ✅ Override action is logged

### 5. Email Analysis

**Test Steps:**
1. Configure AI service with valid credentials
2. Click "Analyze Email" 
3. Review analysis results

**Expected Results:**
- ✅ Button shows loading state
- ✅ Analysis completes successfully
- ✅ Key points are identified
- ✅ Sentiment analysis provided
- ✅ Action items extracted
- ✅ Results display in Analysis tab

### 6. Response Generation

**Test Steps:**
1. Adjust response sliders (length, tone, urgency)
2. Add custom instructions
3. Click "Generate Response"
4. Review generated response

**Expected Results:**
- ✅ Response matches selected parameters
- ✅ Tone and length appropriate
- ✅ Response is contextually relevant
- ✅ Results display in Response tab

### 7. Response Refinement

**Test Steps:**
1. Generate initial response
2. Add refinement instructions
3. Click "Refine Response"
4. Compare refined version

**Expected Results:**
- ✅ Refinement button appears after generation
- ✅ Custom instructions are applied
- ✅ Response improves based on feedback
- ✅ Multiple refinements possible

### 8. Response Actions

**Test Steps:**
1. Generate a response
2. Click "Copy to Clipboard"
3. Click "Insert into Reply"

**Expected Results:**
- ✅ Copy functionality works
- ✅ Insert creates new reply with content
- ✅ Formatting is preserved

## Accessibility Testing

### 1. Keyboard Navigation

**Test Steps:**
1. Use only keyboard to navigate interface
2. Test Tab, Shift+Tab, Arrow keys
3. Verify all controls are reachable

**Expected Results:**
- ✅ All interactive elements focusable
- ✅ Focus indicators visible
- ✅ Logical tab order
- ✅ Keyboard shortcuts work (Alt+A, Alt+R, Alt+S)

### 2. Screen Reader Compatibility

**Test with Windows Narrator or NVDA:**
1. Enable screen reader
2. Navigate through interface
3. Perform key actions

**Expected Results:**
- ✅ All elements properly labeled
- ✅ Status messages announced
- ✅ Form validation errors read aloud
- ✅ Loading states communicated

### 3. High Contrast Mode

**Test Steps:**
1. Enable Windows high contrast mode
2. Review interface visibility
3. Test all functionality

**Expected Results:**
- ✅ All text remains readable
- ✅ Interactive elements clearly visible
- ✅ Focus indicators apparent
- ✅ Functionality unchanged

## Error Handling Testing

### 1. Network Errors

**Test Scenarios:**
- Disconnect internet during AI call
- Use invalid API endpoint
- Exceed API rate limits

**Expected Results:**
- ✅ Clear error messages displayed
- ✅ No app crashes or freezes
- ✅ User can retry or reconfigure

### 2. Invalid Configurations

**Test Scenarios:**
- Empty API key
- Malformed endpoint URL
- Invalid model selection

**Expected Results:**
- ✅ Validation errors shown
- ✅ Specific guidance provided
- ✅ App remains stable

### 3. Office Integration Errors

**Test Scenarios:**
- No email selected
- Email loading fails
- Reply insertion fails

**Expected Results:**
- ✅ Appropriate error messages
- ✅ Graceful degradation
- ✅ Recovery suggestions provided

## Performance Testing

### 1. Load Testing

**Test Steps:**
1. Process various email sizes (100 chars to 10,000+ chars)
2. Test with multiple recipients
3. Measure response times

**Expected Results:**
- ✅ Handles emails up to 50KB
- ✅ Response times under 30 seconds
- ✅ No memory leaks over extended use

### 2. UI Responsiveness

**Test Steps:**
1. Rapid clicking of buttons
2. Quick navigation between tabs
3. Fast setting changes

**Expected Results:**
- ✅ UI remains responsive
- ✅ No duplicate operations
- ✅ State changes handled correctly

## Security Testing

### 1. Data Handling

**Verify:**
- ✅ API keys not logged or displayed
- ✅ Classified content not sent to logs
- ✅ Sensitive data excluded from telemetry
- ✅ Settings stored securely

### 2. Classification Protection

**Test:**
- ✅ SECRET level emails trigger warnings
- ✅ Override attempts are logged
- ✅ User identification in logs (anonymized)
- ✅ Compliance audit trail maintained

## Cross-Platform Testing

### 1. Outlook Versions

Test on:
- ✅ Outlook Desktop (Microsoft 365)
- ✅ Outlook Web App
- ✅ Different Windows versions (10, 11)

### 2. Browser Compatibility

For web version:
- ✅ Chrome
- ✅ Edge
- ✅ Firefox
- ✅ Safari (if available)

## Automated Testing Checklist

### Unit Tests (Future Implementation)
- [ ] Service layer functions
- [ ] Utility functions
- [ ] Classification detection
- [ ] Settings management

### Integration Tests (Future Implementation)
- [ ] AI service connections
- [ ] Office.js integration
- [ ] End-to-end workflows

## Regression Testing

Before each release, verify:
- ✅ All primary workflows work
- ✅ Settings migrate correctly
- ✅ Previous bugs remain fixed
- ✅ Performance hasn't degraded

## Test Data Cleanup

After testing:
1. Clear test API usage (if metered)
2. Remove test entries from Windows logs
3. Reset add-in settings
4. Clear browser cache

## Bug Reporting Template

When issues are found:

```
**Bug Title:** Brief description

**Environment:**
- Outlook version: 
- Windows version:
- Add-in version:
- Browser (if web):

**Steps to Reproduce:**
1. 
2. 
3. 

**Expected Result:**

**Actual Result:**

**Screenshots/Logs:**

**Workaround (if any):**
```

## Test Sign-off Criteria

Before production deployment:
- ✅ All functional tests pass
- ✅ Accessibility requirements met
- ✅ Security validation complete
- ✅ Performance within acceptable limits
- ✅ Cross-platform compatibility verified
- ✅ Error handling tested thoroughly
- ✅ User acceptance testing completed
