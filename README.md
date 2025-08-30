# PromptEmail Outlook Add-in

AI-Powered Email Analysis Outlook Add-in that enhances email productivity through large language model integration.

## Features

- **AI-Powered Email Analysis**: Intelligent email analysis, tone detection, and response assistance
- **Multi-Provider AI Support**: OpenAI-compatible providers, Ollama (local), and custom on-site endpoints
- **Accessible Interface**: Full keyboard navigation, screen reader support, and high-contrast mode
- **Advanced Settings Management**: Persistent settings with Office 365 roaming and local backup
- **Real-time Email Processing**: Extract email content, analyze sentiment, and generate contextual responses
- **Cloud Telemetry Pipeline**: Secure telemetry via AWS API Gateway eliminating CORS and credential issues
- **Windows Logging Integration**: PowerShell-based event logging to Windows Application Log

## Quick Start

### Prerequisites

- Node.js 16+ and npm
- Outlook Desktop (Microsoft 365 subscription)
- PowerShell 5.1+ or PowerShell Core 7+ (Windows)
- AWS CLI configured (for deployment to S3)

### Installation

1. Clone the repository and install dependencies:
```bash
npm install
```

2. Build the project:
```bash
npm run build
```

3. Configure deployment environment in `tools\deployment-environments.json`

4. Deploy using the build and deploy script:
```bash
.\tools\deploy_web_assets.ps1 -Environment Prod
```

5. Sideload the manifest (`manifest.xml`) in Outlook

### Development

- **Start development watcher**: `npm run dev`
- **Build for production**: `npm run build`
- **Deploy to environment**: `.\tools\deploy_web_assets.ps1 -Environment Dev`
- **Validate manifest**: `npm run validate-manifest`

## Architecture

### Frontend Stack
- **Framework**: Vanilla JavaScript with ES6 modules and classes
- **Build System**: Webpack 5 with HtmlWebpackPlugin and CopyWebpackPlugin
- **Styling**: CSS3 with CSS custom properties and Grid/Flexbox layouts
- **Bundling**: Separate bundles for taskpane and commands functionality

### Core Services Architecture
- **EmailAnalyzer**: Extracts and processes email content from Office.js API
- **AIService**: Multi-provider AI integration (OpenAI-compatible, Ollama, Custom on-site providers)
- **SettingsManager**: Dual-storage settings (Office.js roaming + localStorage backup)
- **Logger**: Cloud telemetry via AWS API Gateway and Windows Application Log integration
- **UIController**: State management, loading states, and user feedback
- **AccessibilityManager**: ARIA live regions, keyboard navigation, screen reader support

### Deployment Architecture
- **Hosting**: AWS S3 static website hosting
- **Telemetry Pipeline**: AWS API Gateway → Lambda → EC2 Splunk Enterprise for secure data collection
- **Build Pipeline**: PowerShell-based build and deployment automation
- **Environment Management**: Multi-environment support (Dev/Test/Prod) with URL rewriting
- **Asset Management**: Automated file discovery and URL updating for deployments

### Office Integration
- **Manifest**: Office Add-in manifest with ribbon integration
- **API Integration**: Office.js API for email reading, writing, and user context
- **Extension Points**: Message read and compose command surfaces
- **Authentication**: Office 365 user profile integration

## Security & Compliance

- **Audit Logging**: Comprehensive event logging to Windows Application Log with sanitized data
- **Data Privacy**: Sensitive content filtering in logs and telemetry
- **API Key Security**: Secure storage using Office.js roaming settings with localStorage fallback
- **Content Sanitization**: Automatic removal of sensitive data from all logging and telemetry

## Telemetry & Monitoring

The application provides comprehensive telemetry for operational monitoring, security compliance, and user experience analytics. All telemetry is transmitted through a secure AWS API Gateway pipeline with exponential backoff retry logic.

### Telemetry Pipeline Architecture
- **Collection**: Client-side JavaScript Logger service with intelligent batching
- **Transport**: HTTPS POST to AWS API Gateway with CORS support
- **Processing**: AWS Lambda functions for data validation and enrichment  
- **Storage**: Splunk Enterprise on EC2 for analysis and dashboards
- **Retry Logic**: Exponential backoff (1s, 2s, 4s, 8s) for transient failures
- **Error Handling**: Permanent error detection with graceful event dropping

### Event Data Dictionary

#### Core Event Structure
All telemetry events share a common base structure:

| Field | Type | Description | Example |
|-------|------|-------------|---------|
| `eventType` | string | Event category identifier | `"session_start"` |
| `timestamp` | string | ISO 8601 timestamp with milliseconds | `"2025-08-26T02:03:47.401Z"` |
| `source` | string | Application identifier | `"PromptEmail"` |
| `version` | string | Application version from package.json | `"1.0.0"` |
| `sessionId` | string | Unique session identifier | `"sess_1756173824497_limudw6qt"` |
| `userContext` | object | User identification context | `{"email": "user@domain.com"}` |

#### Session Events

**`session_start`**
Triggered when the Outlook add-in initializes.

| Field | Type | Description |
|-------|------|-------------|
| `host` | string | Office application host | `"Outlook"` |

**`session_summary`**
Sent on session end with aggregated metrics.

| Field | Type | Description |
|-------|------|-------------|
| `session_duration_ms` | integer | Total session duration in milliseconds |
| `email_analyzed` | boolean | Whether any emails were processed |
| `response_generated` | boolean | Whether AI responses were generated |
| `clipboard_used` | boolean | Whether suggestions were copied |
| `refinement_count` | integer | Number of response refinements |

#### Email Processing Events

**`email_analyzed`**
Captured for each email analysis operation.

| Field | Type | Description |
|-------|------|-------------|
| `model_service` | string | AI provider used | `"ollama"` |
| `model_name` | string | Specific model identifier | `"llama3:latest"` |
| `email_length` | integer | Character count of email content |
| `recipients_count` | integer | Number of email recipients |
| `analysis_success` | boolean | Whether analysis completed successfully |
| `refinement_count` | integer | Number of refinements performed |
| `clipboard_used` | boolean | Whether response was copied |
| `performance_metrics` | object | Timing and performance data |

**`response_copied`**
Logged when user copies AI-generated response to clipboard.

| Field | Type | Description |
|-------|------|-------------|
| `refinement_count` | integer | Number of refinements before copy |
| `response_length` | integer | Character count of copied response |

#### Performance Events

**`model_refresh`**
Triggered when AI models are refreshed from providers.

| Field | Type | Description |
|-------|------|-------------|
| `provider` | string | AI provider being refreshed |
| `models_discovered` | integer | Number of available models |
| `refresh_duration_ms` | integer | Time taken to discover models |
| `success` | boolean | Whether refresh succeeded |
| `trigger` | string | What triggered the refresh |

**`error_event`**
Captures application errors for debugging and reliability metrics.

| Field | Type | Description |
|-------|------|-------------|
| `error_type` | string | Category of error |
| `error_message` | string | Sanitized error description |
| `stack_trace` | string | Truncated stack trace (first 200 chars) |
| `recovery_attempted` | boolean | Whether automatic recovery was tried |
| `user_impact` | string | Severity of user experience impact |

#### Real-Time Event Examples

Based on live telemetry data, here are actual event examples:

```json
{
  "eventType": "session_start",
  "timestamp": "2025-08-26T02:19:07.067Z",
  "source": "PromptEmail",
  "version": "1.0.0",
  "sessionId": "sess_1756174747068_pvgpfsa04",
  "userContext": {"email": "user@domain.com"},
  "host": "Outlook"
}
```

```json
{
  "eventType": "email_analyzed",
  "timestamp": "2025-08-26T02:19:33.491Z",
  "sessionId": "sess_1756174747068_pvgpfsa04",
  "userContext": {"email": "user@domain.com"},
  "model_service": "ollama",
  "model_name": "llama3:latest",
  "email_length": 45,
  "recipients_count": 1,
  "analysis_success": true
}
```

```json
{
  "eventType": "session_summary",
  "timestamp": "2025-08-26T02:19:52.231Z",
  "sessionId": "sess_1756174747068_pvgpfsa04",
  "userContext": {"email": "user@domain.com"},
  "session_duration_ms": 45217,
  "refinement_count": 0,
  "clipboard_used": true,
  "email_analyzed": true,
  "response_generated": true
}
```
#### User Interaction Events

**`feature_usage`**
Tracks which features are being utilized.

| Field | Type | Description |
|-------|------|-------------|
| `feature_name` | string | Feature identifier |
| `interaction_type` | string | Type of user interaction |
| `duration_ms` | integer | Time spent in feature |
| `completion_status` | string | Whether feature use completed |

### Data Privacy & Security

#### Sanitization Rules
- **Email Content**: Never logged or transmitted to protect user privacy
- **Email Subjects**: Never logged in plain text - only hashed subjectHash for correlation
- **API Keys**: Completely filtered from all telemetry  
- **Personal Information**: Email addresses are the only user identifier (required for user journey tracking)
- **Error Messages**: Sensitive paths and credentials stripped from stack traces
- **Email Metadata**: Outlook IDs are hashed or used only for session correlation

#### User Context Policy
- **Consistent Identification**: All events include userContext.email for user journey tracking
- **Session Correlation**: sessionId links all events in a user workflow
- **Privacy Balance**: Minimal user data while enabling operational analytics
- **Compliance**: User identification supports audit requirements for secure data handling

#### Retention Policy
- **Real-time Data**: 90 days for operational monitoring
- **Aggregated Metrics**: 2 years for trend analysis
- **Security Events**: 7 years for compliance audit trail
- **Error Logs**: 1 year for debugging and reliability improvement

#### Compliance Features
- **GDPR**: Anonymized user contexts and data minimization
- **SOC 2**: Secure transmission and access controls
- **Data Security**: Proper handling of sensitive email content and user information
- **Audit Trail**: Complete event lineage for security reviews

## Directory Structure

```
├── public/                     # Built assets (generated by webpack)
├── src/                       # Source code
│   ├── services/              # Core business logic services
│   │   ├── AIService.js       # Multi-provider AI integration
│   │   ├── EmailAnalyzer.js   # Email content extraction
│   │   ├── SettingsManager.js # Persistent settings management
│   │   └── Logger.js          # Windows event logging
│   ├── ui/                    # UI components and accessibility
│   │   ├── UIController.js    # State management and user feedback
│   │   └── AccessibilityManager.js # ARIA and keyboard navigation
│   ├── taskpane/              # Main application interface
│   │   ├── taskpane.html      # Application HTML template
│   │   └── taskpane.js        # Main application controller
│   ├── commands/              # Office ribbon commands
│   │   ├── commands.html      # Command page template
│   │   └── commands.js        # Ribbon command handlers
│   ├── assets/               # Static assets (CSS, icons)
│   ├── config/                # Configuration files
│   │   ├── ai-providers.json # AI provider configurations (includes model lists)
│   │   └── telemetry.json    # Telemetry and logging configuration
│   └── manifest.xml          # Office Add-in manifest
├── tools/                    # Build and deployment scripts
│   ├── deploy_web_assets.ps1  # Main build and deploy automation
│   └── deployment-environments.json # Environment configurations
├── webpack.config.js        # Build configuration
└── package.json            # Project dependencies and scripts
```

## AI Provider Support

The add-in supports multiple AI providers with automatic model discovery:

- **OpenAI-Compatible**: Standard OpenAI API endpoints with API key authentication
- **Ollama**: Local LLM hosting with automatic model detection via `/api/tags`
- **Custom On-Site Providers**: Support for OnSiteProvider-1, OnSiteProvider-2, and other OpenAI-compatible endpoints available in your work environment

### Default Configurations
- Provider endpoints and models are defined in `src/default-providers.json` and `src/default-models.json`
- Runtime model discovery for Ollama installations
- Fallback configurations for offline scenarios

## Deployment

See `DEPLOYMENT_GUIDE.md` for detailed deployment instructions.

## Contributing

1. Fork the repository and create a feature branch
2. Follow the development setup in `DEVELOPER_GUIDE.md`
3. Make your changes with appropriate tests
4. Ensure accessibility compliance (WCAG 2.1 AA)
5. Update documentation as needed
6. Submit a pull request with clear description

## License

MIT License - see LICENSE file for details.
