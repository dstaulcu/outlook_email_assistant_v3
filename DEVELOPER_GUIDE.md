# Developer Guide

This guide provides comprehensive instructions for setting up, developing, and testing the PromptEmail Outlook Add-in.

## Prerequisites

### Required Software
- **Node.js 16+**: [Download from nodejs.org](https://nodejs.org/)
- **npm**: Comes with Node.js (verify with `npm --version`)
- **Outlook Desktop**: Microsoft 365 subscription (Business or Enterprise)
- **PowerShell**: Windows PowerShell 5.1+ or PowerShell Core 7+
- **AWS CLI**: [Download from AWS](https://aws.amazon.com/cli/) (for S3 deployment)
- **Git**: For version control

### Optional Development Tools
- **Visual Studio Code**: Recommended editor
  - Office Add-in Debugger extension
  - PowerShell extension
  - Live Server extension for local testing
- **Fiddler** or **Postman**: For API testing and debugging

## Architecture Overview

### Application Structure
The add-in follows a modular architecture:

```
TaskpaneApp (main controller)
├── EmailAnalyzer (email content extraction)
├── AIService (multi-provider AI integration)
├── ClassificationDetector (security classification)
├── SettingsManager (persistent configuration)
├── Logger (Windows event logging)
├── UIController (state management)
└── AccessibilityManager (ARIA and keyboard support)
```

### Build Pipeline
- **Webpack 5**: Module bundling and asset processing
- **Entry Points**: Separate bundles for taskpane and commands
- **Asset Processing**: Icon optimization, CSS bundling, HTML template processing, JSON configuration files
- **Source Maps**: Development debugging support
- **Configuration Files**: Automatic copying of `default-providers.json` and `default-models.json` to public directory

## Initial Setup

### 1. Clone and Install Dependencies

```bash
# Clone the repository
git clone <your-repo-url>
cd outlook_email_assistant_v3

# Install dependencies
npm install
```

### 2. Environment Configuration

The project uses environment-specific configurations managed through `tools\deployment-environments.json`:

```json
{
  "Dev": {
    "bucketName": "your-dev-bucket",
    "region": "us-east-1",
    "description": "Development environment"
  },
  "Prd": {
    "bucketName": "your-prod-bucket", 
    "region": "us-east-1",
    "description": "Production environment"
  }
}
```

### 3. Create Required Assets

#### Icon Files
The project requires actual PNG icon files in `src/assets/icons/`:

```
src/assets/icons/
├── icon-16.png   (16x16 pixels)
├── icon-32.png   (32x32 pixels) 
├── icon-80.png   (80x80 pixels)
└── icon-128.png  (128x128 pixels)
```

#### Manifest Configuration  
Update `src/manifest.xml` with your deployment URLs and unique ID:

```xml
<Id>YOUR-UNIQUE-GUID-HERE</Id>
<IconUrl DefaultValue="https://your-bucket.s3-website-region.amazonaws.com/icons/icon-32.png"/>
```

### 4. AI Provider Setup

#### Default Configurations
The add-in includes default configurations in:
- `src/default-providers.json`: AI provider endpoints and labels
- `src/default-models.json`: Default models for each provider

#### Ollama Local Setup (Optional)
For local AI development:

```bash
# Install Ollama
curl -fsSL https://ollama.ai/install.sh | sh

# Pull a model
ollama pull llama3

# Verify service
curl http://localhost:11434/api/tags
```

## Development Workflow

### 1. Local Development

#### Start Development Build Watcher
```bash
npm run dev
```

This command:
- Watches for file changes in `src/`
- Rebuilds automatically using Webpack in development mode
- Outputs to `public/` directory with source maps
- Enables hot reloading for faster development

#### Development Mode Features
- Unminified JavaScript for easier debugging
- Source maps for breakpoint debugging
- Console logging enabled
- Detailed error messages

### 2. Local Testing Options

#### Option A: Direct File Testing
For basic testing, you can serve files directly:
```bash
# Simple Python server (if Python installed)
cd public
python -m http.server 3000

# Or Node.js serve (if installed globally)
npx serve public -p 3000
```

#### Option B: Live Server (Recommended)
Using VS Code Live Server extension:
1. Open the `public` folder in VS Code
2. Right-click on `taskpane.html`
3. Select "Open with Live Server"
4. Access at `http://127.0.0.1:5500/taskpane.html`

### 3. Production Build

Create an optimized production build:
```bash
npm run build
```

Production build includes:
- Minified JavaScript and CSS
- Optimized asset handling
- Source map generation
- Bundle size optimization

### 4. Automated Deployment

Deploy to your configured environment:
```bash
# Deploy to development environment
.\tools\deploy_web_assets.ps1 -Environment Dev

# Deploy to production environment  
.\tools\deploy_web_assets.ps1 -Environment Prd

# Dry run to preview changes
.\tools\deploy_web_assets.ps1 -Environment Dev -DryRun

# Force deployment (skips validation)
.\tools\deploy_web_assets.ps1 -Environment Prd -Force
```

### 5. Manifest Validation

Validate Office Add-in manifest:
```bash
npm run validate-manifest
```

Common validation issues:
- Invalid URLs (must be HTTPS in production)
- Missing or invalid icon files
- Incorrect GUID format
- Schema validation errors

## Sideloading for Testing

### Method 1: Outlook Desktop

1. Open Outlook Desktop
2. Go to **File** > **Manage Add-ins** > **My Add-ins**
3. Click **Add a custom add-in** > **Add from file**
4. Select your `manifest.xml` file
5. Click **Install**

### Method 2: Outlook Web App

1. Open Outlook on the web
2. Click **Settings** > **View all Outlook settings**
3. Go to **General** > **Manage add-ins**
4. Click **Add a custom add-in** > **Add from file**
5. Upload your `manifest.xml` file

## Deployment

### 1. Deploy to S3

```bash
# Build and deploy in one command
npm run build && npm run deploy

# Or deploy with dry run first
pwsh -ExecutionPolicy Bypass -File deploy.ps1 -DryRun
pwsh -ExecutionPolicy Bypass -File deploy.ps1
```

### 2. Update Manifest

After deployment, ensure your `manifest.xml` references the correct S3 URLs.

### 3. Test Deployed Version

1. Update your sideloaded add-in with the new manifest
2. Clear Outlook cache if needed
3. Test all functionality

## Debugging

## Debugging

### Browser Developer Tools

1. Open Outlook Desktop and load your add-in
2. In the taskpane, right-click and select **Inspect Element**
3. Use the developer tools:
   - **Console**: View application logs and errors
   - **Network**: Monitor API calls to AI providers
   - **Sources**: Set breakpoints with source map support
   - **Application**: Inspect localStorage and Office.js roaming settings

### Visual Studio Code Debugging

1. Install the "Office Add-in Debugger" extension
2. Create `.vscode/launch.json`:

```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "name": "Debug Office Add-in",
      "type": "office-addin",
      "request": "attach",
      "url": "https://your-domain.com/taskpane.html",
      "port": 9229
    }
  ]
}
```

### PowerShell and Event Log Debugging

Monitor Windows Application Log for add-in events:
```powershell
# View recent PromptEmail events
Get-EventLog -LogName Application -Source "PromptEmail" -Newest 10

# Monitor in real-time
Get-EventLog -LogName Application -Source "PromptEmail" -After (Get-Date).AddMinutes(-5)
```

### AI Service Debugging

#### Test AI Provider Connections
```javascript
// In browser console, test AI service manually
const aiService = new AIService();

// Test Ollama connection
aiService.callAI('Hello, test message', {
  provider: 'ollama',
  model: 'llama3',
  baseUrl: 'http://localhost:11434'
});
```

#### Common AI Integration Issues
- **CORS errors**: Verify provider supports browser requests
- **API key issues**: Check key format and permissions  
- **Model availability**: Verify model exists for provider
- **Rate limiting**: Monitor API quotas and limits

## Troubleshooting

### Critical Office Configuration Issues

#### Add-in Button Not Appearing (Most Common Issue)

**❗ CRITICAL: Optional Connected Experiences**
The #1 cause of add-in buttons not appearing is disabled "Optional Connected Experiences":

1. **Location**: File → Options → Trust Center → Trust Center Settings → Privacy Options
2. **Setting**: "Optional connected experiences" must be **ENABLED**
3. **Impact**: When disabled, ALL web-based add-ins will fail to load with no error messages
4. **Solution**: Re-enable the setting and restart Outlook

**Other Office Settings That Break Add-ins:**

1. **Add-in Trust Settings**:
   - Location: File → Options → Trust Center → Trust Center Settings → Add-ins
   - Uncheck "Require Application Add-ins to be signed by Trusted Publisher" (for development)
   - Uncheck "Disable all Application Add-ins"

2. **Internet Zone Security**:
   - Overly restrictive internet security can block external URLs
   - Add your S3 domain to trusted sites if needed

3. **Macro Security Settings**:
   - Can sometimes interfere with add-in initialization
   - Set to "Disable all macros with notification" or less restrictive

### Common Development Issues

#### Add-in Not Loading
1. **Manifest validation errors**:
   ```bash
   npm run validate-manifest
   ```

2. **Certificate/HTTPS issues**:
   - Ensure all URLs in manifest are HTTPS (except localhost)
   - Check SSL certificate validity
   - Clear browser certificates if needed

3. **Office.js API errors**:
   ```javascript
   // Check Office.js is loaded
   console.log('Office loaded:', typeof Office !== 'undefined');
   
   // Verify context
   console.log('Mailbox context:', Office.context?.mailbox?.item);
   ```

#### Build and Deployment Issues  
1. **Webpack build failures**:
   ```bash
   # Clear node_modules and reinstall
   rm -rf node_modules package-lock.json
   npm install
   
   # Check for conflicting dependencies
   npm audit
   ```

2. **PowerShell execution policy**:
   ```powershell
   # Enable script execution
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **AWS S3 deployment errors**:
   ```bash
   # Verify AWS CLI configuration
   aws configure list
   
   # Test S3 access
   aws s3 ls s3://your-bucket-name
   ```

#### Runtime Issues
1. **Settings not persisting**:
   - Check Office.js roaming settings availability
   - Verify localStorage fallback is working
   - Check browser storage quotas

2. **AI provider connection failures**:
   - Verify API keys are correctly stored
   - Check network connectivity and CORS settings
   - Test endpoints directly with curl/Postman

3. **Classification detection not working**:
   - Verify email content format
   - Check pattern matching in `ClassificationDetector.js`
   - Review console logs for detection results

### Cache Clearing Procedures

#### Clear Outlook Add-in Cache
```powershell
# Clear Office Web Extensions cache
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\*" -Recurse -Force

# Clear browser cache for Office (manual process)
# 1. Go to Edge/Chrome settings
# 2. Clear browsing data for office.com domain
# 3. Restart Outlook
```

#### Clear Development Assets
```bash
# Clean build directory
rm -rf public/*

# Clear npm cache
npm cache clean --force

# Rebuild from scratch
npm run build
```

## Code Architecture Deep Dive

### Core Application Class (`TaskpaneApp`)

The main application controller manages the entire add-in lifecycle:

```javascript
class TaskpaneApp {
  constructor() {
    // Initialize all service dependencies
    this.emailAnalyzer = new EmailAnalyzer();
    this.aiService = new AIService();  
    this.classificationDetector = new ClassificationDetector();
    this.logger = new Logger();
    this.settingsManager = new SettingsManager();
    this.accessibilityManager = new AccessibilityManager();
    this.uiController = new UIController();
  }

  async initialize() {
    // Full initialization sequence
    await this.initializeOffice();
    await this.settingsManager.loadSettings();
    this.setupUI();
    this.accessibilityManager.initialize();
    await this.loadCurrentEmail();
  }
}
```

### Service Layer Architecture

#### EmailAnalyzer Service
- Extracts email content using Office.js API
- Handles both read and compose modes
- Parses recipients, attachments, and metadata
- Provides email context for AI analysis

#### AIService Architecture  
- **Multi-provider support**: OpenAI-compatible, Ollama, Custom on-site providers
- **Dynamic model discovery**: Automatic Ollama model detection
- **Response parsing**: Provider-specific response handling
- **Error handling**: Comprehensive error recovery and user feedback

#### SettingsManager Persistence Strategy
- **Primary storage**: Office.js roamingSettings for cross-device sync
- **Fallback storage**: localStorage for offline scenarios  
- **Data validation**: Schema validation for settings integrity
- **Change notification**: Event-driven settings updates

### Security Implementation

#### ClassificationDetector Patterns
```javascript
// Pattern matching for various classification formats
this.patterns = [
  /^(UNCLASSIFIED|CONFIDENTIAL|SECRET|TOP SECRET|TS)$/gim,
  /\(([UCS]|CONFIDENTIAL|SECRET|TOP SECRET|TS)\)/gim
];
```

#### Data Sanitization
```javascript
sanitizeData(data) {
  const sensitiveFields = [
    'apiKey', 'password', 'token', 'secret',
    'emailBody', 'content', 'personalInfo'
  ];
  // Remove sensitive data from logs
}
```

### UI State Management

#### Loading States
The UIController manages complex loading states:
```javascript
setButtonLoading(buttonId, isLoading) {
  // Store original state, show spinner, update ARIA labels
}
```

#### Accessibility Implementation  
```javascript
// ARIA live regions for screen reader announcements
setupAriaLive() {
  const announceRegion = document.createElement('div');
  announceRegion.setAttribute('aria-live', 'assertive');
  announceRegion.setAttribute('aria-atomic', 'true');
}
```

## Testing Strategies

### Unit Testing Approach

While the project doesn't currently include unit tests, here's the recommended testing structure:

#### Service Layer Testing
```javascript
// Test AI service provider switching
describe('AIService', () => {
  it('should switch providers correctly', async () => {
    const aiService = new AIService();
    const response = await aiService.analyzeEmail(mockEmail, {
      provider: 'ollama',
      model: 'llama3'
    });
    expect(response).toBeDefined();
  });
});
```

#### Email Analysis Testing
```javascript
describe('EmailAnalyzer', () => {
  it('should extract email content correctly', async () => {
    const analyzer = new EmailAnalyzer();
    // Mock Office.js context
    const mockEmail = await analyzer.getCurrentEmail();
    expect(mockEmail.subject).toBeDefined();
  });
});
```

### Integration Testing

#### Office.js API Testing
```javascript
// Test Office.js integration in browser console
Office.onReady(() => {
  console.log('Host:', Office.context.host);
  console.log('Platform:', Office.context.platform);
  console.log('Requirements:', Office.context.requirements.sets);
});
```

#### End-to-End Testing Workflow
1. **Manual Testing Checklist**:
   - Load add-in in Outlook Desktop
   - Test both read and compose modes
   - Verify ribbon button functionality
   - Test all AI providers with sample content
   - Verify accessibility with keyboard navigation
   - Test classification detection with sample content

2. **Automated Testing Setup** (recommended):
   - Use Playwright or Cypress for E2E testing
   - Mock Office.js API for consistent testing
   - Test deployment pipeline with staging environment

### Performance Testing

#### Bundle Analysis  
```bash
# Generate bundle analysis report
npm run build -- --analyze
```

#### Memory Profiling
```javascript
// Monitor memory usage during AI operations
const beforeMemory = performance.memory.usedJSHeapSize;
await aiService.analyzeEmail(emailData, config);
const afterMemory = performance.memory.usedJSHeapSize;
console.log('Memory used:', (afterMemory - beforeMemory) / 1024 / 1024, 'MB');
```

## Code Standards and Contributing

### 1. Code Style Guidelines

#### JavaScript Style
- **ES6+ Features**: Use modern JavaScript features (classes, modules, async/await)
- **Variable Naming**: Use camelCase for variables and functions, PascalCase for classes
- **Constants**: Use UPPER_SNAKE_CASE for constants
- **Destructuring**: Use destructuring for object properties when appropriate
- **Arrow Functions**: Prefer arrow functions for callbacks and short functions

#### Code Organization
```javascript
// File structure example
class ServiceName {
  constructor() {
    // Initialize properties
    this.property = value;
  }

  /**
   * Method documentation with JSDoc
   * @param {string} parameter - Parameter description
   * @returns {Promise<Object>} Return value description
   */
  async methodName(parameter) {
    // Implementation
  }

  // Private methods prefixed with underscore
  _privateMethod() {
    // Private implementation
  }
}
```

#### Error Handling Standards
```javascript
// Consistent error handling pattern
async performOperation() {
  try {
    const result = await this.riskyOperation();
    return result;
  } catch (error) {
    console.error('Operation failed:', error);
    this.logger.logEvent('operation_failed', { error: error.message });
    throw new Error('User-friendly error message');
  }
}
```

#### Logging Standards
```javascript
// Logging levels and structure
this.logger.logEvent('event_type', {
  // Always include
  timestamp: new Date().toISOString(),
  userId: this.getCurrentUserId(),
  sessionId: this.sessionId,
  
  // Event-specific data (sanitized)
  // Never include: email content, API keys, personal data
});
```

### 2. Development Workflow

#### Branch Management
```bash
# Feature development workflow
git checkout -b feature/description-of-feature
# Make changes
git add .
git commit -m "feat: description of changes"
git push origin feature/description-of-feature
# Create pull request
```

#### Commit Message Format
Follow conventional commit format:
```
type(scope): description

feat: add new AI provider support
fix: resolve classification detection issue  
docs: update deployment guide
style: fix code formatting
refactor: reorganize service architecture
test: add unit tests for EmailAnalyzer
chore: update dependencies
```

#### Code Review Checklist
- [ ] Code follows established style guidelines
- [ ] All functions have JSDoc documentation
- [ ] Error handling is comprehensive and user-friendly
- [ ] No sensitive data in logs or client-side code
- [ ] Accessibility considerations addressed
- [ ] Performance impact considered
- [ ] Browser compatibility maintained
- [ ] Security implications reviewed

### 3. Testing Requirements

#### Unit Testing (Recommended Implementation)
```javascript
// Example test structure
describe('AIService', () => {
  let aiService;
  
  beforeEach(() => {
    aiService = new AIService();
  });

  describe('analyzeEmail', () => {
    it('should return valid analysis for normal email', async () => {
      const mockEmail = {
        subject: 'Test Subject',
        body: 'Test content',
        from: 'test@example.com'
      };
      
      const result = await aiService.analyzeEmail(mockEmail, {
        provider: 'mock',
        model: 'test-model'
      });
      
      expect(result).toHaveProperty('sentiment');
      expect(result).toHaveProperty('summary');
    });
    
    it('should handle API failures gracefully', async () => {
      // Test error handling
    });
  });
});
```

#### Integration Testing
```javascript
// Office.js integration testing
describe('Office Integration', () => {
  beforeAll(async () => {
    // Mock Office.js environment
    global.Office = {
      context: {
        mailbox: {
          item: mockMailboxItem,
          userProfile: mockUserProfile
        }
      },
      onReady: jest.fn()
    };
  });
  
  it('should extract email content correctly', async () => {
    const emailAnalyzer = new EmailAnalyzer();
    const email = await emailAnalyzer.getCurrentEmail();
    expect(email.subject).toBeDefined();
  });
});
```

### 4. Documentation Standards

#### JSDoc Comments
```javascript
/**
 * Analyzes email content using configured AI provider
 * @param {Object} emailData - Email data from EmailAnalyzer
 * @param {string} emailData.subject - Email subject line
 * @param {string} emailData.body - Email body content  
 * @param {string} emailData.from - Sender email address
 * @param {Object} config - AI configuration object
 * @param {string} config.provider - AI provider name ('openai', 'ollama', etc.)
 * @param {string} config.model - Model name for the provider
 * @param {string} [config.apiKey] - API key (optional for some providers)
 * @returns {Promise<Object>} Analysis results
 * @throws {Error} When AI provider is unavailable or invalid
 */
async analyzeEmail(emailData, config) {
  // Implementation
}
```

#### README Updates
When adding new features, update relevant documentation:
- Feature description in main README.md
- Development setup instructions in DEVELOPER_GUIDE.md
- Deployment considerations in DEPLOYMENT_GUIDE.md
- API changes in code comments

### 5. Security Guidelines

#### Client-Side Security
- Never store API keys or secrets in client-side code
- Sanitize all user inputs before processing
- Validate data types and ranges for all parameters
- Use Content Security Policy headers where possible

#### Data Privacy
```javascript
// Data sanitization example
sanitizeForLogging(data) {
  const sanitized = { ...data };
  
  // Remove sensitive fields
  const sensitiveFields = [
    'apiKey', 'password', 'token', 'secret',
    'emailBody', 'content', 'personalInfo'
  ];
  
  sensitiveFields.forEach(field => {
    if (sanitized[field]) {
      delete sanitized[field];
    }
  });
  
  return sanitized;
}
```

#### Classification Handling
- Always warn users before processing classified content
- Log user override decisions for audit purposes
- Implement proper access controls for sensitive features
- Follow organizational security policies

### 6. Performance Guidelines

#### Code Optimization
- Minimize bundle size by avoiding unnecessary dependencies
- Use dynamic imports for large optional features
- Implement efficient caching strategies
- Optimize DOM manipulations and event handlers

#### Memory Management
```javascript
// Proper cleanup example
class ResourceManager {
  constructor() {
    this.resources = new Map();
  }
  
  addResource(id, resource) {
    this.resources.set(id, resource);
  }
  
  cleanup() {
    // Clean up resources to prevent memory leaks
    this.resources.forEach((resource) => {
      if (resource.cleanup) {
        resource.cleanup();
      }
    });
    this.resources.clear();
  }
}
```

### 7. Accessibility Requirements

#### WCAG 2.1 AA Compliance
- Ensure all interactive elements are keyboard accessible
- Provide appropriate ARIA labels and roles
- Maintain sufficient color contrast ratios (4.5:1 minimum)
- Test with screen readers (NVDA, JAWS, VoiceOver)

#### Implementation Examples
```javascript
// Accessible button creation
createAccessibleButton(text, action, ariaLabel = null) {
  const button = document.createElement('button');
  button.textContent = text;
  button.setAttribute('aria-label', ariaLabel || text);
  button.addEventListener('click', action);
  button.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      action();
    }
  });
  return button;
}
```

### 8. Deployment Best Practices

#### Environment Management
- Use environment-specific configurations
- Never commit credentials or secrets
- Test thoroughly in staging environment before production
- Implement proper rollback procedures

#### Version Management
```json
// package.json version strategy
{
  "version": "1.2.3",
  // Major.Minor.Patch
  // Major: Breaking changes
  // Minor: New features (backward compatible)  
  // Patch: Bug fixes
}
```

#### Release Notes
Maintain clear release notes for each deployment:
```markdown
## Version 1.2.3 (2025-01-15)

### Added
- Enhanced Ollama local model support
- Enhanced accessibility features

### Fixed  
- Classification detection for multi-line headers
- Memory leak in AI service cleanup

### Security
- Updated dependencies with security patches
```

This comprehensive guide should help maintain code quality and consistency across the project while ensuring security, accessibility, and performance standards are met.

## File Structure Reference

```
outlook_email_assistant_v3/
├── public/                 # Build output (generated)
├── src/
│   ├── assets/
│   │   ├── css/           # Stylesheets
│   │   └── icons/         # Icon files
│   ├── commands/          # Ribbon command handlers
│   ├── services/          # Business logic services
│   ├── taskpane/          # Main taskpane code
│   └── ui/               # UI components
├── manifest.xml          # Office Add-in manifest
├── deploy.ps1           # Deployment script
├── package.json         # Project configuration
└── webpack.config.js    # Build configuration
```

## API Configuration

### Supported AI Services

- **OpenAI-Compatible**: Set service to 'openai', provide API key
- **Ollama**: Set service to 'ollama', typically no API key required for local installations
- **Custom On-Site Providers**: Set service to 'custom', provide endpoint URL and credentials as needed

### Environment Variables (Optional)

Create a `.env` file for default settings:
```
DEFAULT_AI_SERVICE=openai
DEFAULT_MODEL=gpt-4
LOG_LEVEL=info
```

## Security Notes

- Never commit API keys to version control
- Use environment variables for sensitive configuration
- Test classification detection with sample classified emails
- Verify logging excludes sensitive content

## Next Steps

1. Complete icon design and creation
2. Set up AWS S3 bucket and deployment
3. Configure AI service credentials
4. Test with real emails in Outlook
5. Deploy to production S3 bucket
6. Share manifest with users for installation
