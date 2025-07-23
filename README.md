# PromptEmail Outlook Add-in

AI-Powered Email Analysis Outlook Add-in that enhances email productivity through large language model integration.

## Features

- **AI-Powered Email Analysis**: Intelligent email classification and response assistance
- **Security-First Design**: Built-in classification detection and compliance logging
- **Accessible Interface**: Keyboard navigation and screen reader support
- **Configurable AI Models**: Support for self-hosted and cloud LLM endpoints
- **Persistent Preferences**: Settings saved across sessions

## Quick Start

### Prerequisites

- Node.js 16+ and npm
- Outlook Desktop (Microsoft 365)
- AWS CLI configured (for deployment)

### Installation

1. Clone the repository and install dependencies:
```bash
npm install
```

2. Build the project:
```bash
npm run build
```

3. Configure your AWS S3 bucket in `deploy.ps1`

4. Deploy to S3:
```bash
npm run deploy
```

5. Sideload the manifest (`manifest.xml`) in Outlook

### Development

- **Start development server**: `npm run dev`
- **Build for production**: `npm run build`
- **Validate manifest**: `npm run validate-manifest`

## Architecture

- **Frontend**: Vanilla JavaScript + HTML/CSS
- **Hosting**: AWS S3 static website hosting
- **Integration**: Office Add-in manifest for Outlook ribbon and taskpane
- **Logging**: PowerShell-based Windows Application Log integration

## Security & Compliance

- Email classification detection (UNCLASSIFIED, SECRET, etc.)
- User override warnings with audit logging
- Secure API key storage
- No sensitive content in telemetry

## Directory Structure

```
├── public/                 # Static assets and HTML
├── src/                   # Source code
│   ├── services/          # Modular service layer
│   ├── ui/               # UI components and controls
│   └── utils/            # Utility functions
├── manifest.xml          # Office Add-in manifest
├── deploy.ps1           # Deployment script
└── webpack.config.js    # Build configuration
```

## Deployment

See `DEPLOYMENT_GUIDE.md` for detailed deployment instructions.

## Contributing

1. Create a feature branch
2. Make your changes
3. Test thoroughly
4. Submit a pull request

## License

MIT License - see LICENSE file for details.
