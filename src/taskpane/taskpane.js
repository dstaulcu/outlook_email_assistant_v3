// PromptEmail Taskpane JavaScript
// Main application logic for the email analysis interface

import '../assets/css/taskpane.css';
import { EmailAnalyzer } from '../services/EmailAnalyzer';
import { AIService } from '../services/AIService';
import { ClassificationDetector } from '../services/ClassificationDetector';
import { Logger } from '../services/Logger';
import { SettingsManager } from '../services/SettingsManager';
import { AccessibilityManager } from '../ui/AccessibilityManager';
import { UIController } from '../ui/UIController';

class TaskpaneApp {
    async fetchDefaultProvidersConfig() {
        // Fetch ai-providers.json from public config directory
        try {
            const response = await fetch('/config/ai-providers.json');
            if (!response.ok) throw new Error('Failed to fetch ai-providers.json');
            return await response.json();
        } catch (e) {
            console.warn('[Default Providers] Could not load ai-providers.json:', e);
            return {};
        }
    }

    switchToResponseTab() {
        // Switch to the response tab in the UI
        const responseTabButton = document.querySelector('.tab-button[aria-controls="panel-response"]');
        if (responseTabButton) {
            console.debug('Switching to response tab');
            responseTabButton.click();
        } else {
            console.error('Response tab button not found');
        }
    }
    showResponseSection() {
        // Show the response section in the UI
        const responseSection = document.getElementById('response-section');
        if (responseSection) {
            responseSection.classList.remove('hidden');
        }
    }
    constructor() {
        // Add unique instance ID for debugging
        this.instanceId = Date.now() + '-' + Math.random().toString(36).substr(2, 9);
        console.debug('TaskpaneApp instance created:', this.instanceId);
        
        this.emailAnalyzer = new EmailAnalyzer();
        this.aiService = new AIService();
        this.classificationDetector = new ClassificationDetector();
        this.logger = new Logger();
        this.settingsManager = new SettingsManager();
        this.accessibilityManager = new AccessibilityManager();
        this.uiController = new UIController();
        
        this.currentEmail = null;
        this.currentAnalysis = null;
        this.currentResponse = null;
        this.pendingAction = null; // Track what action user wants after warning override
        this.classificationOverrideGranted = false; // Track if user has already overridden classification for current email
        this.sessionStartTime = Date.now();
        
        // Telemetry tracking properties
        this.refinementCount = 0;
        this.hasUsedClipboard = false;

        // Model selection UI elements
        this.modelServiceSelect = null;
        this.modelSelectGroup = null;
        this.modelSelect = null;
        
        // Set up session end tracking
        this.setupSessionTracking();
    }

    setupSessionTracking() {
        // Track when user navigates away or closes the taskpane
        window.addEventListener('beforeunload', () => {
            this.logSessionSummary();
        });
        
        // Track when the taskpane loses focus (user switches to another part of Outlook)
        window.addEventListener('blur', () => {
            // Log session summary with a slight delay to allow for quick focus changes
            setTimeout(() => {
                if (!document.hasFocus()) {
                    this.logSessionSummary();
                }
            }, 1000);
        });
    }

    logSessionSummary() {
        if (this.sessionSummaryLogged) return; // Prevent duplicate logging
        this.sessionSummaryLogged = true;
        
        const sessionDuration = Date.now() - this.sessionStartTime;
        this.logger.logEvent('session_summary', {
            session_duration_ms: sessionDuration,
            refinement_count: this.refinementCount,
            clipboard_used: this.hasUsedClipboard,
            email_analyzed: this.currentEmail !== null,
            response_generated: this.currentResponse !== null
        }, 'Information', this.getUserEmailForTelemetry());
    }

    async initialize() {
        try {
            // Initialize Office.js
            await this.initializeOffice();
            
            // Load user settings
            await this.settingsManager.loadSettings();
            
            // Apply accessibility settings immediately after loading
            const currentSettings = await this.settingsManager.getSettings();
            if (currentSettings['high-contrast']) {
                console.debug('Applying high contrast during initialization');
                this.toggleHighContrast(true);
            }
            
            // Load provider config before UI setup
            this.defaultProvidersConfig = await this.fetchDefaultProvidersConfig();
            
            // Update AIService with provider configuration
            this.aiService.updateProvidersConfig(this.defaultProvidersConfig);
            
            // Setup UI
            await this.setupUI();
            
            // Update version display
            this.updateVersionDisplay();
            
            // Setup accessibility
            this.accessibilityManager.initialize();
            
            // Initialize Splunk telemetry if enabled
            await this.initializeTelemetry();
            
            // Load current email
            await this.loadCurrentEmail();
            
            // Check if user needs initial setup (first time user or missing API key)
            await this.checkForInitialSetupNeeded();
            
            // Try automatic analysis if conditions are met
            await this.attemptAutoAnalysis();
            
            // Hide loading, show main content
            this.uiController.hideLoading();
            this.uiController.showMainContent();
            
            // Log session start
            this.logger.logEvent('session_start', {
                host: Office.context.host
            });
            
        } catch (error) {
            console.error('Failed to initialize TaskpaneApp:', error);
            this.uiController.showError('Failed to initialize application. Please try again.');
        }
    }

    async initializeOffice() {
        return new Promise((resolve, reject) => {
            Office.onReady((info) => {
                if (info.host === Office.HostType.Outlook) {
                    // Cache user context immediately when Office is ready
                    this.logger.cacheUserContext();
                    resolve();
                } else {
                    reject(new Error('This add-in is designed for Outlook only'));
                }
            });
        });
    }

    async initializeTelemetry() {
        console.debug('Initializing telemetry...');
        
        try {
            // Logger already initialized telemetry config in constructor, just check if it's ready
            // If not initialized yet, wait for it
            if (!this.logger.telemetryConfig) {
                console.debug('Waiting for telemetry config to load...');
                // Give it a moment to load
                await new Promise(resolve => setTimeout(resolve, 100));
            }
            
            // Start telemetry auto-flush if enabled
            if (this.logger.telemetryConfig?.telemetry?.enabled) {
                const provider = this.logger.telemetryConfig.telemetry.provider;
                if (provider === 'api_gateway') {
                    this.logger.startApiGatewayAutoFlush();
                    console.info(`${provider} telemetry enabled and auto-flush started`);
                }
            }
            
        } catch (error) {
            console.error('Failed to initialize telemetry:', error);
        }
    }

    async setupUI() {
        // Bind event listeners
        this.bindEventListeners();
        
        // Initialize sliders
        this.initializeSliders();
        
        // Setup tabs
        this.initializeTabs();
        
        // Model selection UI elements
        this.modelServiceSelect = document.getElementById('model-service');
        this.modelSelectGroup = document.getElementById('model-select-group');
        this.modelSelect = document.getElementById('model-select');
        
        // Populate model service dropdown from defaultProvidersConfig BEFORE loading settings
        if (this.modelServiceSelect && this.defaultProvidersConfig) {
            this.modelServiceSelect.innerHTML = Object.entries(this.defaultProvidersConfig)
                .filter(([key, val]) => key !== 'custom' && key !== '_config')
                .map(([key, val]) => `<option value="${key}">${val.label}</option>`)
                .join('');
        }
        
        // Load settings into UI (this will now properly select the saved model-service value)
        await this.loadSettingsIntoUI();
        // Hide AI config placeholder in main UI by default
        const aiConfigPlaceholder = document.getElementById('ai-config-placeholder');
        if (aiConfigPlaceholder) {
            aiConfigPlaceholder.classList.add('hidden');
            aiConfigPlaceholder.innerHTML = '';
        }
        if (this.modelServiceSelect && this.modelSelectGroup && this.modelSelect) {
            // Set initial baseUrl to console
            if (this.modelServiceSelect.value && this.defaultProvidersConfig && this.defaultProvidersConfig[this.modelServiceSelect.value]) {
                const baseUrl = this.defaultProvidersConfig[this.modelServiceSelect.value].baseUrl || '';
                console.debug(`Model provider: ${this.modelServiceSelect.value}, base URL: ${baseUrl}`);
            }
            await this.updateModelDropdown();
        }
    }

    bindEventListeners() {
        // Main action buttons
        document.getElementById('analyze-email').addEventListener('click', () => this.analyzeEmail());
        document.getElementById('generate-response').addEventListener('click', () => this.generateResponse());
        document.getElementById('refine-response').addEventListener('click', () => this.refineResponse());

        // Classification warning buttons
        document.getElementById('proceed-anyway').addEventListener('click', () => this.proceedWithWarning());
        document.getElementById('cancel-analysis').addEventListener('click', () => this.cancelAnalysis());

        // Response actions
        document.getElementById('copy-response').addEventListener('click', () => this.copyResponse());
        
        // Settings
        document.getElementById('open-settings').addEventListener('click', () => this.openSettings());
        document.getElementById('close-settings').addEventListener('click', () => this.closeSettings());
        document.getElementById('reset-settings').addEventListener('click', () => this.resetSettings());
        
        // Help and navigation links
        document.getElementById('api-key-help-btn').addEventListener('click', () => this.showProviderHelp());
        document.getElementById('source-link').addEventListener('click', (e) => this.openSource(e));
        document.getElementById('issues-link').addEventListener('click', (e) => this.openIssues(e));
        document.getElementById('wiki-link').addEventListener('click', (e) => this.openWiki(e));
        document.getElementById('telemetry-link').addEventListener('click', (e) => this.openTelemetry(e));
        
        // Model service change
        document.getElementById('model-service').addEventListener('change', (e) => this.onModelServiceChange(e));
        
        // Settings checkboxes
        document.getElementById('high-contrast').addEventListener('change', (e) => this.toggleHighContrast(e.target.checked));
        document.getElementById('screen-reader-mode').addEventListener('change', (e) => this.toggleScreenReaderMode(e.target.checked));
        
        // Auto-save settings with special handling for provider-specific fields
        ['custom-instructions'].forEach(id => {
            const element = document.getElementById(id);
            if (element) {
                element.addEventListener('blur', () => this.saveSettings());
            }
        });
        
        // Special handling for provider-specific fields (API key and endpoint URL)
        ['api-key', 'endpoint-url'].forEach(id => {
            const element = document.getElementById(id);
            if (element) {
                element.addEventListener('blur', () => {
                    this.saveCurrentProviderSettings(this.modelServiceSelect?.value);
                    // Also trigger model lookup when endpoint or key changes
                    this.updateModelDropdown();
                });
            }
        });
    }

    /**
     * Update version display with dynamic version from package.json
     */
    updateVersionDisplay() {
        const versionDisplay = document.getElementById('version-display');
        if (versionDisplay) {
            const version = process.env.PACKAGE_VERSION || '1.0.0';
            versionDisplay.textContent = `v${version}`;
        }
    }

    initializeSliders() {
        const sliders = [
            { id: 'response-length', values: ['Very Brief', 'Brief', 'Medium', 'Detailed', 'Very Detailed'] },
            { id: 'response-tone', values: ['Very Casual', 'Casual', 'Professional', 'Formal', 'Very Formal'] },
            { id: 'response-urgency', values: ['Very Relaxed', 'Relaxed', 'Normal', 'Urgent', 'Very Urgent'] }
        ];

        sliders.forEach(({ id, values }) => {
            const slider = document.getElementById(id);
            const valueDisplay = document.getElementById(id.replace('response-', '') + '-value');
            
            slider.addEventListener('input', (e) => {
                const value = parseInt(e.target.value) - 1;
                valueDisplay.textContent = values[value];
                this.saveSettings();
            });
        });
    }

    initializeTabs() {
        const tabButtons = document.querySelectorAll('.tab-button');
        const tabPanels = document.querySelectorAll('.tab-panel');
        
        tabButtons.forEach(button => {
            button.addEventListener('click', (e) => {
                const targetPanel = e.target.getAttribute('aria-controls');
                
                // Update buttons
                tabButtons.forEach(btn => {
                    btn.classList.remove('active');
                    btn.setAttribute('aria-selected', 'false');
                });
                
                // Update panels
                tabPanels.forEach(panel => {
                    panel.classList.remove('active');
                });
                
                // Activate current
                e.target.classList.add('active');
                e.target.setAttribute('aria-selected', 'true');
                document.getElementById(targetPanel).classList.add('active');
            });
        });
    }

    async loadCurrentEmail() {
        try {
            this.currentEmail = await this.emailAnalyzer.getCurrentEmail();
            this.classificationOverrideGranted = false; // Reset override flag for new email
            
            // Ensure context is properly stored on currentEmail for later use
            if (this.currentEmail && this.currentEmail.context) {
                console.debug('Email context loaded:', this.currentEmail.context);
            }
            
            await this.displayEmailSummary(this.currentEmail);
        } catch (error) {
            console.error('Failed to load current email:', error);
            this.uiController.showError('Failed to load email. Please select an email and try again.');
        }
    }

    async checkForInitialSetupNeeded(showSettingsIfNeeded = true) {
        try {
            const currentSettings = await this.settingsManager.getSettings();
            const selectedService = currentSettings['model-service'] || 'onsite1'; // Default provider
            
            // Check if user has an API key configured for the default/selected provider
            const providerConfigs = currentSettings['provider-configs'] || {};
            const selectedProviderConfig = providerConfigs[selectedService] || {};
            const apiKey = selectedProviderConfig['api-key'] || currentSettings['api-key'] || '';
            
            // Also check if this appears to be a first-time user (no last-updated timestamp)
            const isFirstTime = !currentSettings['last-updated'];
            
            // If no API key is set for the selected provider, or if it's a first-time user
            if (!apiKey.trim() || isFirstTime) {
                if (showSettingsIfNeeded) {
                    console.info('Initial setup needed - showing settings tab');
                    
                    // Show a welcome message for first-time users
                    if (isFirstTime) {
                        this.uiController.showStatus('Welcome! Please configure your AI provider settings to get started.');
                    } else {
                        this.uiController.showStatus('API key required. Please configure your API key in settings.');
                    }
                    
                    // Switch to settings tab
                    const settingsTab = document.querySelector('button[data-tab="settings"]');
                    if (settingsTab) {
                        settingsTab.click();
                        
                        // Highlight the API key field if it exists
                        setTimeout(() => {
                            const apiKeyField = document.getElementById('api-key');
                            if (apiKeyField) {
                                apiKeyField.style.borderColor = '#007bff';
                                apiKeyField.style.borderWidth = '2px';
                                apiKeyField.style.boxShadow = '0 0 0 0.2rem rgba(0, 123, 255, 0.25)';
                                apiKeyField.focus();
                                
                                // Add a helpful tooltip or message
                                let helpDiv = document.getElementById('api-key-help');
                                if (!helpDiv) {
                                    helpDiv = document.createElement('div');
                                    helpDiv.id = 'api-key-help';
                                    helpDiv.style.cssText = 'color: #007bff; font-size: 14px; margin-top: 5px; padding: 8px; background-color: #e7f3ff; border-radius: 4px; border: 1px solid #b8daff;';
                                    helpDiv.innerHTML = '💡 Enter your API key here to start using the AI assistant. You can get your API key from your AI service provider.';
                                    apiKeyField.parentNode.appendChild(helpDiv);
                                }
                                
                                // Remove highlight after user starts typing
                                const removeHighlight = () => {
                                    apiKeyField.style.borderColor = '';
                                    apiKeyField.style.borderWidth = '';
                                    apiKeyField.style.boxShadow = '';
                                    if (helpDiv) {
                                        helpDiv.remove();
                                    }
                                    apiKeyField.removeEventListener('input', removeHighlight);
                                };
                                apiKeyField.addEventListener('input', removeHighlight);
                            }
                        }, 500);
                    }
                    
                    // Log this event for analytics
                    this.logger.logEvent('initial_setup_prompted', {
                        selected_service: selectedService,
                        has_api_key: !!apiKey.trim(),
                        is_first_time: isFirstTime
                    }, 'Information', this.getUserEmailForTelemetry());
                }
                
                return true; // Indicates setup is needed
            }
            
            return false; // No setup needed
        } catch (error) {
            console.error('Error checking for initial setup:', error);
            return false; // Continue normally on error
        }
    }

    async displayEmailSummary(email) {
        console.debug('Displaying email summary:', email);
        
        // Email overview section has been removed for cleaner UI
        // Classification checks still run in background for AI compatibility
        
        // Enhanced classification detection for AI provider compatibility
        let classificationResult = null;
        
        if (email.body) {
            classificationResult = this.classificationDetector.detectClassification(email.body);
            console.debug('Classification result:', classificationResult);
        }
        
        // Store classification result for later use
        if (classificationResult) {
            email.classificationResult = classificationResult;
        }

        // Context-aware UI adaptation (works behind the scenes)
        this.adaptUIForContext(email.context);
    }

    /**
     * Adapts the UI based on email context (sent vs inbox vs compose)
     * @param {Object} context - Context information from EmailAnalyzer
     */
    adaptUIForContext(context) {
        console.debug('Adapting UI for context:', context);
        
        if (!context) {
            console.warn('No context provided for UI adaptation');
            return;
        }

        // Log detailed context information for debugging
        console.debug('Context details:', {
            isSentMail: context.isSentMail,
            isInbox: context.isInbox,
            isCompose: context.isCompose,
            folderType: context.folderType,
            debugInfo: context.debugInfo
        });
        
        // Log telemetry for context detection
        this.logger.logEvent('email_context_detected', {
            context_type: context.isSentMail ? 'sent' : (context.isCompose ? 'compose' : 'inbox'),
            folder_detection_available: context.debugInfo ? context.debugInfo.folderDetectionAttempted : false,
            folder_api_result: context.debugInfo ? context.debugInfo.folderApiResult : 'unknown',
            detection_method: context.debugInfo ? context.debugInfo.detectionMethod : 'unknown',
            email_comparison_used: context.debugInfo ? context.debugInfo.emailComparisonUsed : false,
            folder_type: context.folderType || 'unknown'
        }, 'Information', this.getUserEmailForTelemetry());
        
        try {
            // Get UI elements that need adaptation
            const analysisSection = document.getElementById('panel-analysis');
            const responseSection = document.getElementById('panel-response');
            
            // Find buttons and UI elements for context-aware behavior
            const analyzeButton = document.getElementById('analyze-email');
            const generateResponseButton = document.getElementById('generate-response');
            
            // Apply context-specific adaptations (affects button behavior)
            if (context.isCompose) {
                console.debug('Applying compose mode UI adaptations');
                this.adaptUIForComposeMode();
            } else if (context.isSentMail) {
                console.debug('Applying sent mail UI adaptations');
                this.adaptUIForSentMail();
            } else {
                console.debug('Applying inbox mail UI adaptations');
                this.adaptUIForInboxMail();
            }

        } catch (error) {
            console.error('Error adapting UI for context:', error);
        }
    }

    /**
     * Adapts UI for compose mode (writing new email)
     */
    adaptUIForComposeMode() {
        console.debug('Adapting UI for compose mode');
        
        // Hide analysis features since we're composing
        this.setElementVisibility('analyze-email', false);
        this.setElementVisibility('panel-analysis', false);
        
        // Show writing assistance features
        this.setButtonText('generate-response', '✍️ Writing Assistant');
        this.setElementVisibility('generate-response', true);
        
        // Update tab labels if they exist
        this.setElementText('tab-analysis', '📝 Composition');
        this.setElementText('tab-response', '✍️ Writing Help');
    }

    /**
     * Adapts UI for sent mail (viewing previously sent emails)
     */
    adaptUIForSentMail() {
        console.debug('Adapting UI for sent mail');
        
        // Show analysis with different focus
        this.setButtonText('analyze-email', '📋 Analyze Sent Message');
        this.setElementVisibility('analyze-email', true);
        
        // Change response generation to follow-up suggestions
        this.setButtonText('generate-response', '📅 Follow-up Suggestions');
        this.setElementVisibility('generate-response', true);
        
        // Update tab labels
        this.setElementText('tab-analysis', '📋 Sent Analysis');
        this.setElementText('tab-response', '📅 Follow-up');
    }

    /**
     * Adapts UI for inbox mail (received emails)
     */
    adaptUIForInboxMail() {
        console.debug('Adapting UI for inbox mail (received)');
        
        // Standard inbox functionality
        this.setButtonText('analyze-email', '🔍 Analyze Email');
        this.setElementVisibility('analyze-email', true);
        
        this.setButtonText('generate-response', '✉️ Generate Response');
        this.setElementVisibility('generate-response', true);
        
        // Standard tab labels
        this.setElementText('tab-analysis', '🔍 Analysis');
        this.setElementText('tab-response', '✉️ Response');
    }

    /**
     * Gets a human-readable context label
     * @param {Object} context - Context information
     * @returns {string} Context label
     */
    getContextLabel(context) {
        if (context.isCompose) return '📝 COMPOSING';
        if (context.isSentMail) return '📤 SENT MAIL';
        if (context.isInbox) return '📥 INBOX';
        return '📧 EMAIL';
    }

    /**
     * Get CSS class for context display styling
     * @param {Object} context - Email context object
     * @returns {string} CSS class name
     */
    getContextClass(context) {
        if (context.isCompose) return 'context-compose';
        if (context.isSentMail) return 'context-sent';
        if (context.isInbox) return 'context-inbox';
        return 'context-inbox'; // default
    }

    /**
     * Helper method to set element visibility
     * @param {string} elementId - Element ID
     * @param {boolean} visible - Whether element should be visible
     */
    setElementVisibility(elementId, visible) {
        const element = document.getElementById(elementId);
        if (element) {
            element.style.display = visible ? '' : 'none';
        }
    }

    /**
     * Helper method to set button text
     * @param {string} elementId - Button element ID
     * @param {string} text - New button text
     */
    setButtonText(elementId, text) {
        const element = document.getElementById(elementId);
        if (element) {
            element.textContent = text;
        }
    }

    /**
     * Helper method to set element text content
     * @param {string} elementId - Element ID
     * @param {string} text - New text content
     */
    setElementText(elementId, text) {
        const element = document.getElementById(elementId);
        if (element) {
            element.textContent = text;
        }
    }

    async attemptAutoAnalysis() {
        console.debug('Checking if automatic analysis should be performed...');
        
        // Only auto-analyze if we have an email
        if (!this.currentEmail) {
            console.debug('No email available for auto-analysis');
            return;
        }

        // Skip auto-analysis if user is in initial setup mode (no API key configured)
        const needsSetup = await this.checkForInitialSetupNeeded(false);
        if (needsSetup) {
            console.debug('Initial setup needed, skipping auto-analysis');
            return;
        }

        try {
            // Get current AI provider settings
            const currentSettings = await this.settingsManager.getSettings();
            console.debug('[DEBUG] Auto-analysis settings check:', {
                fullSettings: currentSettings,
                modelService: currentSettings['model-service'],
                modelServiceType: typeof currentSettings['model-service'],
                modelServiceLength: currentSettings['model-service']?.length
            });
            
            const selectedService = currentSettings['model-service'];
            
            // Also check what the UI element shows
            const modelServiceElement = document.getElementById('model-service');
            console.debug('[DEBUG] UI element check:', {
                elementExists: !!modelServiceElement,
                elementValue: modelServiceElement?.value,
                elementType: typeof modelServiceElement?.value,
                optionsCount: modelServiceElement?.options?.length,
                selectedIndex: modelServiceElement?.selectedIndex,
                allOptions: modelServiceElement ? Array.from(modelServiceElement.options).map(opt => ({value: opt.value, text: opt.text, selected: opt.selected})) : 'N/A'
            });
            
            if (!selectedService) {
                console.warn('No AI service configured, skipping auto-analysis');
                return;
            }

            // Check for classification compatibility with selected provider
            const classification = this.classificationDetector.detectClassification(this.currentEmail.body);
            console.debug('Email classification for auto-analysis:', classification);
            
            // Skip auto-analysis for SECRET or above
            if (classification.level > 2) {
                console.debug('Classification level too high for auto-analysis');
                return;
            }

            // Test AI service health
            const config = this.getAIConfiguration();
            const isHealthy = await this.aiService.testConnection(config);
            
            if (!isHealthy) {
                console.debug('AI service not healthy, skipping auto-analysis');
                return;
            }

            console.info('Conditions met, performing automatic analysis...');
            await this.performAnalysisWithResponse();
            
        } catch (error) {
            console.error('Error during auto-analysis check:', error);
            // Don't show error to user, just skip auto-analysis
        }
    }

    async performAnalysisWithResponse() {
        const analysisStartTime = Date.now();
        let analysisEndTime, responseStartTime, responseEndTime;
        
        try {
            this.uiController.showStatus('Auto-analyzing email...');
            
            // Get AI configuration
            const config = this.getAIConfiguration();
            
            // Get classification information for telemetry
            const classificationResult = this.currentEmail.classificationResult || 
                this.classificationDetector.detectClassification(this.currentEmail.body);
            
            // Perform analysis
            this.currentAnalysis = await this.aiService.analyzeEmail(this.currentEmail, config);
            analysisEndTime = Date.now();
            
            // Display results
            this.displayAnalysis(this.currentAnalysis);
            
            // Auto-generate response as well (consolidating user actions)
            console.info('Auto-generating response after analysis...');
            responseStartTime = Date.now();
            const responseConfig = this.getResponseConfiguration();
            
            // Check email context to determine response type
            const emailContext = this.currentEmail.context || { isSentMail: false };
            
            if (emailContext.isSentMail) {
                // Generate follow-up suggestions for sent mail
                console.info('Generating follow-up suggestions for sent mail...');
                this.currentResponse = await this.aiService.generateFollowupSuggestions(
                    this.currentEmail, 
                    this.currentAnalysis, 
                    { ...config, ...responseConfig }
                );
            } else {
                // Generate response for received mail
                console.info('Generating response for received mail...');
                this.currentResponse = await this.aiService.generateResponse(
                    this.currentEmail, 
                    this.currentAnalysis, 
                    { ...config, ...responseConfig }
                );
            }
            responseEndTime = Date.now();
            
            // Display the response
            this.displayResponse(this.currentResponse);
            this.showResponseSection();
            
            // Switch to response tab for convenience
            this.switchToResponseTab();
            
            // Show refine button so user can modify the auto-generated response
            this.showRefineButton();
            
            // Log successful auto-analysis and response generation with flattened performance metrics
            this.logger.logEvent('auto_analysis_completed', {
                model_service: config.service,
                model_name: config.model,
                email_length: this.currentEmail.bodyLength,
                auto_response_generated: true,
                email_context: this.currentEmail.context ? (this.currentEmail.context.isSentMail ? 'sent' : 'inbox') : 'unknown',
                generation_type: 'standard_response',
                refinement_count: this.refinementCount,
                clipboard_used: this.hasUsedClipboard,
                // Flattened performance metrics
                analysis_duration_ms: analysisEndTime - analysisStartTime,
                response_generation_duration_ms: responseEndTime - responseStartTime,
                total_duration_ms: responseEndTime - analysisStartTime
            }, 'Information', this.getUserEmailForTelemetry());
            
            this.uiController.showStatus('Email analyzed and draft response generated automatically.');
            
        } catch (error) {
            console.error('Auto-analysis failed:', error);
            this.uiController.showStatus('Automatic analysis failed. You can still analyze manually.');
        }
    }

    async analyzeEmail() {
        if (!this.currentEmail) {
            this.uiController.showError('No email selected. Please select an email first.');
            return;
        }

        // Check for classification compatibility with selected provider
        const classification = this.classificationDetector.detectClassification(this.currentEmail.body);
        console.debug('Email classification check:', classification);
        
        // Skip classification checks if user has already overridden for this email
        if (this.classificationOverrideGranted) {
            console.debug('Classification override already granted, proceeding with analysis');
            await this.performAnalysis();
            return;
        }
        
        // Get current AI provider settings
        const currentSettings = await this.settingsManager.getSettings();
        const selectedService = currentSettings['model-service'];
        
        // Check for restricted classifications (SECRET and above) - show warning for user override
        if (classification.restricted) {
            this.pendingAction = 'analyze';
            this.showClassificationWarning(classification);
            return;
        }

        await this.performAnalysis();
    }

    showClassificationWarning(classification) {
        const warningPanel = document.getElementById('classification-warning');
        const message = document.getElementById('classification-message');
        
        let warningText = `This email is classified as ${classification.text}.`;
        
        if (classification.markings && classification.markings.length > 0) {
            warningText += ` Found ${classification.markings.length} classification marking${classification.markings.length > 1 ? 's' : ''}.`;
        }
        
        warningText += ` Proceeding may violate security policies and will be logged for compliance review.`;
        
        message.textContent = warningText;
        warningPanel.classList.remove('hidden');
        
        // Classification warning shown - logged to browser console only
        console.warn('Classification warning displayed:', classification.text);
    }

    async proceedWithWarning() {
        // Hide warning
        document.getElementById('classification-warning').classList.add('hidden');
        
        // Get current classification for console logging only
        const classification = this.classificationDetector.detectClassification(this.currentEmail.body);
        
        // Classification warning overridden - logged to browser console only
        console.warn('Classification warning overridden by user:', classification.text);
        
        // Proceed with the action the user was trying to perform
        if (this.pendingAction === 'analyze') {
            await this.performAnalysis();
        } else if (this.pendingAction === 'generateResponse') {
            await this.continueResponseGeneration();
        } else {
            // Default to analysis if no pending action
            await this.performAnalysis();
        }
        
        // Mark that user has overridden classification for this email
        this.classificationOverrideGranted = true;
        
        // Clear pending action
        this.pendingAction = null;
    }

    async continueResponseGeneration() {
        // This function contains the response generation logic without classification checks
        // since the user has already been warned and chosen to proceed
        try {
            this.uiController.showStatus('Generating response...');
            this.uiController.setButtonLoading('generate-response', true);
            
            // Get configuration
            const config = this.getAIConfiguration();
            const responseConfig = this.getResponseConfiguration();
            
            // Ensure we have analysis data - if not, run analysis first
            let analysisData = this.currentAnalysis;
            if (!analysisData) {
                console.warn('No current analysis available, running analysis first');
                this.uiController.showStatus('Analyzing email before generating response...');
                
                try {
                    // Run analysis first (this will bypass classification checks since we're in override mode)
                    await this.performAnalysis();
                    analysisData = this.currentAnalysis;
                    
                    if (!analysisData) {
                        // If analysis still failed, create minimal default
                        console.warn('Analysis failed, using default analysis');
                        analysisData = {
                            keyPoints: ['Email content needs response'],
                            sentiment: 'neutral',
                            responseStrategy: 'respond professionally and appropriately'
                        };
                    }
                } catch (analysisError) {
                    console.warn('Analysis failed, using default analysis:', analysisError);
                    analysisData = {
                        keyPoints: ['Email content needs response'],
                        sentiment: 'neutral',
                        responseStrategy: 'respond professionally and appropriately'
                    };
                }
                
                this.uiController.showStatus('Generating response...');
            }
            
            // Generate response
            this.currentResponse = await this.aiService.generateResponse(
                this.currentEmail,
                analysisData,
                { ...config, ...responseConfig }
            );

            console.info('Response generated:', this.currentResponse);
            
            // Display response
            this.displayResponse(this.currentResponse);
            this.switchToResponseTab();
            this.showRefineButton();
            
            this.uiController.showStatus('Response generated successfully.');
            
        } catch (error) {
            console.error('Response generation failed:', error);
            this.uiController.showError('Failed to generate response. Please try again.');
        } finally {
            this.uiController.setButtonLoading('generate-response', false);
        }
    }

    cancelAnalysis() {
        document.getElementById('classification-warning').classList.add('hidden');
        this.uiController.showStatus('Analysis cancelled due to classification restrictions.');
    }

    async performAnalysis() {
        const analysisStartTime = Date.now();
        
        try {
            this.uiController.showStatus('Analyzing email...');
            this.uiController.setButtonLoading('analyze-email', true);
            
            // Get AI configuration
            const config = this.getAIConfiguration();
            
            // Get classification information for telemetry
            const classificationResult = this.currentEmail.classificationResult || 
                this.classificationDetector.detectClassification(this.currentEmail.body);
            
            // Perform analysis
            this.currentAnalysis = await this.aiService.analyzeEmail(this.currentEmail, config);
            const analysisEndTime = Date.now();
            
            // Display results
            this.displayAnalysis(this.currentAnalysis);
            this.showResponseSection();
            
            // Log successful analysis with flattened performance telemetry
            this.logger.logEvent('email_analyzed', {
                model_service: config.service,
                model_name: config.model,
                email_length: this.currentEmail.bodyLength,
                recipients_count: this.currentEmail.recipients.split(',').length,
                analysis_success: true,
                refinement_count: this.refinementCount,
                clipboard_used: this.hasUsedClipboard,
                // Flattened performance metrics
                analysis_duration_ms: analysisEndTime - analysisStartTime
            }, 'Information', this.getUserEmailForTelemetry());
            
            this.uiController.showStatus('Email analysis completed successfully.');
            
        } catch (error) {
            console.error('Analysis failed:', error);
            
            // Provide more specific error messages based on error type
            let userMessage = 'Analysis failed. Please check your configuration and try again.';
            let showSettings = false;
            
            if (error.message && error.message.includes('Authentication failed')) {
                userMessage = 'Analysis failed: Invalid or missing API key. Please check your API key in the settings panel.';
                showSettings = true;
            } else if (error.message && error.message.includes('Access forbidden')) {
                userMessage = 'Analysis failed: API key permissions issue. Please verify your key has the correct permissions.';
                showSettings = true;
            } else if (error.message && error.message.includes('Service not found')) {
                userMessage = 'Analysis failed: Service endpoint not found. Please verify your endpoint URL in settings.';
                showSettings = true;
            } else if (error.message && error.message.includes('Rate limit exceeded')) {
                userMessage = 'Analysis failed: Rate limit exceeded. Please wait a moment and try again.';
            }
            
            // Show the error with additional context
            this.uiController.showError(userMessage);
            
            // If it's a configuration issue, also provide a way to access settings
            if (showSettings) {
                // Switch to settings tab to help user fix the issue
                setTimeout(() => {
                    const settingsTab = document.querySelector('button[data-tab="settings"]');
                    if (settingsTab) {
                        settingsTab.click();
                    }
                }, 2000);
            }
        } finally {
            this.uiController.setButtonLoading('analyze-email', false);
        }
    }

    async generateResponse() {
        if (!this.currentEmail) {
            this.uiController.showError('No email to respond to. Please analyze an email first.');
            return;
        }

        // Check if this is sent mail context - handle differently
        if (this.currentEmail.context && this.currentEmail.context.isSentMail) {
            await this.generateFollowupSuggestions();
            return;
        }

        // Check for classification compatibility with selected provider BEFORE generating
        const classification = this.classificationDetector.detectClassification(this.currentEmail.body);
        console.debug('Email classification check for response generation:', classification);
        
        // Skip classification checks if user has already overridden for this email
        if (this.classificationOverrideGranted) {
            console.debug('Classification override already granted, proceeding with response generation');
            await this.continueResponseGeneration();
            return;
        }
        
        // Get current AI provider settings
        const currentSettings = await this.settingsManager.getSettings();
        const selectedService = currentSettings['model-service'];
        
        // Check for restricted classifications (SECRET and above) - show warning for user override
        if (classification.restricted) {
            this.pendingAction = 'generateResponse';
            this.showClassificationWarning(classification);
            return;
        }

        try {
            this.uiController.showStatus('Generating response...');
            this.uiController.setButtonLoading('generate-response', true);
            
            // Get configuration
            const config = this.getAIConfiguration();
            const responseConfig = this.getResponseConfiguration();
            
            // Ensure we have analysis data - if not, run analysis first
            let analysisData = this.currentAnalysis;
            if (!analysisData) {
                console.warn('No current analysis available, running analysis first');
                this.uiController.showStatus('Analyzing email before generating response...');
                
                try {
                    // Run analysis first
                    await this.performAnalysis();
                    analysisData = this.currentAnalysis;
                    
                    if (!analysisData) {
                        // If analysis still failed, create minimal default
                        console.warn('Analysis failed, using default analysis');
                        analysisData = {
                            keyPoints: ['Email content needs response'],
                            sentiment: 'neutral',
                            responseStrategy: 'respond professionally and appropriately'
                        };
                    }
                } catch (analysisError) {
                    console.warn('Analysis failed, using default analysis:', analysisError);
                    analysisData = {
                        keyPoints: ['Email content needs response'],
                        sentiment: 'neutral',
                        responseStrategy: 'respond professionally and appropriately'
                    };
                }
                
                this.uiController.showStatus('Generating response...');
            }
            
            // Generate response
            this.currentResponse = await this.aiService.generateResponse(
                this.currentEmail, 
                analysisData,
                { ...config, ...responseConfig }
            );
            
            console.info('Response generated:', this.currentResponse);
            
            // Display response
            this.displayResponse(this.currentResponse);
            this.switchToResponseTab();
            this.showRefineButton();
            
            this.uiController.showStatus('Response generated successfully.');
            
        } catch (error) {
            console.error('Response generation failed:', error);
            
            // Provide more specific error messages based on error type
            let userMessage = 'Failed to generate response. Please try again.';
            let showSettings = false;
            
            if (error.message && error.message.includes('Authentication failed')) {
                userMessage = 'Response generation failed: Invalid or missing API key. Please check your API key in the settings panel.';
                showSettings = true;
            } else if (error.message && error.message.includes('Access forbidden')) {
                userMessage = 'Response generation failed: API key permissions issue. Please verify your key has the correct permissions.';
                showSettings = true;
            } else if (error.message && error.message.includes('Service not found')) {
                userMessage = 'Response generation failed: Service endpoint not found. Please verify your endpoint URL in settings.';
                showSettings = true;
            } else if (error.message && error.message.includes('Rate limit exceeded')) {
                userMessage = 'Response generation failed: Rate limit exceeded. Please wait a moment and try again.';
            }
            
            this.uiController.showError(userMessage);
            
            // If it's a configuration issue, provide guidance to fix it
            if (showSettings) {
                setTimeout(() => {
                    const settingsTab = document.querySelector('button[data-tab="settings"]');
                    if (settingsTab) {
                        settingsTab.click();
                    }
                }, 2000);
            }
        } finally {
            this.uiController.setButtonLoading('generate-response', false);
        }
    }

    async generateFollowupSuggestions() {
        if (!this.currentEmail) {
            this.uiController.showError('No email available for follow-up suggestions.');
            return;
        }

        try {
            this.uiController.showStatus('Generating follow-up suggestions...');
            this.uiController.setButtonLoading('generate-response', true);
            
            // Get configuration
            const config = this.getAIConfiguration();
            const responseConfig = this.getResponseConfiguration();
            
            // Ensure we have analysis data - if not, run analysis first
            let analysisData = this.currentAnalysis;
            if (!analysisData) {
                console.warn('No current analysis available, running analysis first');
                this.uiController.showStatus('Analyzing sent email before generating follow-up suggestions...');
                
                try {
                    await this.performAnalysis();
                    analysisData = this.currentAnalysis;
                    
                    if (!analysisData) {
                        analysisData = {
                            keyPoints: ['Sent email content analyzed'],
                            sentiment: 'neutral',
                            responseStrategy: 'generate appropriate follow-up suggestions'
                        };
                    }
                } catch (analysisError) {
                    console.warn('Analysis failed, using default analysis:', analysisError);
                    analysisData = {
                        keyPoints: ['Sent email content analyzed'],
                        sentiment: 'neutral', 
                        responseStrategy: 'generate appropriate follow-up suggestions'
                    };
                }
                
                this.uiController.showStatus('Generating follow-up suggestions...');
            }
            
            // Generate follow-up suggestions instead of response
            this.currentResponse = await this.aiService.generateFollowupSuggestions(
                this.currentEmail, 
                analysisData,
                { ...config, ...responseConfig }
            );
            
            console.info('Follow-up suggestions generated:', this.currentResponse);
            
            // Log telemetry for follow-up suggestions generation
            this.logger.logEvent('followup_suggestions_generated', {
                model_service: config.service,
                model_name: config.model,
                email_length: this.currentEmail.bodyLength,
                recipients_count: this.currentEmail.recipients.split(',').length,
                suggestions_length: this.currentResponse.suggestions ? this.currentResponse.suggestions.length : 0,
                analysis_available: !!analysisData,
                generation_success: true,
                refinement_count: this.refinementCount
            }, 'Information', this.getUserEmailForTelemetry());
            
            // Display suggestions
            this.displayResponse(this.currentResponse);
            this.switchToResponseTab();
            this.showRefineButton();
            
            this.uiController.showStatus('Follow-up suggestions generated successfully.');
            
        } catch (error) {
            console.error('Follow-up suggestion generation failed:', error);
            
            // Log telemetry for failed follow-up suggestions
            this.logger.logEvent('followup_suggestions_failed', {
                error_message: error.message,
                model_service: config ? config.service : 'unknown',
                email_length: this.currentEmail ? this.currentEmail.bodyLength : 0,
                analysis_available: !!analysisData
            }, 'Error', this.getUserEmailForTelemetry());
            
            // Provide more specific error messages based on error type
            let userMessage = 'Failed to generate follow-up suggestions. Please try again.';
            let showSettings = false;
            
            if (error.message && error.message.includes('Authentication failed')) {
                userMessage = 'Follow-up generation failed: Invalid or missing API key. Please check your API key in the settings panel.';
                showSettings = true;
            } else if (error.message && error.message.includes('Access forbidden')) {
                userMessage = 'Follow-up generation failed: API key permissions issue. Please verify your key has the correct permissions.';
                showSettings = true;
            } else if (error.message && error.message.includes('Service not found')) {
                userMessage = 'Follow-up generation failed: Service endpoint not found. Please verify your endpoint URL in settings.';
                showSettings = true;
            } else if (error.message && error.message.includes('Rate limit exceeded')) {
                userMessage = 'Follow-up generation failed: Rate limit exceeded. Please wait a moment and try again.';
            }
            
            this.uiController.showError(userMessage);
            
            // If it's a configuration issue, provide guidance to fix it
            if (showSettings) {
                setTimeout(() => {
                    const settingsTab = document.querySelector('button[data-tab="settings"]');
                    if (settingsTab) {
                        settingsTab.click();
                    }
                }, 2000);
            }
        } finally {
            this.uiController.setButtonLoading('generate-response', false);
        }
    }

    async refineResponse() {
        const customInstructions = document.getElementById('custom-instructions').value.trim();
        const currentSettings = this.settingsManager.getSettings();
        
        // Check if there are any refinement inputs (custom instructions OR changed settings)
        const hasCustomInstructions = customInstructions.length > 0;
        const hasSettingsChanges = true; // Always allow refinement based on current slider settings
        
        if (!hasCustomInstructions && !hasSettingsChanges) {
            this.uiController.showError('Please provide custom instructions or adjust response settings (length, tone, urgency) for refining the response.');
            return;
        }

        try {
            this.uiController.showStatus('Refining response...');
            this.uiController.setButtonLoading('refine-response', true);
            
            const config = this.getAIConfiguration();
            
            // Pass both custom instructions and current response settings
            // Convert settings to the same format used by generateResponse
            const responseConfig = this.getResponseConfiguration();
            this.currentResponse = await this.aiService.refineResponse(
                this.currentResponse,
                customInstructions,
                config,
                responseConfig // Use responseConfig instead of currentSettings for consistent format
            );
            
            this.displayResponse(this.currentResponse);
            this.uiController.showStatus('Response refined successfully.');
            
            // Increment refinement counter for telemetry
            this.refinementCount++;
            
            // Log response refinement event
            this.logger.logEvent('response_refined', {
                refinement_count: this.refinementCount,
                clipboard_used: this.hasUsedClipboard,
                has_custom_instructions: hasCustomInstructions,
                custom_instructions_length: customInstructions.length
            }, 'Information', this.getUserEmailForTelemetry());
            
        } catch (error) {
            console.error('Response refinement failed:', error);
            this.uiController.showError('Failed to refine response. Please try again.');
        } finally {
            this.uiController.setButtonLoading('refine-response', false);
        }
    }

    getAIConfiguration() {
        let model = this.getSelectedModel();
        // Always prioritize the modelSelect dropdown value if available
        if (this.modelSelect && this.modelSelect.value) {
            model = this.modelSelect.value;
        }
        // Fallback to saved settings if UI element is not available
        else {
            const settings = this.settingsManager.getSettings();
            if (settings['model-select']) {
                model = settings['model-select'];
            }
        }
        
        const service = this.modelServiceSelect ? this.modelServiceSelect.value : '';
        
        // Get provider-specific configuration
        let apiKey = '';
        let endpointUrl = '';
        
        if (service) {
            const providerConfig = this.settingsManager.getProviderConfig(service);
            apiKey = providerConfig['api-key'] || '';
            endpointUrl = providerConfig['endpoint-url'] || '';
            
            // Override with UI values if they exist (for immediate use before saving)
            const apiKeyElement = document.getElementById('api-key');
            const endpointUrlElement = document.getElementById('endpoint-url');
            if (apiKeyElement && apiKeyElement.value) {
                apiKey = apiKeyElement.value;
            }
            if (endpointUrlElement && endpointUrlElement.value) {
                endpointUrl = endpointUrlElement.value;
            }
        }
        
        return {
            service,
            apiKey,
            endpointUrl,
            model
        };
    }

    getResponseConfiguration() {
        const responseLengthElement = document.getElementById('response-length');
        const responseToneElement = document.getElementById('response-tone');
        const responseUrgencyElement = document.getElementById('response-urgency');
        const customInstructionsElement = document.getElementById('custom-instructions');
        
        return {
            length: responseLengthElement ? parseInt(responseLengthElement.value) : 50,
            tone: responseToneElement ? parseInt(responseToneElement.value) : 50,
            urgency: responseUrgencyElement ? parseInt(responseUrgencyElement.value) : 50,
            customInstructions: customInstructionsElement ? customInstructionsElement.value : ''
        };
    }

    getSelectedModel() {
        const service = this.modelServiceSelect ? this.modelServiceSelect.value : '';
        
        // Get default model from provider configuration instead of hardcoded map
        if (this.defaultProvidersConfig && this.defaultProvidersConfig[service]) {
            return this.defaultProvidersConfig[service].defaultModel || this.getFallbackModel();
        }
        
        // Final fallback from global config or ultimate hardcoded fallback
        return this.getFallbackModel();
    }

    getDefaultModelForProvider(provider) {
        if (!provider || !this.defaultProvidersConfig) {
            return this.getFallbackModel();
        }
        
        if (this.defaultProvidersConfig[provider] && this.defaultProvidersConfig[provider].defaultModel) {
            return this.defaultProvidersConfig[provider].defaultModel;
        }
        
        return this.getFallbackModel();
    }

    providerNeedsApiKey(provider) {
        // Ollama typically runs locally and doesn't need an API key
        if (provider === 'ollama') {
            return false;
        }
        
        // Most other providers (OpenAI, Claude, etc.) require API keys
        if (provider === 'openai' || provider === 'anthropic' || provider === 'claude') {
            return true;
        }
        
        // For onsite1/onsite2 or custom providers, assume they need API keys unless explicitly configured otherwise
        if (this.defaultProvidersConfig && this.defaultProvidersConfig[provider]) {
            // Check if the provider config indicates no API key needed
            return this.defaultProvidersConfig[provider].requiresApiKey !== false;
        }
        
        // Default to requiring API key for unknown providers
        return true;
    }

    getFallbackModel() {
        // Use fallback model from global config, or ultimate hardcoded fallback for internal deployments
        return this.defaultProvidersConfig?._config?.fallbackModel || 'llama3:latest';
    }

    async updateModelDropdown() {
        if (!this.modelServiceSelect || !this.modelSelectGroup || !this.modelSelect) return;
        
        const aiConfigPlaceholder = document.getElementById('ai-config-placeholder');
        this.modelSelectGroup.style.display = 'none';
        this.modelSelect.innerHTML = '';
        let models = [];
        let preferred = '';
        let errorMsg = '';
        if (this.modelServiceSelect.value === 'ollama') {
            this.modelSelectGroup.style.display = '';
            this.modelSelect.innerHTML = '<option value="">Loading...</option>';
            const endpointUrlElement = document.getElementById('endpoint-url');
            const baseUrl = (endpointUrlElement && endpointUrlElement.value) || 'http://localhost:11434';
            try {
                models = await AIService.fetchOllamaModels(baseUrl);
                this.modelSelect.innerHTML = models.length
                    ? models.map(m => `<option value="${m}">${m}</option>`).join('')
                    : '<option value="">No models found</option>';
                preferred = this.defaultProvidersConfig?.ollama?.defaultModel || this.getFallbackModel();
                if (preferred && models.includes(preferred)) {
                    this.modelSelect.value = preferred;
                } else if (models.length) {
                    this.modelSelect.value = models[0];
                }
                
                // Save the model selection to settings if one was set
                if (this.modelSelect.value) {
                    await this.saveSettings();
                }
            } catch (err) {
                errorMsg = `Error fetching models: ${err.message || err}`;
                // Use knownModels from provider config as fallback
                const serviceKey = this.modelServiceSelect.value;
                const providerConfig = this.defaultProvidersConfig?.[serviceKey];
                if (providerConfig?.knownModels && providerConfig.knownModels.length > 0) {
                    models = providerConfig.knownModels;
                    this.modelSelect.innerHTML = models.map(m => `<option value="${m}">${m}</option>`).join('');
                    // Still try to set the preferred model
                    preferred = this.defaultProvidersConfig?.ollama?.defaultModel || this.getFallbackModel();
                    if (preferred && models.includes(preferred)) {
                        this.modelSelect.value = preferred;
                    } else if (models.length) {
                        this.modelSelect.value = models[0];
                    }
                } else {
                    this.modelSelect.innerHTML = '<option value="">Error fetching models</option>';
                }
            }
        } else if (this.modelServiceSelect.value !== 'ollama') {
            // Handle OpenAI-compatible services (openai, onsite1, onsite2, etc.)
            this.modelSelectGroup.style.display = '';
            this.modelSelect.innerHTML = '<option value="">Loading...</option>';
            
            const serviceKey = this.modelServiceSelect.value;
            
            // Get endpoint URL: user input -> provider config -> configured fallback
            let endpoint = '';
            const endpointUrlElement = document.getElementById('endpoint-url');
            if (endpointUrlElement && endpointUrlElement.value) {
                endpoint = endpointUrlElement.value;
            } else if (this.defaultProvidersConfig && this.defaultProvidersConfig[serviceKey] && this.defaultProvidersConfig[serviceKey].baseUrl) {
                endpoint = this.defaultProvidersConfig[serviceKey].baseUrl;
            } else {
                endpoint = this.defaultProvidersConfig?._config?.fallbackBaseUrl || 'http://localhost:11434/v1';
            }
            
            if (endpoint.endsWith('/')) endpoint = endpoint.slice(0, -1);
            const apiKey = document.getElementById('api-key').value;
            try {
                models = await AIService.fetchOpenAICompatibleModels(endpoint, apiKey);
                this.modelSelect.innerHTML = models.length
                    ? models.map(m => `<option value="${m}">${m}</option>`).join('')
                    : '<option value="">No models found</option>';
            } catch (err) {
                errorMsg = err.message || `Error fetching models: ${err}`;
                
                // Show more helpful error display for authentication issues
                if (err.message && err.message.includes('Authentication failed')) {
                    // Highlight API key field or show settings reminder
                    const apiKeyField = document.getElementById('api-key');
                    if (apiKeyField) {
                        apiKeyField.style.borderColor = '#dc3545';
                        apiKeyField.style.borderWidth = '2px';
                        // Remove highlight after 5 seconds
                        setTimeout(() => {
                            apiKeyField.style.borderColor = '';
                            apiKeyField.style.borderWidth = '';
                        }, 5000);
                    }
                }
                
                // Use knownModels from provider config as fallback, or global config fallback
                const serviceKey = this.modelServiceSelect.value;
                const providerConfig = this.defaultProvidersConfig?.[serviceKey];
                models = providerConfig?.knownModels || [this.getFallbackModel()];
                this.modelSelect.innerHTML = models.map(m => `<option value="${m}">${m}</option>`).join('');
            }
            // Get preferred model from the specific service's config, not just openai
            preferred = this.defaultProvidersConfig?.[serviceKey]?.defaultModel || this.getFallbackModel();
            if (preferred && models.includes(preferred)) {
                this.modelSelect.value = preferred;
            } else if (models.length) {
                this.modelSelect.value = models[0];
            }
        } else {
            // Hide model dropdown for services that don't support model selection
            this.modelSelectGroup.style.display = 'none';
        }
        // Show error if needed
        let errorDiv = document.getElementById('model-select-error');
        if (errorMsg) {
            if (!errorDiv) {
                errorDiv = document.createElement('div');
                errorDiv.id = 'model-select-error';
                errorDiv.style.cssText = 'color: #dc3545; background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 4px; padding: 8px; margin-top: 5px; font-size: 14px; line-height: 1.4;';
                this.modelSelectGroup.appendChild(errorDiv);
            }
            
            // Create more helpful error messages with action suggestions
            let displayMessage = errorMsg;
            if (errorMsg.includes('Authentication failed')) {
                displayMessage = '🔐 ' + errorMsg + '\n\n💡 Tip: Check that your API key is entered correctly and has not expired.';
            } else if (errorMsg.includes('Access forbidden')) {
                displayMessage = '🚫 ' + errorMsg + '\n\n💡 Tip: Contact your administrator to verify API key permissions.';
            } else if (errorMsg.includes('Service not found')) {
                displayMessage = '🔗 ' + errorMsg + '\n\n💡 Tip: Verify your endpoint URL is correct and the service is running.';
            } else if (errorMsg.includes('Rate limit exceeded')) {
                displayMessage = '⏰ ' + errorMsg + '\n\n💡 Tip: Wait a few moments before trying again.';
            }
            
            errorDiv.innerHTML = displayMessage.replace(/\n/g, '<br>');
            errorDiv.style.display = 'block';
        } else if (errorDiv) {
            errorDiv.style.display = 'none';
        }
        // Hide AI config placeholder in main UI if model discovery succeeds
        if (aiConfigPlaceholder) {
            aiConfigPlaceholder.classList.add('hidden');
            aiConfigPlaceholder.innerHTML = '';
        }
    }

    displayAnalysis(analysis) {
        const container = document.getElementById('email-analysis');
        
        // Build due dates section if present
        let dueDatesHtml = '';
        if (analysis.dueDates && analysis.dueDates.length > 0) {
            const dueDateItems = analysis.dueDates.map(dueDate => {
                const urgentClass = dueDate.isUrgent ? 'urgent-due-date' : '';
                const dateDisplay = dueDate.date !== 'unspecified' ? dueDate.date : 'Date not specified';
                const timeDisplay = dueDate.time !== 'unspecified' ? ` at ${dueDate.time}` : '';
                
                return `<li class="due-date-item ${urgentClass}">
                    <strong>${this.escapeHtml(dueDate.description)}</strong><br>
                    <span class="due-date-info">Due: ${dateDisplay}${timeDisplay}</span>
                    ${dueDate.isUrgent ? '<span class="urgent-badge">URGENT</span>' : ''}
                </li>`;
            }).join('');
            
            dueDatesHtml = `
                <h3 class="due-dates-header">⏰ Due Dates & Deadlines</h3>
                <ul class="due-dates-list">
                    ${dueDateItems}
                </ul>
            `;
        }
        
        container.innerHTML = `
            <div class="analysis-content">
                ${dueDatesHtml}
                
                <h3>Key Points</h3>
                <ul>
                    ${analysis.keyPoints.map(point => `<li>${this.escapeHtml(point)}</li>`).join('')}
                </ul>

                <h3>Intent & Sentiment</h3>
                <ul>
                    <li><strong>Purpose:</strong> ${this.escapeHtml(analysis.intent || 'Not specified')}</li>
                    <li><strong>Tone:</strong> ${this.escapeHtml(analysis.sentiment)}</li>
                    <li><strong>Urgency:</strong> ${analysis.urgencyLevel}/5 - ${this.escapeHtml(analysis.urgencyReason || 'No reason provided')}</li>
                </ul>

                <h3>Recommended Actions</h3>
                <ul>
                    ${analysis.actions.map(action => `<li>${this.escapeHtml(action)}</li>`).join('')}
                </ul>
                
                ${analysis.responseStrategy ? `
                <h3>Response Strategy</h3>
                <ul>
                    <li>${this.escapeHtml(analysis.responseStrategy)}</li>
                </ul>
                ` : ''}
            </div>
        `;
    }

    displayResponse(response) {
        console.debug('Displaying response:', response);
        const container = document.getElementById('response-draft');
        
        if (!container) {
            console.error('response-draft container not found');
            return;
        }
        
        if (!response || (!response.text && !response.suggestions)) {
            console.error('Invalid response object:', response);
            container.innerHTML = '<div class="error">Error: Invalid response received</div>';
            return;
        }
        
        // Handle both regular responses (text) and follow-up suggestions (suggestions)
        const responseContent = response.text || response.suggestions;
        
        // Use separate formatting for display (less aggressive than clipboard)
        const cleanText = this.formatTextForDisplay(responseContent);
        
        container.innerHTML = `
            <div class="response-content">
                <div class="response-text" id="response-text-content">
                    ${this.escapeHtml(cleanText).replace(/\n/g, '<br>')}
                </div>
            </div>
        `;
        
        console.info('Response displayed successfully');
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    async copyResponse() {
        try {
            // Get the original response text from the currentResponse object for better formatting
            let responseText = '';
            
            if (this.currentResponse && (this.currentResponse.text || this.currentResponse.suggestions)) {
                // Handle both regular responses (text) and follow-up suggestions (suggestions)
                responseText = this.currentResponse.text || this.currentResponse.suggestions;
            } else {
                // Fallback to displayed content if currentResponse not available
                const responseElement = document.getElementById('response-text-content');
                responseText = responseElement ? responseElement.textContent : '';
            }
            
            if (!responseText) {
                this.uiController.showError('No response text to copy.');
                return;
            }
            
            // Format the text properly for clipboard with proper line breaks
            const formattedText = this.formatTextForClipboard(responseText);
            
            await navigator.clipboard.writeText(formattedText);
            this.uiController.showStatus('Response copied to clipboard.');
            
            // Track clipboard usage for telemetry
            this.hasUsedClipboard = true;
            
            // Log clipboard usage event
            this.logger.logEvent('response_copied', {
                content_type: this.currentResponse.suggestions ? 'followup_suggestions' : 'standard_response',
                email_context: this.currentEmail.context ? (this.currentEmail.context.isSentMail ? 'sent' : 'inbox') : 'unknown',
                refinement_count: this.refinementCount,
                response_length: formattedText.length
            }, 'Information', this.getUserEmailForTelemetry());
        } catch (error) {
            console.error('Failed to copy response:', error);
            this.uiController.showError('Failed to copy response to clipboard.');
        }
    }

    /**
     * Format text for display in the TaskPane (more conservative than clipboard)
     * @param {string} text - The text to format
     * @returns {string} Formatted text for display
     */
    formatTextForDisplay(text) {
        // Start with the cleaned text
        let formatted = text.trim();
        
        // Remove ALL forms of tabs and tab-like characters aggressively
        formatted = formatted.replace(/\t/g, '');  // Regular tabs
        formatted = formatted.replace(/\u0009/g, ''); // Unicode tab
        formatted = formatted.replace(/\u00A0/g, ' '); // Non-breaking space to regular space
        formatted = formatted.replace(/\u2009/g, ' '); // Thin space to regular space
        formatted = formatted.replace(/\u200B/g, ''); // Zero-width space
        formatted = formatted.replace(/\u2000-\u200F/g, ''); // Various Unicode spaces
        formatted = formatted.replace(/\u2028/g, '\n'); // Line separator to newline
        formatted = formatted.replace(/\u2029/g, '\n\n'); // Paragraph separator to double newline
        
        // Remove excessive spaces
        formatted = formatted.replace(/[ ]{2,}/g, ' '); // Multiple spaces to single space
        
        // Normalize line endings
        formatted = formatted.replace(/\r\n?/g, '\n');
        
        // Remove leading/trailing whitespace from each line, including any hidden characters
        formatted = formatted.split('\n').map(line => {
            return line.replace(/^[\s\t\u00A0\u2000-\u200F\u2028\u2029]+|[\s\t\u00A0\u2000-\u200F\u2028\u2029]+$/g, '');
        }).join('\n');
        
        // Remove empty lines at the beginning and end
        formatted = formatted.replace(/^\n+/, '').replace(/\n+$/, '');
        
        // Only add minimal paragraph breaks - don't be as aggressive as clipboard version
        // Just ensure there's a break after greeting if it doesn't exist
        if (formatted.match(/(Hi|Hello|Dear)\s+[^,]+,\s*[A-Z]/)) {
            formatted = formatted.replace(/((?:Hi|Hello|Dear)\s+[^,]+,)\s*([A-Z])/gi, '$1\n\n$2');
        }
        
        console.debug('formatTextForDisplay - Original:', JSON.stringify(text));
        console.debug('formatTextForDisplay - Formatted:', JSON.stringify(formatted));
        
        return formatted;
    }

    /**
     * Format text for clipboard with proper line breaks and paragraph spacing
     * @param {string} text - The text to format
     * @returns {string} Formatted text with proper spacing
     */
    formatTextForClipboard(text) {
        // Start with the cleaned text
        let formatted = text.trim();
        
        // Remove any existing tabs and excessive spaces
        formatted = formatted.replace(/\t+/g, ' ');
        formatted = formatted.replace(/[ ]{2,}/g, ' ');
        
        // Normalize line endings to \n
        formatted = formatted.replace(/\r\n?/g, '\n');
        
        // If the text doesn't already have proper paragraph breaks, add them
        if (!formatted.includes('\n\n')) {
            // Add breaks after common greetings
            formatted = formatted.replace(/((?:Hi|Hello|Dear)\s+[^,]+,)\s*/gi, '$1\n\n');
            
            // Add breaks before common closings
            formatted = formatted.replace(/\s*((?:Best\s+)?(?:regards?|sincerely|thanks?|cheers),?\s*\n?\s*[\w\s]+)$/gi, '\n\n$1');
            
            // Add breaks after sentences that are likely to end paragraphs
            formatted = formatted.replace(/([.!?])\s+([A-Z][a-z])/g, '$1\n\n$2');
            
            // Clean up any triple+ line breaks
            formatted = formatted.replace(/\n{3,}/g, '\n\n');
        }
        
        // Ensure signature is on its own line
        formatted = formatted.replace(/([.!?])\s*((?:Best\s+)?(?:regards?|sincerely|thanks?|cheers),?)\s*([A-Z][\w\s]+)$/gi, '$1\n\n$2\n$3');
        
        // Final cleanup
        formatted = formatted.trim();
        
        console.debug('[formatTextForClipboard] Original:', JSON.stringify(text));
        console.debug('[formatTextForClipboard] Formatted:', JSON.stringify(formatted));
        
        return formatted;
    }

    showRefineButton() {
        document.getElementById('refine-response').classList.remove('hidden');
    }

    async onModelServiceChange(event) {
        console.debug('[DEBUG] onModelServiceChange triggered:', {
            value: event.target.value,
            oldValue: event.target.dataset.oldValue || 'undefined'
        });
        
        const customEndpoint = document.getElementById('custom-endpoint');
        if (customEndpoint) {
            if (event.target.value === 'custom') {
                customEndpoint.classList.remove('hidden');
            } else {
                customEndpoint.classList.add('hidden');
            }
        }
        
        // Save current provider's settings before switching
        const oldProvider = event.target.dataset.oldValue;
        if (oldProvider && oldProvider !== 'undefined') {
            await this.saveCurrentProviderSettings(oldProvider);
        }
        
        // Load new provider's settings
        await this.loadProviderSettings(event.target.value);
        
        // Update provider labels in UI
        this.updateProviderLabels(event.target.value);
        
        // Store old value for next time
        event.target.dataset.oldValue = event.target.value;
        
        // Update model dropdown first (this will set default model and save settings)
        await this.updateModelDropdown();
        
        console.debug('[DEBUG] About to save settings after model service change');
        await this.saveSettings();
    }

    openSettings() {
        document.getElementById('settings-panel').classList.remove('hidden');
    }

    closeSettings() {
        document.getElementById('settings-panel').classList.add('hidden');
    }

    async resetSettings() {
        try {
            // Create a simple confirmation using the existing UI
            const confirmed = await this.showConfirmDialog(
                'Reset All Settings',
                'Are you sure you want to reset all settings to defaults? This will:\n\n' +
                '• Clear all API keys for all providers\n' +
                '• Reset all preferences to default values\n' +
                '• Clear all custom configurations\n' +
                '• Reset to default provider and model\n\n' +
                'This action cannot be undone.'
            );
            
            if (!confirmed) {
                return;
            }
            
            // Use SettingsManager to properly clear all settings
            const success = await this.settingsManager.clearAllSettings();
            
            if (success) {
                // Get default provider and model from S3 config
                const defaultProvider = this.defaultProvidersConfig?._config?.defaultProvider || 'ollama';
                const defaultModel = this.getDefaultModelForProvider(defaultProvider);
                
                // Set default provider and model
                if (this.modelServiceSelect) {
                    this.modelServiceSelect.value = defaultProvider;
                }
                
                // Save the default settings
                await this.settingsManager.saveSettings({
                    'model-service': defaultProvider,
                    'model-select': defaultModel
                });
                
                // Check if the default provider requires an API key
                const needsApiKey = this.providerNeedsApiKey(defaultProvider);
                
                if (needsApiKey) {
                    // Show success message with API key instruction
                    await this.showInfoDialog('Settings Reset - Action Required', 
                        `Settings have been reset to defaults.\n\nDefault provider: ${defaultProvider}\nDefault model: ${defaultModel}\n\n⚠️ IMPORTANT: This provider requires an API key.\n\nAfter the page reloads:\n1. Open Settings (⚙️)\n2. Enter your ${defaultProvider.toUpperCase()} API key\n3. Close Settings to save\n\nThe application will now reload.`);
                } else {
                    // Show success message for local providers
                    await this.showInfoDialog('Success', 
                        `Settings have been reset to defaults.\n\nDefault provider: ${defaultProvider}\nDefault model: ${defaultModel}\n\nThe application will now reload.`);
                }
                
                window.location.reload();
            } else {
                // Show error message
                await this.showInfoDialog('Error', 'Failed to reset settings. Please try again or contact support.');
            }
            
        } catch (error) {
            console.error('Error during settings reset:', error);
            await this.showInfoDialog('Error', 'An error occurred while resetting settings. Please try again.');
        }
    }

    // Simple dialog replacement for Office Add-in environment
    showConfirmDialog(title, message) {
        return new Promise((resolve) => {
            // Create a simple overlay dialog since Office Add-ins don't support native dialogs
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
                background: rgba(0,0,0,0.5); z-index: 10000; display: flex; 
                align-items: center; justify-content: center;
            `;
            
            const dialog = document.createElement('div');
            dialog.style.cssText = `
                background: white; padding: 20px; border-radius: 8px; max-width: 400px; 
                box-shadow: 0 4px 12px rgba(0,0,0,0.3); text-align: center;
            `;
            
            dialog.innerHTML = `
                <h3 style="margin-top: 0; color: #d73502;">${title}</h3>
                <p style="white-space: pre-line; margin: 16px 0;">${message}</p>
                <div style="margin-top: 20px;">
                    <button id="confirm-yes" style="margin-right: 10px; padding: 8px 16px; background: #d73502; color: white; border: none; border-radius: 4px; cursor: pointer;">Reset Settings</button>
                    <button id="confirm-no" style="padding: 8px 16px; background: #ccc; color: black; border: none; border-radius: 4px; cursor: pointer;">Cancel</button>
                </div>
            `;
            
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
            
            dialog.querySelector('#confirm-yes').onclick = () => {
                document.body.removeChild(overlay);
                resolve(true);
            };
            
            dialog.querySelector('#confirm-no').onclick = () => {
                document.body.removeChild(overlay);
                resolve(false);
            };
            
            // Close on overlay click
            overlay.onclick = (e) => {
                if (e.target === overlay) {
                    document.body.removeChild(overlay);
                    resolve(false);
                }
            };
        });
    }

    showInfoDialog(title, message) {
        return new Promise((resolve) => {
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
                background: rgba(0,0,0,0.5); z-index: 10000; display: flex; 
                align-items: center; justify-content: center;
            `;
            
            const dialog = document.createElement('div');
            dialog.style.cssText = `
                background: white; padding: 20px; border-radius: 8px; max-width: 400px; 
                box-shadow: 0 4px 12px rgba(0,0,0,0.3); text-align: center;
            `;
            
            dialog.innerHTML = `
                <h3 style="margin-top: 0; color: ${title === 'Error' ? '#d73502' : '#0078d4'};">${title}</h3>
                <p style="white-space: pre-line; margin: 16px 0;">${message}</p>
                <div style="margin-top: 20px;">
                    <button id="info-ok" style="padding: 8px 16px; background: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer;">OK</button>
                </div>
            `;
            
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
            
            dialog.querySelector('#info-ok').onclick = () => {
                document.body.removeChild(overlay);
                resolve();
            };
            
            // Close on overlay click
            overlay.onclick = (e) => {
                if (e.target === overlay) {
                    document.body.removeChild(overlay);
                    resolve();
                }
            };
        });
    }

    showHelpDialog(title, message, helpUrl) {
        return new Promise((resolve) => {
            // Create a custom dialog with "Open Help" and "Close" buttons
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
                background: rgba(0,0,0,0.5); z-index: 10000; display: flex; 
                align-items: center; justify-content: center;
            `;
            
            const dialog = document.createElement('div');
            dialog.style.cssText = `
                background: white; padding: 20px; border-radius: 8px; max-width: 400px; 
                box-shadow: 0 4px 12px rgba(0,0,0,0.3); text-align: center;
            `;
            
            dialog.innerHTML = `
                <h3 style="margin-top: 0; color: #0078d4;">${title}</h3>
                <p style="white-space: pre-line; margin: 16px 0; text-align: left;">${message}</p>
                <p style="margin: 16px 0; font-size: 14px; color: #666;">
                    <strong>Help URL:</strong><br>
                    <a href="${helpUrl}" target="_blank" style="color: #0078d4; word-break: break-all;">${helpUrl}</a>
                </p>
                <div style="margin-top: 20px; display: flex; gap: 10px; justify-content: center;">
                    <button id="help-open" style="padding: 8px 16px; background: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer;">Open Help</button>
                    <button id="help-close" style="padding: 8px 16px; background: #6c757d; color: white; border: none; border-radius: 4px; cursor: pointer;">Close</button>
                </div>
            `;
            
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
            
            dialog.querySelector('#help-open').onclick = () => {
                document.body.removeChild(overlay);
                resolve(true);
            };
            
            dialog.querySelector('#help-close').onclick = () => {
                document.body.removeChild(overlay);
                resolve(false);
            };
            
            // Close on overlay click
            overlay.onclick = (e) => {
                if (e.target === overlay) {
                    document.body.removeChild(overlay);
                    resolve(false);
                }
            };
        });
    }

    async showProviderHelp() {
        const currentProvider = this.modelServiceSelect?.value || 'ollama';
        const providerConfig = this.defaultProvidersConfig?.[currentProvider];
        
        if (providerConfig) {
            const helpText = providerConfig.helpText || 'No help available for this provider.';
            const helpUrl = providerConfig.helpUrl;
            
            if (helpUrl) {
                // Show custom dialog with "Open Help" and "Close" buttons
                const openHelp = await this.showHelpDialog(
                    `Help: ${providerConfig.label || currentProvider}`,
                    helpText,
                    helpUrl
                );
                
                if (openHelp) {
                    // Open the help URL in a new window
                    try {
                        window.open(helpUrl, '_blank', 'noopener,noreferrer');
                    } catch (error) {
                        console.error('Error opening help URL:', error);
                        await this.showInfoDialog('Error', 'Unable to open help page. Please visit the URL manually.');
                    }
                }
            } else {
                // Show info dialog for providers without URLs
                await this.showInfoDialog(
                    `Help: ${providerConfig.label || currentProvider}`,
                    helpText
                );
            }
        } else {
            await this.showInfoDialog('Help', 'No help available for the current provider.');
        }
    }

    openSource(event) {
        event.preventDefault();
        
        // Get GitHub repository URL from configuration
        const githubUrl = this.defaultProvidersConfig?._config?.githubRepository || 
                         'https://github.com/your-username/outlook-email-assistant';
        
        try {
            window.open(githubUrl, '_blank', 'noopener,noreferrer');
        } catch (error) {
            console.error('Error opening source repository:', error);
            // Fallback: copy URL to clipboard if available
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(githubUrl).then(() => {
                    this.showInfoDialog('Source Repository', 
                        `Unable to open browser. Repository URL copied to clipboard:\n\n${githubUrl}`);
                }).catch(() => {
                    this.showInfoDialog('Source Repository', 
                        `Unable to open browser. Please visit:\n\n${githubUrl}`);
                });
            } else {
                this.showInfoDialog('Source Repository', 
                    `Please visit the repository at:\n\n${githubUrl}`);
            }
        }
    }

    openIssues(event) {
        event.preventDefault();
        
        // Get issues URL from configuration
        const issuesUrl = this.defaultProvidersConfig?._config?.issuesUrl || 
                         'https://github.com/your-username/outlook-email-assistant/issues';
        
        try {
            window.open(issuesUrl, '_blank', 'noopener,noreferrer');
        } catch (error) {
            console.error('Error opening issues page:', error);
            // Fallback: copy URL to clipboard if available
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(issuesUrl).then(() => {
                    this.showInfoDialog('Issues Page', 
                        `Unable to open browser. Issues URL copied to clipboard:\n\n${issuesUrl}`);
                }).catch(() => {
                    this.showInfoDialog('Issues Page', 
                        `Unable to open browser. Please visit:\n\n${issuesUrl}`);
                });
            } else {
                this.showInfoDialog('Issues Page', 
                    `Please visit the issues page at:\n\n${issuesUrl}`);
            }
        }
    }

    openWiki(event) {
        event.preventDefault();
        
        // Get wiki URL from configuration
        const wikiUrl = this.defaultProvidersConfig?._config?.wikiUrl || 
                       'https://github.com/your-username/outlook-email-assistant/wiki';
        
        try {
            window.open(wikiUrl, '_blank', 'noopener,noreferrer');
        } catch (error) {
            console.error('Error opening wiki:', error);
            // Fallback: copy URL to clipboard if available
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(wikiUrl).then(() => {
                    this.showInfoDialog('Project Wiki', 
                        `Unable to open browser. Wiki URL copied to clipboard:\n\n${wikiUrl}`);
                }).catch(() => {
                    this.showInfoDialog('Project Wiki', 
                        `Unable to open browser. Please visit:\n\n${wikiUrl}`);
                });
            } else {
                this.showInfoDialog('Project Wiki', 
                    `Please visit the wiki at:\n\n${wikiUrl}`);
            }
        }
    }

    openTelemetry(event) {
        event.preventDefault();
        
        // Get telemetry URL from configuration
        const telemetryUrl = this.defaultProvidersConfig?._config?.telemetryUrl || 
                            'https://your-splunk-instance.com:8000/en-US/app/search/search';
        
        try {
            window.open(telemetryUrl, '_blank', 'noopener,noreferrer');
        } catch (error) {
            console.error('Error opening telemetry dashboard:', error);
            // Fallback: copy URL to clipboard if available
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(telemetryUrl).then(() => {
                    this.showInfoDialog('Telemetry Dashboard', 
                        `Unable to open browser. Dashboard URL copied to clipboard:\n\n${telemetryUrl}`);
                }).catch(() => {
                    this.showInfoDialog('Telemetry Dashboard', 
                        `Unable to open browser. Please visit:\n\n${telemetryUrl}`);
                });
            } else {
                this.showInfoDialog('Telemetry Dashboard', 
                    `Please visit the dashboard at:\n\n${telemetryUrl}`);
            }
        }
    }

    toggleHighContrast(enabled) {
        console.debug('toggleHighContrast called:', enabled);
        document.body.classList.toggle('high-contrast', enabled);
        console.debug('body classes after toggle:', document.body.classList.toString());
        this.saveSettings();
    }

    toggleScreenReaderMode(enabled) {
        this.accessibilityManager.setScreenReaderMode(enabled);
        this.saveSettings();
    }

    async loadSettingsIntoUI() {
        const settings = this.settingsManager.getSettings();

        // Load form values (excluding provider-specific fields)
        Object.keys(settings).forEach(key => {
            // Skip provider-specific fields and custom-instructions
            if (key === 'custom-instructions' || key === 'api-key' || key === 'endpoint-url' || key === 'provider-configs') return;
            const element = document.getElementById(key);
            if (element) {
                if (element.type === 'checkbox') {
                    element.checked = settings[key];
                } else {
                    element.value = settings[key] || '';
                }
            }
        });

        // Always blank custom-instructions on load
        const customInstructions = document.getElementById('custom-instructions');
        if (customInstructions) {
            customInstructions.value = '';
        }

        // If no model-service is set, default to the configured default provider
        const modelServiceSelect = document.getElementById('model-service');
        if (modelServiceSelect && (!settings['model-service'] || !modelServiceSelect.value)) {
            if (modelServiceSelect.options.length > 0) {
                // Use configured default provider, or fall back to first option
                const defaultProvider = this.defaultProvidersConfig?._config?.defaultProvider;
                let selectedOption = null;
                
                if (defaultProvider) {
                    // Try to find the default provider in the options
                    for (let option of modelServiceSelect.options) {
                        if (option.value === defaultProvider) {
                            selectedOption = option.value;
                            break;
                        }
                    }
                }
                
                // Fall back to first option if default provider not found
                const chosenOption = selectedOption || modelServiceSelect.options[0].value;
                modelServiceSelect.value = chosenOption;
                console.debug(`[TaskPane-${this.instanceId}] Setting default model-service to: ${chosenOption} (configured: ${defaultProvider})`);
                
                // Save this default to settings
                const updatedSettings = this.settingsManager.getSettings();
                updatedSettings['model-service'] = chosenOption;
                await this.settingsManager.saveSettings(updatedSettings);
                console.debug(`[TaskPane-${this.instanceId}] Saved default model-service to settings: ${chosenOption}`);
            }
        }

        // Load provider-specific settings for the current service
        const currentService = settings['model-service'] || this.defaultProvidersConfig?._config?.defaultProvider || 'openai';
        await this.loadProviderSettings(currentService);
        this.updateProviderLabels(currentService);

        // Trigger change events
        if (settings['model-service']) {
            document.getElementById('model-service').dispatchEvent(new Event('change'));
        }

        if (settings['high-contrast']) {
            console.debug('Applying high contrast setting on load:', settings['high-contrast']);
            this.toggleHighContrast(true);
        }

        if (settings['screen-reader-mode']) {
            this.toggleScreenReaderMode(true);
        }
    }

    saveSettings() {
        const formSettings = {};
        
        // Collect all form values except provider-specific ones
        const inputs = document.querySelectorAll('input, select, textarea');
        console.debug('[DEBUG] saveSettings: Found', inputs.length, 'form elements');
        
        inputs.forEach((input, index) => {
            if (input.id) {
                // Skip provider-specific fields as they're handled separately
                if (input.id === 'api-key' || input.id === 'endpoint-url') {
                    return;
                }
                
                const value = input.type === 'checkbox' ? input.checked : input.value;
                formSettings[input.id] = value;
                
                if (input.id === 'model-service') {
                    console.debug('[DEBUG] model-service element details:', {
                        index: index,
                        id: input.id,
                        type: input.type,
                        value: input.value,
                        selectedIndex: input.selectedIndex,
                        options: input.options ? Array.from(input.options).map(opt => opt.value) : 'N/A',
                        settingsValue: value
                    });
                }
            }
        });
        
        console.debug('[DEBUG] saveSettings collected:', formSettings);
        
        // Merge form settings with existing settings to preserve provider-configs
        const currentSettings = this.settingsManager.getSettings();
        const mergedSettings = { ...currentSettings, ...formSettings };
        
        this.settingsManager.saveSettings(mergedSettings);
    }

    /**
     * Save current provider's API key and endpoint settings
     * @param {string} provider - The provider key to save settings for
     */
    async saveCurrentProviderSettings(provider) {
        if (!provider || provider === 'undefined') return;
        
        const apiKeyElement = document.getElementById('api-key');
        const endpointUrlElement = document.getElementById('endpoint-url');
        
        const apiKey = apiKeyElement ? apiKeyElement.value.trim() : '';
        const endpointUrl = endpointUrlElement ? endpointUrlElement.value.trim() : '';
        
        console.debug(`[DEBUG] Saving settings for provider ${provider}:`, { 
            apiKeyLength: apiKey.length, 
            endpointUrl,
            hasApiKey: !!apiKey,
            elementValue: apiKeyElement ? `length=${apiKeyElement.value.length}` : 'no element'
        });
        
        await this.settingsManager.setProviderConfig(provider, apiKey, endpointUrl);
        console.debug(`Saved settings for provider ${provider}:`, { apiKey: apiKey ? '[HIDDEN]' : '[EMPTY]', endpointUrl });
    }

    /**
     * Load provider-specific settings into the UI
     * @param {string} provider - The provider key to load settings for
     */
    async loadProviderSettings(provider) {
        if (!provider || provider === 'undefined') return;
        
        const providerConfig = this.settingsManager.getProviderConfig(provider);
        
        const apiKeyElement = document.getElementById('api-key');
        const endpointUrlElement = document.getElementById('endpoint-url');
        
        if (apiKeyElement) {
            apiKeyElement.value = providerConfig['api-key'] || '';
        }
        
        if (endpointUrlElement) {
            // Determine which endpoint to use
            let endpointToUse = providerConfig['endpoint-url'] || '';
            
            if (this.defaultProvidersConfig && this.defaultProvidersConfig[provider]) {
                const defaultEndpoint = this.defaultProvidersConfig[provider].baseUrl || '';
                
                // For onsite providers, check if stored endpoint is the old incorrect OpenAI URL
                if (provider.startsWith('onsite') && defaultEndpoint) {
                    // If stored endpoint is the old OpenAI URL, replace it with the correct baseUrl
                    if (endpointToUse === 'https://api.openai.com/v1') {
                        endpointToUse = defaultEndpoint;
                    } else if (!endpointToUse) {
                        // If no stored endpoint, use the baseUrl from ai-providers.json
                        endpointToUse = defaultEndpoint;
                    }
                    // Otherwise keep the user's custom endpoint
                } else if (!endpointToUse && defaultEndpoint) {
                    // For other providers, use default only if no stored endpoint
                    endpointToUse = defaultEndpoint;
                }
            }
            
            endpointUrlElement.value = endpointToUse;
        }
        
        console.debug(`Loaded settings for provider ${provider}:`, { 
            apiKey: providerConfig['api-key'] ? '[HIDDEN]' : '', 
            endpointUrl: endpointUrlElement ? endpointUrlElement.value : 'no element'
        });
    }

    /**
     * Update provider labels in the UI to show which provider is currently selected
     * @param {string} provider - The current provider key
     */
    updateProviderLabels(provider) {
        const providerLabel = this.defaultProvidersConfig?.[provider]?.label || provider;
        
        const apiKeyLabel = document.getElementById('api-key-provider-label');
        const endpointUrlLabel = document.getElementById('endpoint-url-provider-label');
        
        if (apiKeyLabel) {
            apiKeyLabel.textContent = `(${providerLabel})`;
        }
        
        if (endpointUrlLabel) {
            endpointUrlLabel.textContent = `(${providerLabel})`;
        }
    }

    getUserId() {
        // Use the logger's consistent user context
        return this.logger?.getUserContext()?.userId || 'unknown';
    }

    /**
     * Get the current Outlook user's email address for telemetry context
     * @returns {string|null} Current user's email address or null
     */
    getUserEmailForTelemetry() {
        // For telemetry, we always want the actual user's email (the person using the add-in),
        // not the recipient's email, regardless of whether it's sent or inbox mail
        try {
            const userProfile = Office.context.mailbox.userProfile;
            return userProfile ? userProfile.emailAddress : null;
        } catch (error) {
            console.warn('Unable to get user profile for telemetry:', error);
            return null;
        }
    }

    getEmailIdentifiersForTelemetry() {
        if (!this.currentEmail) {
            return null;
        }

        // Create a hash of the subject for correlation without revealing content
        const subjectHash = this.currentEmail.subject ? 
            this.hashString(this.currentEmail.subject) : null;

        // Get available Office.js identifiers that don't reveal content
        const identifiers = {
            // Primary identifiers for email tracking
            conversationId: this.currentEmail.conversationId || null,
            itemId: this.currentEmail.itemId || null,
            itemClass: this.currentEmail.itemClass || null,
            
            // Content-safe metadata
            subjectHash: subjectHash,
            bodyLength: this.currentEmail.bodyLength || 0,
            hasAttachments: this.currentEmail.hasAttachments || false,
            hasInternetMessageId: this.currentEmail.hasInternetMessageId || false,
            
            // Email context
            itemType: this.currentEmail.itemType || null,
            isReply: this.currentEmail.isReply || false,
            date: this.currentEmail.date?.toISOString() || null
        };

        // Try to get additional identifiers if available from Office context
        try {
            if (Office.context.mailbox.item) {
                // Add any additional runtime identifiers
                if (Office.context.mailbox.item.itemId && !identifiers.itemId) {
                    identifiers.itemId = Office.context.mailbox.item.itemId;
                }
            }
        } catch (error) {
            console.debug('Could not access additional Office identifiers:', error);
        }

        return identifiers;
    }

    hashString(str) {
        // Simple hash function for subject correlation without revealing content
        let hash = 0;
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32-bit integer
        }
        return hash.toString(36); // Return as base-36 string
    }
}

// Initialize the application when Office.js is ready
Office.onReady(() => {
    const app = new TaskpaneApp();
    app.initialize().catch(error => {
        console.error('Failed to initialize application:', error);
    });
});
