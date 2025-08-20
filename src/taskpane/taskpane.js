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
        }, 'Information', this.getRecipientEmailForTelemetry());
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
            
            // Setup accessibility
            this.accessibilityManager.initialize();
            
            // Initialize Splunk telemetry if enabled
            await this.initializeTelemetry();
            
            // Load current email
            await this.loadCurrentEmail();
            
            // Try automatic analysis if conditions are met
            await this.attemptAutoAnalysis();
            
            // Hide loading, show main content
            this.uiController.hideLoading();
            this.uiController.showMainContent();
            
            // Log session start
            this.logger.logEvent('session_start', {
                timestamp: new Date().toISOString(),
                version: '1.0.0',
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
            
            // Start Splunk auto-flush if enabled
            if (this.logger.telemetryConfig?.telemetry?.enabled && 
                this.logger.telemetryConfig.telemetry.provider === 'splunk_hec') {
                this.logger.startSplunkAutoFlush();
                console.info('Splunk telemetry enabled and auto-flush started');
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
            this.modelServiceSelect.addEventListener('change', () => {
                // Print resultant baseUrl to console for developer/support
                const providerKey = this.modelServiceSelect.value;
                if (this.defaultProvidersConfig && this.defaultProvidersConfig[providerKey]) {
                    const baseUrl = this.defaultProvidersConfig[providerKey].baseUrl || '';
                    console.debug(`Model provider key: ${providerKey}, base URL: ${baseUrl}`);
                }
                this.updateModelDropdown();
            });
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
        
        // Help and GitHub links
        document.getElementById('api-key-help-btn').addEventListener('click', () => this.showProviderHelp());
        document.getElementById('github-link').addEventListener('click', (e) => this.openGitHubRepository(e));
        
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
                element.addEventListener('blur', () => this.saveCurrentProviderSettings(this.modelServiceSelect?.value));
            }
        });
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
            await this.displayEmailSummary(this.currentEmail);
        } catch (error) {
            console.error('Failed to load current email:', error);
            this.uiController.showError('Failed to load email. Please select an email and try again.');
        }
    }

    async displayEmailSummary(email) {
        console.debug('Displaying email summary:', email);
        
        // Only update the fields we're actually showing
        const subjectElement = document.getElementById('email-subject');
        if (subjectElement) {
            subjectElement.textContent = email.subject || 'No Subject';
        }
        
        // Enhanced classification display logic using ClassificationDetector
        let classificationResult = null;
        let classificationText = 'UNCLASSIFIED';
        let classificationColor = 'green';
        
        if (email.body) {
            classificationResult = this.classificationDetector.detectClassification(email.body);
            console.debug('Classification result:', classificationResult);
            
            if (classificationResult.detected) {
                classificationText = classificationResult.text;
                
                // Only show warning color if classification is not supported by provider
                const isSupportedByProvider = await this.isClassificationSupportedByProvider(classificationResult.text);
                classificationColor = isSupportedByProvider ? 'green' : classificationResult.color;
                
                // Symmetrical classification messaging
                if (isSupportedByProvider) {
                    classificationText += ' - Safe for AI processing';
                } else {
                    classificationText += ' - Potentially unsafe for AI processing';
                }
            } else {
                classificationText = 'UNCLASSIFIED - Safe for AI processing';
            }
        }
        
        const classificationElement = document.getElementById("email-classification");
        if (classificationElement) {
            classificationElement.textContent = classificationText;
            classificationElement.className = `classification classification-${classificationColor}`;
            console.debug('Set classification display:', classificationText, classificationColor);
        }
        
        // Store classification result for later use
        if (classificationResult) {
            email.classificationResult = classificationResult;
        }
        
        // Commented out fields for debugging purposes
        // document.getElementById('email-from').textContent = email.from || 'Unknown';
        // document.getElementById('email-recipients').textContent = email.recipients || 'Unknown';
        // document.getElementById('email-length').textContent = `${email.bodyLength || 0} characters`;
    }

    async attemptAutoAnalysis() {
        console.debug('Checking if automatic analysis should be performed...');
        
        // Only auto-analyze if we have an email
        if (!this.currentEmail) {
            console.debug('No email available for auto-analysis');
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
            
            // Check if provider supports this classification level
            const isCompatible = await this.checkProviderClassificationCompatibility(selectedService, classification);
            
            if (!isCompatible) {
                console.debug('Provider incompatible with classification, skipping auto-analysis');
                return;
            }
            
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
        try {
            this.uiController.showStatus('Auto-analyzing email...');
            
            // Get AI configuration
            const config = this.getAIConfiguration();
            
            // Get classification information for telemetry
            const classificationResult = this.currentEmail.classificationResult || 
                this.classificationDetector.detectClassification(this.currentEmail.body);
            
            // Perform analysis
            this.currentAnalysis = await this.aiService.analyzeEmail(this.currentEmail, config);
            
            // Display results
            this.displayAnalysis(this.currentAnalysis);
            
            // Auto-generate response as well (consolidating user actions)
            console.info('Auto-generating response after analysis...');
            const responseConfig = this.getResponseConfiguration();
            
            // Generate response using analysis data
            this.currentResponse = await this.aiService.generateResponse(
                this.currentEmail, 
                this.currentAnalysis, 
                { ...config, ...responseConfig }
            );
            
            // Display the response
            this.displayResponse(this.currentResponse);
            this.showResponseSection();
            
            // Switch to response tab for convenience
            this.switchToResponseTab();
            
            // Show refine button so user can modify the auto-generated response
            this.showRefineButton();
            
            // Log successful auto-analysis and response generation
            this.logger.logEvent('auto_analysis_completed', {
                model_service: config.service,
                model_name: config.model,
                email_length: this.currentEmail.bodyLength,
                classification: classificationResult.text,
                classification_level: classificationResult.level,
                auto_response_generated: true,
                refinement_count: this.refinementCount,
                clipboard_used: this.hasUsedClipboard,
                performance_metrics: {
                    start_time: Date.now()
                }
            }, 'Information', this.getRecipientEmailForTelemetry());
            
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
        
        // Check if provider supports this classification level
        const isCompatible = await this.checkProviderClassificationCompatibility(selectedService, classification);
        
        if (!isCompatible) {
            this.pendingAction = 'analyze';
            this.showProviderClassificationWarning(selectedService, classification);
            return;
        }
        
        // Check for restricted classifications (SECRET and above) - show warning for user override
        if (classification.restricted) {
            this.pendingAction = 'analyze';
            this.showClassificationWarning(classification);
            return;
        }

        await this.performAnalysis();
    }

    async checkProviderClassificationCompatibility(serviceProvider, classification) {
        console.debug('Checking provider compatibility:', serviceProvider, classification);
        
        try {
            const providersConfig = await this.fetchDefaultProvidersConfig();
            const providerInfo = providersConfig[serviceProvider];
            
            if (!providerInfo) {
                console.warn('No provider configuration found for:', serviceProvider);
                return true; // Allow if no config found
            }
            
            if (providerInfo.supportedClassifications) {
                const compatible = providerInfo.supportedClassifications.includes(classification.text);
                console.debug(`Classification compatibility: ${classification.text} vs ${serviceProvider} supported: [${providerInfo.supportedClassifications.join(', ')}] = ${compatible}`);
                return compatible;
            }
            
            return true; // Allow if no classification restrictions
        } catch (error) {
            console.error('Error checking provider compatibility:', error);
            return true; // Allow on error
        }
    }

    showProviderClassificationWarning(provider, classification) {
        const providersConfig = this.defaultProvidersConfig;
        const providerInfo = providersConfig[provider];
        const providerNote = providerInfo?.classificationNote || 'Classification restrictions apply';
        
        // Show the warning panel instead of just an error message
        const warningPanel = document.getElementById('classification-warning');
        const message = document.getElementById('classification-message');
        
        let warningText = `The selected AI provider "${provider}" does not support ${classification.text} classified content.\n\n${providerNote}\n\nProceeding may violate security policies and will be logged for compliance review.`;
        
        message.textContent = warningText;
        warningPanel.classList.remove('hidden');
        
        // Log the incompatibility
        this.logger.logEvent('classification_incompatible', {
            ...this.getEmailIdentifiersForTelemetry(),
            provider: provider,
            classification: classification.text,
            provider_supported_classifications: providerInfo?.supportedClassifications
        });
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
        
        // Enhanced telemetry logging
        this.logger.logEvent('classification_warning_shown', {
            ...this.getEmailIdentifiersForTelemetry(),
            classification: classification.text,
            restricted: classification.restricted,
            markings_found: classification.markings?.length || 0,
            details: classification.details,
            timestamp: new Date().toISOString()
        });
    }

    async proceedWithWarning() {
        // Hide warning
        document.getElementById('classification-warning').classList.add('hidden');
        
        // Get current classification and provider info for detailed logging
        const classification = this.classificationDetector.detectClassification(this.currentEmail.body);
        const currentSettings = await this.settingsManager.getSettings();
        const selectedService = currentSettings['model-service'];
        const providerConfig = this.defaultProvidersConfig?.[selectedService];
        
        // Enhanced telemetry logging for security compliance
        this.logger.logEvent('classification_warning_overridden', {
            ...this.getEmailIdentifiersForTelemetry(),
            classification_detected: classification.text,
            classification_restricted: classification.restricted,
            classification_markings_count: classification.markings?.length || 0,
            provider_used: selectedService,
            provider_supported_classifications: providerConfig?.supportedClassifications,
            timestamp: new Date().toISOString(),
            warning_type: 'user_override'
        });
        
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

    async isClassificationSupportedByProvider(classificationText) {
        try {
            const currentSettings = await this.settingsManager.getSettings();
            const selectedService = currentSettings['model-service'];
            const providersConfig = await this.fetchDefaultProvidersConfig();
            const providerInfo = providersConfig[selectedService];
            
            if (!providerInfo || !providerInfo.supportedClassifications) {
                return true; // Assume supported if no config (fail open for compatibility)
            }
            
            return providerInfo.supportedClassifications.includes(classificationText);
        } catch (error) {
            console.error('Error checking classification support:', error);
            return true; // Assume supported on error (fail open for compatibility)
        }
    }

    cancelAnalysis() {
        document.getElementById('classification-warning').classList.add('hidden');
        this.uiController.showStatus('Analysis cancelled due to classification restrictions.');
    }

    async performAnalysis() {
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
            
            // Display results
            this.displayAnalysis(this.currentAnalysis);
            this.showResponseSection();
            
            // Log successful analysis with enhanced telemetry
            this.logger.logEvent('email_analyzed', {
                model_service: config.service,
                model_name: config.model,
                email_length: this.currentEmail.bodyLength,
                recipients_count: this.currentEmail.recipients.split(',').length,
                classification: classificationResult.text,
                classification_level: classificationResult.level,
                classification_detected: classificationResult.detected,
                has_markings: classificationResult.markings ? classificationResult.markings.length > 0 : false,
                analysis_success: true,
                refinement_count: this.refinementCount,
                clipboard_used: this.hasUsedClipboard,
                performance_metrics: {
                    start_time: Date.now()
                }
            }, 'Information', this.getRecipientEmailForTelemetry());
            
            this.uiController.showStatus('Email analysis completed successfully.');
            
        } catch (error) {
            console.error('Analysis failed:', error);
            this.uiController.showError('Analysis failed. Please check your API configuration and try again.');
        } finally {
            this.uiController.setButtonLoading('analyze-email', false);
        }
    }

    async generateResponse() {
        if (!this.currentEmail) {
            this.uiController.showError('No email to respond to. Please analyze an email first.');
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
        
        // Check if provider supports this classification level
        const isCompatible = await this.checkProviderClassificationCompatibility(selectedService, classification);
        
        if (!isCompatible) {
            this.pendingAction = 'generateResponse';
            this.showProviderClassificationWarning(selectedService, classification);
            return;
        }
        
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
            this.uiController.showError('Failed to generate response. Please try again.');
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
            }, 'Information', this.getRecipientEmailForTelemetry());
            
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
                const response = await fetch(`${endpoint}/models`, {
                    headers: {
                        'Authorization': `Bearer ${apiKey}`
                    }
                });
                if (!response.ok) throw new Error(`HTTP ${response.status}`);
                const data = await response.json();
                models = (data.data || []).map(m => m.id);
                this.modelSelect.innerHTML = models.length
                    ? models.map(m => `<option value="${m}">${m}</option>`).join('')
                    : '<option value="">No models found</option>';
            } catch (err) {
                errorMsg = `Error fetching models: ${err.message || err}`;
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
                errorDiv.style.color = 'red';
                this.modelSelectGroup.appendChild(errorDiv);
            }
            errorDiv.textContent = errorMsg;
        } else if (errorDiv) {
            errorDiv.remove();
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
        
        if (!response || !response.text) {
            console.error('Invalid response object:', response);
            container.innerHTML = '<div class="error">Error: Invalid response received</div>';
            return;
        }
        
        // Use separate formatting for display (less aggressive than clipboard)
        const cleanText = this.formatTextForDisplay(response.text);
        
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
            
            if (this.currentResponse && this.currentResponse.text) {
                responseText = this.currentResponse.text;
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
                refinement_count: this.refinementCount,
                response_length: formattedText.length
            }, 'Information', this.getRecipientEmailForTelemetry());
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

    openGitHubRepository(event) {
        event.preventDefault();
        
        // Get GitHub repository URL from configuration
        const githubUrl = this.defaultProvidersConfig?._config?.githubRepository || 
                         'https://github.com/your-username/outlook-email-assistant';
        
        try {
            window.open(githubUrl, '_blank', 'noopener,noreferrer');
        } catch (error) {
            console.error('Error opening GitHub repository:', error);
            // Fallback: copy URL to clipboard if available
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(githubUrl).then(() => {
                    this.showInfoDialog('GitHub Repository', 
                        `Unable to open browser. Repository URL copied to clipboard:\n\n${githubUrl}`);
                }).catch(() => {
                    this.showInfoDialog('GitHub Repository', 
                        `Unable to open browser. Please visit:\n\n${githubUrl}`);
                });
            } else {
                this.showInfoDialog('GitHub Repository', 
                    `Please visit the repository at:\n\n${githubUrl}`);
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
        // In a real implementation, this would get the actual user ID
        return Office.context.mailbox.userProfile.emailAddress || 'unknown';
    }

    /**
     * Extract the first recipient email address for telemetry context
     * @returns {string|null} First recipient email or null
     */
    getRecipientEmailForTelemetry() {
        if (!this.currentEmail?.recipients) {
            return null;
        }
        
        // Parse recipient string to extract email addresses
        // Recipients format: "To: Name <email@domain.com>, Name2 <email2@domain.com>"
        const recipientMatch = this.currentEmail.recipients.match(/<([^>]+)>/);
        return recipientMatch ? recipientMatch[1] : null;
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
            normalizedSubject: this.currentEmail.normalizedSubject || null,
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
