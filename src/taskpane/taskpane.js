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
    async fetchDefaultModelsConfig() {
        // Fetch ai-models.json from public config directory
        try {
            const response = await fetch('/config/ai-models.json');
            if (!response.ok) throw new Error('Failed to fetch ai-models.json');
            return await response.json();
        } catch (e) {
            console.warn('[Default Models] Could not load ai-models.json:', e);
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
        this.sessionStartTime = Date.now();
        
        // Telemetry tracking properties
        this.refinementCount = 0;
        this.hasUsedClipboard = false;

        // Model selection UI elements
        this.modelServiceSelect = null;
        this.modelSelectGroup = null;
        this.modelSelect = null;
    }

    async initialize() {
        try {
            // Initialize Office.js
            await this.initializeOffice();
            
            // Load user settings
            await this.settingsManager.loadSettings();
            
            // Load provider config before UI setup
            this.defaultProvidersConfig = await this.fetchDefaultProvidersConfig();
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
        this.baseUrlInput = document.getElementById('base-url');
        
        // Populate model service dropdown from defaultProvidersConfig BEFORE loading settings
        if (this.modelServiceSelect && this.defaultProvidersConfig) {
            this.modelServiceSelect.innerHTML = Object.entries(this.defaultProvidersConfig)
                .filter(([key, val]) => key !== 'custom')
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
        
        // Model service change
        document.getElementById('model-service').addEventListener('change', (e) => this.onModelServiceChange(e));
        
        // Settings checkboxes
        document.getElementById('high-contrast').addEventListener('change', (e) => this.toggleHighContrast(e.target.checked));
        document.getElementById('screen-reader-mode').addEventListener('change', (e) => this.toggleScreenReaderMode(e.target.checked));
        
        // Auto-save settings
        ['api-key', 'endpoint-url', 'custom-instructions'].forEach(id => {
            const element = document.getElementById(id);
            if (element) {
                element.addEventListener('blur', () => this.saveSettings());
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
            this.displayEmailSummary(this.currentEmail);
        } catch (error) {
            console.error('Failed to load current email:', error);
            this.uiController.showError('Failed to load email. Please select an email and try again.');
        }
    }

    displayEmailSummary(email) {
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
                classificationColor = classificationResult.color;
                
                // Show detailed classification info if markings were found
                if (classificationResult.markings && classificationResult.markings.length > 0) {
                    classificationText += ` (${classificationResult.markings.length} marking${classificationResult.markings.length > 1 ? 's' : ''} found)`;
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
        
        // Get current AI provider settings
        const currentSettings = await this.settingsManager.getSettings();
        const selectedService = currentSettings['model-service'];
        
        // Check if provider supports this classification level
        const isCompatible = await this.checkProviderClassificationCompatibility(selectedService, classification);
        
        if (!isCompatible) {
            this.showProviderClassificationWarning(selectedService, classification);
            return;
        }
        
        // Legacy check for SECRET or above
        if (classification.level > 2) { // SECRET or above
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
            
            if (providerInfo.maxClassificationLevel !== undefined) {
                const compatible = classification.level <= providerInfo.maxClassificationLevel;
                console.debug(`Classification compatibility: ${classification.text} (level ${classification.level}) vs ${serviceProvider} (max level ${providerInfo.maxClassificationLevel}) = ${compatible}`);
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
        
        const message = `The selected AI provider "${provider}" does not support ${classification.text} classified content.\n\n${providerNote}\n\nPlease select a different provider or use unclassified content.`;
        
        this.uiController.showError(message);
        
        // Log the incompatibility
        this.logger.logEvent('classification_incompatible', {
            provider: provider,
            classification: classification.text,
            classificationLevel: classification.level,
            maxSupportedLevel: providerInfo?.maxClassificationLevel
        });
    }

    showClassificationWarning(classification) {
        const warningPanel = document.getElementById('classification-warning');
        const message = document.getElementById('classification-message');
        
        message.textContent = `This email is classified as ${classification.text}. Proceeding may violate security policies.`;
        warningPanel.classList.remove('hidden');
        
        // Log the warning
        this.logger.logEvent('classification_warning_shown', {
            classification: classification.text,
            level: classification.level,
            subject: this.currentEmail.subject
        });
    }

    async proceedWithWarning() {
        // Hide warning
        document.getElementById('classification-warning').classList.add('hidden');
        
        // Log override
        this.logger.logEvent('classification_warning_overridden', {
            subject: this.currentEmail.subject,
            user_id: this.getUserId(),
            timestamp: new Date().toISOString()
        });
        
        // Proceed with analysis
        await this.performAnalysis();
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

        try {
            this.uiController.showStatus('Generating response...');
            this.uiController.setButtonLoading('generate-response', true);
            
            // Get configuration
            const config = this.getAIConfiguration();
            const responseConfig = this.getResponseConfiguration();
            
            // Ensure we have analysis data, or create default
            let analysisData = this.currentAnalysis;
            if (!analysisData) {
                console.warn('No current analysis available, creating default analysis');
                analysisData = {
                    keyPoints: ['Email content needs response'],
                    sentiment: 'neutral',
                    responseStrategy: 'respond professionally and appropriately'
                };
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
            this.currentResponse = await this.aiService.refineResponse(
                this.currentResponse,
                customInstructions,
                config,
                currentSettings // Pass current settings for length, tone, urgency
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
        
        const apiKeyElement = document.getElementById('api-key');
        const endpointUrlElement = document.getElementById('endpoint-url');
        return {
            service: this.modelServiceSelect ? this.modelServiceSelect.value : '',
            apiKey: apiKeyElement ? apiKeyElement.value : '',
            endpointUrl: endpointUrlElement ? endpointUrlElement.value : '',
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
        const modelMap = {
            'openai': 'gpt-4',
            'ollama': '',
            'custom': 'custom'
        };
        return modelMap[service] || 'gpt-4';
    }

    async updateModelDropdown() {
        if (!this.modelServiceSelect || !this.modelSelectGroup || !this.modelSelect) return;
        // Load default models config (cache for session)
        if (!this.defaultModelsConfig) {
            this.defaultModelsConfig = await this.fetchDefaultModelsConfig();
        }
        const aiConfigPlaceholder = document.getElementById('ai-config-placeholder');
        this.modelSelectGroup.style.display = 'none';
        this.modelSelect.innerHTML = '';
        let models = [];
        let preferred = '';
        let errorMsg = '';
        if (this.modelServiceSelect.value === 'ollama') {
            this.modelSelectGroup.style.display = '';
            this.modelSelect.innerHTML = '<option value="">Loading...</option>';
            const baseUrl = (this.baseUrlInput && this.baseUrlInput.value) || 'http://localhost:11434';
            try {
                models = await AIService.fetchOllamaModels(baseUrl);
                this.modelSelect.innerHTML = models.length
                    ? models.map(m => `<option value="${m}">${m}</option>`).join('')
                    : '<option value="">No models found</option>';
                preferred = this.defaultModelsConfig && this.defaultModelsConfig.ollama;
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
                this.modelSelect.innerHTML = '<option value="">Error fetching models</option>';
            }
        } else if (this.modelServiceSelect.value === 'openai') {
            this.modelSelectGroup.style.display = '';
            this.modelSelect.innerHTML = '<option value="">Loading...</option>';
            let endpoint = this.baseUrlInput && this.baseUrlInput.value ? this.baseUrlInput.value : 'https://api.openai.com/v1';
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
                models = (data.data || []).map(m => m.id).filter(id => id.startsWith('gpt-') || id.startsWith('ft-') || id.startsWith('davinci') || id.startsWith('babbage') || id.startsWith('curie'));
                this.modelSelect.innerHTML = models.length
                    ? models.map(m => `<option value="${m}">${m}</option>`).join('')
                    : '<option value="">No models found</option>';
            } catch (err) {
                errorMsg = `Error fetching models: ${err.message || err}`;
                models = ['gpt-4', 'gpt-4o', 'gpt-3.5-turbo', 'gpt-3.5-turbo-16k'];
                this.modelSelect.innerHTML = models.map(m => `<option value="${m}">${m}</option>`).join('');
            }
            preferred = this.defaultModelsConfig && this.defaultModelsConfig.openai;
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

    toggleHighContrast(enabled) {
        document.body.classList.toggle('high-contrast', enabled);
        this.saveSettings();
    }

    toggleScreenReaderMode(enabled) {
        this.accessibilityManager.setScreenReaderMode(enabled);
        this.saveSettings();
    }

    async loadSettingsIntoUI() {
        const settings = this.settingsManager.getSettings();

        // Load form values
        Object.keys(settings).forEach(key => {
            // Never load custom-instructions from settings
            if (key === 'custom-instructions') return;
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

        // Ensure base-url input defaults to http://localhost:11434 if not set
        const baseUrlInput = document.getElementById('base-url');
        if (baseUrlInput && (!settings['base-url'] || !baseUrlInput.value)) {
            baseUrlInput.value = 'http://localhost:11434';
        }

        // If no model-service is set, default to the first available option and save it
        const modelServiceSelect = document.getElementById('model-service');
        if (modelServiceSelect && (!settings['model-service'] || !modelServiceSelect.value)) {
            if (modelServiceSelect.options.length > 0) {
                const firstOption = modelServiceSelect.options[0].value;
                modelServiceSelect.value = firstOption;
                console.debug(`[TaskPane-${this.instanceId}] Setting default model-service to: ${firstOption}`);
                
                // Save this default to settings
                const updatedSettings = this.settingsManager.getSettings();
                updatedSettings['model-service'] = firstOption;
                await this.settingsManager.saveSettings(updatedSettings);
                console.debug(`[TaskPane-${this.instanceId}] Saved default model-service to settings: ${firstOption}`);
            }
        }

        // Trigger change events
        if (settings['model-service']) {
            document.getElementById('model-service').dispatchEvent(new Event('change'));
        }

        if (settings['high-contrast']) {
            this.toggleHighContrast(true);
        }

        if (settings['screen-reader-mode']) {
            this.toggleScreenReaderMode(true);
        }
    }

    saveSettings() {
        const settings = {};
        
        // Collect all form values
        const inputs = document.querySelectorAll('input, select, textarea');
        console.debug('[DEBUG] saveSettings: Found', inputs.length, 'form elements');
        
        inputs.forEach((input, index) => {
            if (input.id) {
                const value = input.type === 'checkbox' ? input.checked : input.value;
                settings[input.id] = value;
                
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
        
        console.debug('[DEBUG] saveSettings collected:', settings);
        this.settingsManager.saveSettings(settings);
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
}

// Initialize the application when Office.js is ready
Office.onReady(() => {
    const app = new TaskpaneApp();
    app.initialize().catch(error => {
        console.error('Failed to initialize application:', error);
    });

    // Add logic for Reset Settings button
    const resetBtn = document.getElementById('reset-settings');
    if (resetBtn) {
        resetBtn.addEventListener('click', () => {
            // Remove settings from localStorage
            localStorage.removeItem('settings');
            // Optionally clear Office.js roaming settings
            if (typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
                Office.context.roamingSettings.remove('settings');
                Office.context.roamingSettings.saveAsync();
            }
            // Reload the app to reset UI
            location.reload();
        });
    }
});
