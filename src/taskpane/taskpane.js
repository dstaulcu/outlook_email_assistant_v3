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
        // Fetch default-providers.json from public directory
        try {
            const response = await fetch('/default-providers.json');
            if (!response.ok) throw new Error('Failed to fetch default-providers.json');
            return await response.json();
        } catch (e) {
            console.warn('[Default Providers] Could not load default-providers.json:', e);
            return {};
        }
    }
    async fetchDefaultModelsConfig() {
        // Fetch default-models.json from public directory
        try {
            const response = await fetch('/default-models.json');
            if (!response.ok) throw new Error('Failed to fetch default-models.json');
            return await response.json();
        } catch (e) {
            console.warn('[Default Models] Could not load default-models.json:', e);
            return {};
        }
    }

    switchToResponseTab() {
        // Switch to the response tab in the UI
        const responseTabButton = document.querySelector('.tab-button[aria-controls="panel-response"]');
        if (responseTabButton) {
            console.log('[TaskpaneApp] Switching to response tab');
            responseTabButton.click();
        } else {
            console.error('[TaskpaneApp] Response tab button not found');
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

        // Model selection UI elements
        this.modelServiceSelect = null;
        this.modelSelectGroup = null;
        this.modelSelect = null;
    }

    async initialize() {
        try {
        // Show/hide Analyze Email button based on whether a reply/compose is in progress
        const analyzeBtn = document.getElementById('analyze-email');
        if (analyzeBtn) {
            let isCompose = false;
            try {
                // Compose mode if setAsync is available (reply/forward/compose window)
                isCompose = (
                    typeof Office !== 'undefined' &&
                    Office.context &&
                    Office.context.mailbox &&
                    Office.context.mailbox.item &&
                    Office.context.mailbox.item.body &&
                    typeof Office.context.mailbox.item.body.setAsync === 'function'
                );
            } catch (e) {
                isCompose = false;
            }
            analyzeBtn.style.display = isCompose ? '' : 'none';
        }
            // Initialize Office.js
            await this.initializeOffice();
            
            // Load user settings
            await this.settingsManager.loadSettings();
            
            // Load provider config before UI setup
            this.defaultProvidersConfig = await this.fetchDefaultProvidersConfig();
            // Setup UI
            this.setupUI();
            
            // Setup accessibility
            this.accessibilityManager.initialize();
            
            // Load current email
            await this.loadCurrentEmail();
            
            // Hide loading, show main content
            this.uiController.hideLoading();
            this.uiController.showMainContent();
            
            // Log session start
            this.logger.logEvent('session_start', {
                timestamp: new Date().toISOString(),
                version: '1.0.0'
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

    setupUI() {
        // Bind event listeners
        this.bindEventListeners();
        
        // Initialize sliders
        this.initializeSliders();
        
        // Setup tabs
        this.initializeTabs();
        
        // Load settings into UI
        this.loadSettingsIntoUI();
        // Model selection UI elements
        this.modelServiceSelect = document.getElementById('model-service');
        this.modelSelectGroup = document.getElementById('model-select-group');
        this.modelSelect = document.getElementById('model-select');
        this.baseUrlInput = document.getElementById('base-url');
        // Populate model service dropdown from defaultProvidersConfig
        if (this.modelServiceSelect && this.defaultProvidersConfig) {
            this.modelServiceSelect.innerHTML = Object.entries(this.defaultProvidersConfig)
                .filter(([key, val]) => key !== 'custom')
                .map(([key, val]) => `<option value="${key}">${val.label}</option>`)
                .join('');
        }
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
                    console.log(`[Provider Base URL] ${providerKey}: ${baseUrl}`);
                }
                this.updateModelDropdown();
            });
            // Set initial baseUrl to console
            if (this.modelServiceSelect.value && this.defaultProvidersConfig && this.defaultProvidersConfig[this.modelServiceSelect.value]) {
                const baseUrl = this.defaultProvidersConfig[this.modelServiceSelect.value].baseUrl || '';
                console.log(`[Provider Base URL] ${this.modelServiceSelect.value}: ${baseUrl}`);
            }
            this.updateModelDropdown();
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
        document.getElementById('insert-response').addEventListener('click', () => this.insertResponse());
        
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
        // Only update the fields we're actually showing
        const subjectElement = document.getElementById('email-subject');
        if (subjectElement) {
            subjectElement.textContent = email.subject || 'No Subject';
        }
        
        // Classification display logic
        let classification = email.classification;
        let classificationText;
        if (!classification || classification.toLowerCase() === "unclassified") {
            classificationText = "This email appears to be safe for AI processing.";
        } else {
            classificationText = classification;
        }
        const classificationElement = document.getElementById("email-classification");
        if (classificationElement) {
            classificationElement.textContent = classificationText;
        }
        
        // Commented out fields for debugging purposes
        // document.getElementById('email-from').textContent = email.from || 'Unknown';
        // document.getElementById('email-recipients').textContent = email.recipients || 'Unknown';
        // document.getElementById('email-length').textContent = `${email.bodyLength || 0} characters`;
    }

    async analyzeEmail() {
        if (!this.currentEmail) {
            this.uiController.showError('No email selected. Please select an email first.');
            return;
        }

        // Check for classification
        const classification = this.classificationDetector.detectClassification(this.currentEmail.body);
        if (classification.level > 2) { // SECRET or above
            this.showClassificationWarning(classification);
            return;
        }

        await this.performAnalysis();
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
            
            // Perform analysis
            this.currentAnalysis = await this.aiService.analyzeEmail(this.currentEmail, config);
            
            // Display results
            this.displayAnalysis(this.currentAnalysis);
            this.showResponseSection();
            
            // Log successful analysis
            this.logger.logEvent('email_analyzed', {
                model_service: config.service,
                model_name: config.model,
                email_length: this.currentEmail.bodyLength,
                recipients_count: this.currentEmail.recipients.split(',').length
            });
            
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
            
            console.log('[TaskpaneApp] Response generated:', this.currentResponse);
            
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
        
        if (!customInstructions) {
            this.uiController.showError('Please provide instructions for refining the response.');
            return;
        }

        try {
            this.uiController.showStatus('Refining response...');
            this.uiController.setButtonLoading('refine-response', true);
            
            const config = this.getAIConfiguration();
            
            this.currentResponse = await this.aiService.refineResponse(
                this.currentResponse,
                customInstructions,
                config
            );
            
            this.displayResponse(this.currentResponse);
            this.uiController.showStatus('Response refined successfully.');
            
        } catch (error) {
            console.error('Response refinement failed:', error);
            this.uiController.showError('Failed to refine response. Please try again.');
        } finally {
            this.uiController.setButtonLoading('refine-response', false);
        }
    }

    getAIConfiguration() {
        let model = this.getSelectedModel();
        // Always use the unified modelSelect for model selection
        if (this.modelSelect && this.modelSelect.value) {
            model = this.modelSelect.value;
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
            'anthropic': 'claude-3-sonnet',
            'azure': 'gpt-4',
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
        container.innerHTML = `
            <div class="analysis-content">
                <h3>Key Points</h3>
                <ul>
                    ${analysis.keyPoints.map(point => `<li>${this.escapeHtml(point)}</li>`).join('')}
                </ul>

                <h3>Sentiment</h3>
                <ul>
                    <li>${this.escapeHtml(analysis.sentiment)}</li>
                </ul>

                <h3>Recommended Actions</h3>
                <ul>
                    ${analysis.actions.map(action => `<li>${this.escapeHtml(action)}</li>`).join('')}
                </ul>
            </div>
        `;
    }

    displayResponse(response) {
        console.log('[TaskpaneApp] Displaying response:', response);
        const container = document.getElementById('response-draft');
        
        if (!container) {
            console.error('[TaskpaneApp] response-draft container not found');
            return;
        }
        
        if (!response || !response.text) {
            console.error('[TaskpaneApp] Invalid response object:', response);
            container.innerHTML = '<div class="error">Error: Invalid response received</div>';
            return;
        }
        
        // Clean up the text for display (remove tabs, normalize whitespace)
        let cleanText = response.text.replace(/\t/g, '').trim();
        cleanText = cleanText.replace(/[ ]+/g, ' '); // Replace multiple spaces with single space
        
        container.innerHTML = `
            <div class="response-content">
                <h3>Generated Response</h3>
                <div class="response-text" id="response-text-content">
                    ${this.escapeHtml(cleanText).replace(/\n/g, '<br>')}
                </div>
            </div>
        `;
        
        console.log('[TaskpaneApp] Response displayed successfully');
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    async copyResponse() {
        try {
            const responseText = document.getElementById('response-text-content').textContent;
            await navigator.clipboard.writeText(responseText);
            this.uiController.showStatus('Response copied to clipboard.');
        } catch (error) {
            console.error('Failed to copy response:', error);
            this.uiController.showError('Failed to copy response to clipboard.');
        }
    }

    async insertResponse() {
        try {
            // Get the cleaned text from the displayed content instead of the raw response
            const responseTextElement = document.getElementById('response-text-content');
            const responseText = responseTextElement ? responseTextElement.textContent : '';
            
            console.log('[InsertResponse] Response text:', responseText);
            if (!responseText || responseText.trim().length === 0) {
                this.uiController.showError('No response text to insert.');
                return;
            }
            
            // Check Office.js and mailbox context
            if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox || !Office.context.mailbox.item || !Office.context.mailbox.item.body) {
                this.uiController.showError('Office.js mailbox context is not available. Please run this add-in inside Outlook.');
                console.error('[InsertResponse] Office.js mailbox context missing.');
                return;
            }
            
            // Format text for Outlook with proper line breaks
            const formattedText = this.formatTextForOutlook(responseText);
            console.log('[InsertResponse] Formatted text:', formattedText);
            
            // Try to use setSelectedDataAsync first (inserts at cursor position) with plain text
            if (typeof Office.context.mailbox.item.body.setSelectedDataAsync === 'function') {
                Office.context.mailbox.item.body.setSelectedDataAsync(
                    formattedText,
                    { coercionType: Office.CoercionType.Text },
                    (result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            this.uiController.showStatus('Response inserted at cursor position.');
                        } else {
                            console.error('[InsertResponse] setSelectedDataAsync error:', result);
                            // Fallback to HTML approach
                            this.insertResponseAsHtml(responseText);
                        }
                    }
                );
            } else {
                // Fallback to HTML approach
                this.insertResponseAsHtml(responseText);
            }
        } catch (error) {
            console.error('Failed to insert response:', error);
            this.uiController.showError('Failed to insert response into email.');
        }
    }

    /**
     * Format text for Outlook with proper line breaks
     * @param {string} text - Plain text to format
     * @returns {string} Formatted text with proper line breaks
     */
    formatTextForOutlook(text) {
        // Clean up the text first - remove tabs and normalize whitespace
        let formatted = text.replace(/\t/g, '').trim(); // Remove all tabs
        formatted = formatted.replace(/[ ]+/g, ' '); // Replace multiple spaces with single space
        
        // Use Windows-style line breaks
        formatted = formatted.replace(/\n/g, '\r\n');
        
        // Add extra spacing between existing paragraphs (convert double line breaks to triple)
        formatted = formatted.replace(/\r\n\r\n/g, '\r\n\r\n\r\n');
        
        // If there are no existing paragraph breaks, try to detect natural break points
        if (!formatted.includes('\r\n\r\n')) {
            // Look for common greeting patterns and add line breaks
            formatted = formatted.replace(/(Hi\s+\w+,)/gi, '$1\r\n\r\n');
            formatted = formatted.replace(/(Dear\s+\w+,)/gi, '$1\r\n\r\n');
            formatted = formatted.replace(/(Hello\s+\w+,)/gi, '$1\r\n\r\n');
            
            // Look for "Thanks for" patterns that often start new paragraphs
            formatted = formatted.replace(/\s+(Thanks for [^.!?]*[.!?])/gi, '\r\n\r\n$1');
            
            // Look for sentence endings followed by capital letters (likely paragraph breaks)
            formatted = formatted.replace(/([.!?])\s+([A-Z][a-z])/g, '$1\r\n\r\n$2');
            
            // Handle closing salutations - look for the pattern and ensure proper spacing
            // Match: sentence ending, optional space, salutation, optional comma, space, name
            formatted = formatted.replace(/([.!?])\s*(Best\s+regards?),?\s+(\w+(?:\s+\w+)*)/gi, '$1\r\n\r\n$2,\r\n$3');
            formatted = formatted.replace(/([.!?])\s*(Best),?\s+(\w+(?:\s+\w+)*)/gi, '$1\r\n\r\n$2,\r\n$3');
            formatted = formatted.replace(/([.!?])\s*(Sincerely),?\s+(\w+(?:\s+\w+)*)/gi, '$1\r\n\r\n$2,\r\n$3');
            formatted = formatted.replace(/([.!?])\s*(Thanks?),?\s+(\w+(?:\s+\w+)*)/gi, '$1\r\n\r\n$2,\r\n$3');
            formatted = formatted.replace(/([.!?])\s*(Regards?),?\s+(\w+(?:\s+\w+)*)/gi, '$1\r\n\r\n$2,\r\n$3');
            formatted = formatted.replace(/([.!?])\s*(Cheers),?\s+(\w+(?:\s+\w+)*)/gi, '$1\r\n\r\n$2,\r\n$3');
        }
        
        // Clean up any excessive line breaks (more than 3 consecutive)
        formatted = formatted.replace(/\r\n\r\n\r\n\r\n+/g, '\r\n\r\n\r\n');
        
        // Final cleanup - remove any leading/trailing whitespace
        formatted = formatted.trim();
        
        console.log('[formatTextForOutlook] Original:', JSON.stringify(text));
        console.log('[formatTextForOutlook] Formatted:', JSON.stringify(formatted));
        
        return formatted;
    }

    insertResponseAsHtml(responseText) {
        const htmlContent = this.convertTextToHtml(responseText);
        
        if (typeof Office.context.mailbox.item.body.setSelectedDataAsync === 'function') {
            Office.context.mailbox.item.body.setSelectedDataAsync(
                htmlContent,
                { coercionType: Office.CoercionType.Html },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        this.uiController.showStatus('Response inserted as HTML.');
                    } else {
                        console.error('[InsertResponse] HTML setSelectedDataAsync error:', result);
                        this.insertResponseFallback(htmlContent);
                    }
                }
            );
        } else {
            this.insertResponseFallback(htmlContent);
        }
    }

    insertResponseFallback(htmlContent) {
        Office.context.mailbox.item.body.setAsync(
            htmlContent,
            { coercionType: Office.CoercionType.Html },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    this.uiController.showStatus('Response inserted into email body.');
                } else {
                    console.error('[InsertResponse] setAsync error:', result);
                    this.uiController.showError('Failed to insert response into email.');
                }
            }
        );
    }

    /**
     * Converts plain text to HTML, preserving line breaks and paragraphs
     * @param {string} text - Plain text to convert
     * @returns {string} HTML formatted text
     */
    convertTextToHtml(text) {
        // Escape HTML entities first
        let html = text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
        
        // Split by double line breaks to identify paragraphs
        const paragraphs = html.split(/\n\s*\n/);
        const htmlParagraphs = [];
        
        for (let paragraph of paragraphs) {
            const trimmed = paragraph.trim();
            if (trimmed) {
                // Replace single line breaks with <br> and wrap in <p> with Outlook-friendly styling
                const formattedParagraph = trimmed.replace(/\n/g, '<br>');
                htmlParagraphs.push(`<p style="margin-top: 0; margin-bottom: 1em; line-height: 1.5;">${formattedParagraph}</p>`);
            }
        }
        
        // If no paragraphs detected, treat as single paragraph
        if (htmlParagraphs.length === 0 && html.trim()) {
            const formattedText = html.trim().replace(/\n/g, '<br>');
            htmlParagraphs.push(`<p style="margin-top: 0; margin-bottom: 1em; line-height: 1.5;">${formattedText}</p>`);
        }
        
        // Create the final HTML
        const finalHtml = htmlParagraphs.join('');
        
        console.log('[convertTextToHtml] Converted HTML:', finalHtml);
        return finalHtml;
    }

    showRefineButton() {
        document.getElementById('refine-response').classList.remove('hidden');
    }

    onModelServiceChange(event) {
        const customEndpoint = document.getElementById('custom-endpoint');
        if (customEndpoint) {
            if (event.target.value === 'custom') {
                customEndpoint.classList.remove('hidden');
            } else {
                customEndpoint.classList.add('hidden');
            }
        }
        this.saveSettings();
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

    loadSettingsIntoUI() {
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
        inputs.forEach(input => {
            if (input.id) {
                settings[input.id] = input.type === 'checkbox' ? input.checked : input.value;
            }
        });
        
        this.settingsManager.saveSettings(settings);
    }

    getUserId() {
        // In a real implementation, this would get the actual user ID
        return Office.context.mailbox.userProfile.emailAddress || 'unknown';
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
