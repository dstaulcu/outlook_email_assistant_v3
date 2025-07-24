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
        this.baseUrlInput = null;
        this.modelSelectGroup = null;
        this.modelSelect = null;
    }

    async initialize() {
        try {
            // Initialize Office.js
            await this.initializeOffice();
            
            // Load user settings
            await this.settingsManager.loadSettings();
            
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
        this.baseUrlInput = document.getElementById('base-url');
        this.modelSelectGroup = document.getElementById('model-select-group');
        this.modelSelect = document.getElementById('model-select');

        // Wire up model discovery
        if (this.modelServiceSelect && this.baseUrlInput && this.modelSelectGroup && this.modelSelect) {
            this.modelServiceSelect.addEventListener('change', () => this.updateModelDropdown());
            this.baseUrlInput.addEventListener('change', () => this.updateModelDropdown());
            this.baseUrlInput.addEventListener('blur', () => this.updateModelDropdown());
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
        document.getElementById('email-from').textContent = email.from || 'Unknown';
        document.getElementById('email-subject').textContent = email.subject || 'No Subject';
        document.getElementById('email-recipients').textContent = email.recipients || 'Unknown';
        document.getElementById('email-length').textContent = `${email.bodyLength || 0} characters`;
        // Classification display logic
        let classification = email.classification;
        let classificationText;
        if (!classification || classification.toLowerCase() === "unclassified") {
            classificationText = "This email appears to be safe for AI processing.";
        } else {
            classificationText = classification;
        }
        document.getElementById("email-classification").textContent = classificationText;
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
            
            // Generate response
            this.currentResponse = await this.aiService.generateResponse(
                this.currentEmail, 
                this.currentAnalysis,
                { ...config, ...responseConfig }
            );
            
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
        return {
            service: this.modelServiceSelect ? this.modelServiceSelect.value : '',
            apiKey: document.getElementById('api-key').value,
            endpointUrl: document.getElementById('endpoint-url').value,
            baseUrl: this.baseUrlInput ? this.baseUrlInput.value : '',
            model: (this.modelServiceSelect && this.modelServiceSelect.value === 'ollama' && this.modelSelect && this.modelSelect.value) ? this.modelSelect.value : this.getSelectedModel()
        };
    }

    getResponseConfiguration() {
        return {
            length: parseInt(document.getElementById('response-length').value),
            tone: parseInt(document.getElementById('response-tone').value),
            urgency: parseInt(document.getElementById('response-urgency').value),
            customInstructions: document.getElementById('custom-instructions').value
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
        if (!this.modelServiceSelect || !this.baseUrlInput || !this.modelSelectGroup || !this.modelSelect) return;
        if (this.modelServiceSelect.value === 'ollama') {
            this.modelSelectGroup.style.display = '';
            this.modelSelect.innerHTML = '<option value="">Loading...</option>';
            const baseUrl = this.baseUrlInput.value || 'http://localhost:11434';
            const requestUrl = `${baseUrl.replace(/\/$/, '')}/api/tags`;
            console.log('[Ollama Model Discovery] Requesting:', requestUrl);
            let models = [];
            try {
                models = await AIService.fetchOllamaModels(baseUrl);
                console.log('[Ollama Model Discovery] Models received:', models);
                this.modelSelect.innerHTML = models.length
                    ? models.map(m => `<option value="${m}">${m}</option>`).join('')
                    : '<option value="">No models found</option>';
                // Remove any previous error message
                const errorDiv = document.getElementById('ollama-model-error');
                if (errorDiv) errorDiv.remove();
            } catch (err) {
                console.error('[Ollama Model Discovery] Error:', err);
                this.modelSelect.innerHTML = '<option value="">Error fetching models</option>';
                // Show error in UI
                let errorDiv = document.getElementById('ollama-model-error');
                if (!errorDiv) {
                    errorDiv = document.createElement('div');
                    errorDiv.id = 'ollama-model-error';
                    errorDiv.style.color = 'red';
                    this.modelSelectGroup.appendChild(errorDiv);
                }
                errorDiv.textContent = `Error fetching models: ${err.message || err}`;
            }
        } else {
            this.modelSelectGroup.style.display = 'none';
            this.modelSelect.innerHTML = '';
            const errorDiv = document.getElementById('ollama-model-error');
            if (errorDiv) errorDiv.remove();
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
                <p>${this.escapeHtml(analysis.sentiment)}</p>
                
                <h3>Recommended Actions</h3>
                <ul>
                    ${analysis.actions.map(action => `<li>${this.escapeHtml(action)}</li>`).join('')}
                </ul>
            </div>
        `;
    }

    displayResponse(response) {
        const container = document.getElementById('response-draft');
        container.innerHTML = `
            <div class="response-content">
                <h3>Generated Response</h3>
                <div class="response-text" id="response-text-content">
                    ${this.escapeHtml(response.text).replace(/\n/g, '<br>')}
                </div>
            </div>
        `;
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
            const responseText = document.getElementById('response-text-content').textContent;
            
            // Use Office.js to insert into email body
            Office.context.mailbox.item.body.setAsync(
                responseText,
                { coercionType: Office.CoercionType.Text },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        this.uiController.showStatus('Response inserted into email.');
                    } else {
                        this.uiController.showError('Failed to insert response into email.');
                    }
                }
            );
        } catch (error) {
            console.error('Failed to insert response:', error);
            this.uiController.showError('Failed to insert response into email.');
        }
    }

    showResponseSection() {
        document.getElementById('response-section').classList.remove('hidden');
    }

    switchToResponseTab() {
        document.getElementById('tab-response').click();
    }

    showRefineButton() {
        document.getElementById('refine-response').classList.remove('hidden');
    }

    onModelServiceChange(event) {
        const customEndpoint = document.getElementById('custom-endpoint');
        if (event.target.value === 'custom') {
            customEndpoint.classList.remove('hidden');
        } else {
            customEndpoint.classList.add('hidden');
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
            const element = document.getElementById(key);
            if (element) {
                if (element.type === 'checkbox') {
                    element.checked = settings[key];
                } else {
                    element.value = settings[key] || '';
                }
            }
        });
        
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
});
