/**
 * AI Service for email analysis and response generation
 * Supports multiple AI providers and models
 */

export class AIService {
    constructor() {
        this.supportedServices = {
            openai: {
                endpoint: 'https://api.openai.com/v1/chat/completions',
                model: 'gpt-4',
                maxTokens: 4000
            },
            anthropic: {
                endpoint: 'https://api.anthropic.com/v1/messages',
                model: 'claude-3-sonnet-20240229',
                maxTokens: 4000
            },
            azure: {
                endpoint: '', // Will be set from configuration
                model: 'gpt-4',
                maxTokens: 4000
            }
        };
    }

    /**
     * Analyzes an email using AI
     * @param {Object} emailData - Email data from EmailAnalyzer
     * @param {Object} config - AI configuration
     * @returns {Promise<Object>} Analysis results
     */
    async analyzeEmail(emailData, config) {
        const prompt = this.buildAnalysisPrompt(emailData);
        
        try {
            const response = await this.callAI(prompt, config, 'analysis');
            return this.parseAnalysisResponse(response);
        } catch (error) {
            console.error('Email analysis failed:', error);
            throw new Error('Failed to analyze email: ' + error.message);
        }
    }

    /**
     * Generates a response to an email
     * @param {Object} emailData - Original email data
     * @param {Object} analysis - Email analysis results
     * @param {Object} config - Configuration including AI and response settings
     * @returns {Promise<Object>} Generated response
     */
    async generateResponse(emailData, analysis, config) {
        const prompt = this.buildResponsePrompt(emailData, analysis, config);
        
        try {
            const response = await this.callAI(prompt, config, 'response');
            return this.parseResponseResult(response);
        } catch (error) {
            console.error('Response generation failed:', error);
            throw new Error('Failed to generate response: ' + error.message);
        }
    }

    /**
     * Refines an existing response based on user feedback
     * @param {Object} currentResponse - Current response object
     * @param {string} instructions - User refinement instructions
     * @param {Object} config - AI configuration
     * @returns {Promise<Object>} Refined response
     */
    async refineResponse(currentResponse, instructions, config) {
        const prompt = this.buildRefinementPrompt(currentResponse, instructions);
        
        try {
            const response = await this.callAI(prompt, config, 'refinement');
            return this.parseResponseResult(response);
        } catch (error) {
            console.error('Response refinement failed:', error);
            throw new Error('Failed to refine response: ' + error.message);
        }
    }

    /**
     * Builds the prompt for email analysis
     * @param {Object} emailData - Email data
     * @returns {string} Analysis prompt
     */
    buildAnalysisPrompt(emailData) {
        return `Please analyze the following email and provide insights:

**Email Details:**
From: ${emailData.from}
Subject: ${emailData.subject}
Recipients: ${emailData.recipients}
Length: ${emailData.bodyLength} characters

**Email Content:**
${emailData.cleanBody || emailData.body}

**Analysis Request:**
Please provide a structured analysis including:

1. **Key Points**: List the main points or topics discussed (3-5 bullet points)
2. **Sentiment**: Describe the overall tone and sentiment of the email
3. **Intent**: What is the sender trying to accomplish?
4. **Urgency Level**: Rate the urgency from 1-5 and explain why
5. **Action Items**: What actions are requested or implied?
6. **Recommended Response Strategy**: How should this email be approached in a response?

Format your response as JSON with the following structure:
{
    "keyPoints": ["point1", "point2", "point3"],
    "sentiment": "description of sentiment and tone",
    "intent": "what the sender wants to accomplish",
    "urgencyLevel": number,
    "urgencyReason": "explanation of urgency rating",
    "actions": ["action1", "action2"],
    "responseStrategy": "recommended approach for responding"
}`;
    }

    /**
     * Builds the prompt for response generation
     * @param {Object} emailData - Original email data
     * @param {Object} analysis - Email analysis
     * @param {Object} config - Response configuration
     * @returns {string} Response generation prompt
     */
    buildResponsePrompt(emailData, analysis, config) {
        const lengthMap = {
            1: 'very brief (1-2 sentences)',
            2: 'brief (1 short paragraph)',
            3: 'medium length (2-3 paragraphs)',
            4: 'detailed (3-4 paragraphs)',
            5: 'very detailed (4+ paragraphs)'
        };

        const toneMap = {
            1: 'very casual and friendly',
            2: 'casual but respectful',
            3: 'professional and courteous',
            4: 'formal and business-like',
            5: 'very formal and ceremonious'
        };

        const urgencyMap = {
            1: 'relaxed, no rush',
            2: 'calm and measured',
            3: 'appropriate responsiveness',
            4: 'prompt and attentive',
            5: 'urgent and immediate'
        };

        let prompt = `Generate a professional email response based on the following context:

**Original Email:**
From: ${emailData.from}
Subject: ${emailData.subject}
Content: ${emailData.cleanBody || emailData.body}

**Analysis Summary:**
- Key Points: ${analysis.keyPoints?.join(', ') || 'Not analyzed'}
- Sentiment: ${analysis.sentiment || 'Not analyzed'}
- Recommended Strategy: ${analysis.responseStrategy || 'Not analyzed'}

**Response Requirements:**
- Length: ${lengthMap[config.length] || 'medium length'}
- Tone: ${toneMap[config.tone] || 'professional'}
- Urgency: ${urgencyMap[config.urgency] || 'appropriate'}`;

        if (config.customInstructions && config.customInstructions.trim()) {
            prompt += `\n- Special Instructions: ${config.customInstructions}`;
        }

        prompt += `\n\n**Output Requirements:**
Please generate an appropriate email response that:
1. Addresses the key points from the original email
2. Matches the requested tone and length
3. Is professional and well-structured
4. Includes appropriate greetings and closings

Return only the email response text, ready to be sent. Do not include subject line or email headers.`;

        return prompt;
    }

    /**
     * Builds the prompt for response refinement
     * @param {Object} currentResponse - Current response
     * @param {string} instructions - User instructions
     * @returns {string} Refinement prompt
     */
    buildRefinementPrompt(currentResponse, instructions) {
        return `Please refine the following email response based on the user's feedback:

**Current Response:**
${currentResponse.text}

**User's Refinement Instructions:**
${instructions}

**Requirements:**
- Apply the user's feedback while maintaining professionalism
- Keep the overall structure and flow intact unless specifically requested to change
- Ensure the response remains appropriate for business communication
- Maintain consistency in tone and style

Please provide the refined email response text only.`;
    }

    /**
     * Makes API call to the specified AI service
     * @param {string} prompt - The prompt to send
     * @param {Object} config - AI configuration
     * @param {string} type - Type of request (analysis, response, refinement)
     * @returns {Promise<string>} AI response text
     */
    async callAI(prompt, config, type) {
        const service = config.service || 'openai';
        
        if (service === 'custom') {
            return this.callCustomEndpoint(prompt, config);
        }
        
        const serviceConfig = this.supportedServices[service];
        if (!serviceConfig) {
            throw new Error(`Unsupported AI service: ${service}`);
        }

        const requestBody = this.buildRequestBody(prompt, service, config);
        const headers = this.buildHeaders(service, config);

        const endpoint = service === 'azure' ? config.endpointUrl : serviceConfig.endpoint;
        
        const response = await fetch(endpoint, {
            method: 'POST',
            headers: headers,
            body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`API request failed: ${response.status} ${response.statusText} - ${errorText}`);
        }

        const data = await response.json();
        return this.extractResponseText(data, service);
    }

    /**
     * Calls a custom AI endpoint
     * @param {string} prompt - The prompt
     * @param {Object} config - Configuration
     * @returns {Promise<string>} Response text
     */
    async callCustomEndpoint(prompt, config) {
        if (!config.endpointUrl) {
            throw new Error('Custom endpoint URL is required');
        }

        const requestBody = {
            prompt: prompt,
            max_tokens: 4000,
            temperature: 0.7
        };

        const headers = {
            'Content-Type': 'application/json'
        };

        if (config.apiKey) {
            headers['Authorization'] = `Bearer ${config.apiKey}`;
        }

        const response = await fetch(config.endpointUrl, {
            method: 'POST',
            headers: headers,
            body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
            throw new Error(`Custom endpoint request failed: ${response.status}`);
        }

        const data = await response.json();
        
        // Try to extract response from common response formats
        return data.response || data.text || data.content || JSON.stringify(data);
    }

    /**
     * Builds request body based on AI service
     * @param {string} prompt - The prompt
     * @param {string} service - AI service name
     * @param {Object} config - Configuration
     * @returns {Object} Request body
     */
    buildRequestBody(prompt, service, config) {
        const serviceConfig = this.supportedServices[service];
        
        switch (service) {
            case 'openai':
            case 'azure':
                return {
                    model: config.model || serviceConfig.model,
                    messages: [
                        {
                            role: 'system',
                            content: 'You are a helpful AI assistant that specializes in email analysis and response generation. Provide clear, professional, and actionable insights.'
                        },
                        {
                            role: 'user',
                            content: prompt
                        }
                    ],
                    max_tokens: serviceConfig.maxTokens,
                    temperature: 0.7
                };
                
            case 'anthropic':
                return {
                    model: config.model || serviceConfig.model,
                    max_tokens: serviceConfig.maxTokens,
                    messages: [
                        {
                            role: 'user',
                            content: prompt
                        }
                    ],
                    temperature: 0.7
                };
                
            default:
                throw new Error(`Unsupported service for request body: ${service}`);
        }
    }

    /**
     * Builds headers for API request
     * @param {string} service - AI service name
     * @param {Object} config - Configuration
     * @returns {Object} Headers object
     */
    buildHeaders(service, config) {
        const headers = {
            'Content-Type': 'application/json'
        };

        switch (service) {
            case 'openai':
            case 'azure':
                headers['Authorization'] = `Bearer ${config.apiKey}`;
                break;
                
            case 'anthropic':
                headers['x-api-key'] = config.apiKey;
                headers['anthropic-version'] = '2023-06-01';
                break;
        }

        return headers;
    }

    /**
     * Extracts response text from API response
     * @param {Object} data - API response data
     * @param {string} service - AI service name
     * @returns {string} Response text
     */
    extractResponseText(data, service) {
        switch (service) {
            case 'openai':
            case 'azure':
                return data.choices?.[0]?.message?.content || '';
                
            case 'anthropic':
                return data.content?.[0]?.text || '';
                
            default:
                throw new Error(`Unsupported service for response extraction: ${service}`);
        }
    }

    /**
     * Parses analysis response from AI
     * @param {string} responseText - Raw response text
     * @returns {Object} Parsed analysis
     */
    parseAnalysisResponse(responseText) {
        try {
            // Try to parse as JSON first
            const parsed = JSON.parse(responseText);
            return {
                keyPoints: parsed.keyPoints || [],
                sentiment: parsed.sentiment || 'Unable to determine',
                intent: parsed.intent || 'Unable to determine',
                urgencyLevel: parsed.urgencyLevel || 3,
                urgencyReason: parsed.urgencyReason || 'Standard priority',
                actions: parsed.actions || [],
                responseStrategy: parsed.responseStrategy || 'Respond professionally'
            };
        } catch (error) {
            // Fallback to text parsing
            return this.parseAnalysisFromText(responseText);
        }
    }

    /**
     * Fallback parsing for non-JSON analysis responses
     * @param {string} text - Response text
     * @returns {Object} Parsed analysis
     */
    parseAnalysisFromText(text) {
        return {
            keyPoints: ['Analysis completed', 'See full response for details'],
            sentiment: 'Professional communication',
            intent: 'Information sharing',
            urgencyLevel: 3,
            urgencyReason: 'Standard business communication',
            actions: ['Review content', 'Respond appropriately'],
            responseStrategy: text.substring(0, 200) + '...'
        };
    }

    /**
     * Parses response generation result
     * @param {string} responseText - Generated response text
     * @returns {Object} Response object
     */
    parseResponseResult(responseText) {
        return {
            text: responseText.trim(),
            generatedAt: new Date().toISOString(),
            wordCount: responseText.trim().split(/\s+/).length
        };
    }
}
