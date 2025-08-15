/**
 * AI Service for email analysis and response generation
 * Supports multiple AI providers and models
 */

export class AIService {
    /**
     * Extracts the response text from the API response data for each service
     * @param {Object} data - The response data from the API
     * @param {string} service - The AI service name
     * @returns {string} The extracted response text
     */
    extractResponseText(data, service) {
        switch (service) {
            case 'openai':
            case 'azure':
                return data.choices?.[0]?.message?.content || '';
            case 'ollama':
                // Ollama returns response in data.message.content
                return data.message?.content || data.response || data.text || JSON.stringify(data);
            case 'anthropic':
                return data.content?.[0]?.text || '';
            default:
                // Fallback: try common fields or stringify
                return data.response || data.text || data.content || JSON.stringify(data);
        }
    }
    constructor() {
        this.supportedServices = {
            openai: {
                endpoint: 'https://api.openai.com/v1/chat/completions',
                model: 'gpt-4',
                maxTokens: 4000
            },
            ollama: {
                endpoint: 'http://localhost:11434/api/chat', // Default Ollama base_url
                model: 'llama2', // Default Ollama model
                apiKey: '', // Ollama does not require apiKey by default
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
    /**
     * Fetch available models from Ollama using /api/tags
     * @param {string} baseUrl - The base URL for Ollama
     * @returns {Promise<Array>} - Array of model names
     */
    static async fetchOllamaModels(baseUrl) {
        try {
            const url = `${baseUrl.replace(/\/$/, '')}/api/tags`;
            const response = await fetch(url);
            if (!response.ok) throw new Error(`Failed to fetch models: ${response.status}`);
            const data = await response.json();
            // Ollama returns { models: [{ name: ... }, ...] }
            return (data.models || []).map(m => m.name);
        } catch (err) {
            console.error('Error fetching Ollama models:', err);
            return [];
        }
    }
    
    async analyzeEmail(emailData, config) {
        console.log('[AIService] Starting email analysis...');
        console.log('[AIService] Email data:', emailData);
        console.log('[AIService] Config:', config);
        
        const prompt = this.buildAnalysisPrompt(emailData);
        console.log('[AIService] Built analysis prompt:', prompt);
        
        try {
            console.log('[AIService] Calling AI for analysis...');
            const response = await this.callAI(prompt, config, 'analysis');
            console.log('[AIService] Raw analysis response:', response);
            
            const parsed = this.parseAnalysisResponse(response);
            console.log('[AIService] Parsed analysis result:', parsed);
            return parsed;
        } catch (error) {
            console.error('[AIService] Email analysis failed:', error);
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
        console.log('[AIService] Starting response generation...');
        console.log('[AIService] Email data:', emailData);
        console.log('[AIService] Analysis:', analysis);
        console.log('[AIService] Config:', config);
        
        // Ensure analysis is not null - provide default if missing
        if (!analysis) {
            console.warn('[AIService] Analysis is null, providing default analysis structure');
            analysis = {
                keyPoints: ['No analysis available'],
                sentiment: 'neutral',
                responseStrategy: 'respond professionally'
            };
        }
        
        const prompt = this.buildResponsePrompt(emailData, analysis, config);
        console.log('[AIService] Built response prompt:', prompt);
        
        try {
            console.log('[AIService] Calling AI for response generation...');
            const response = await this.callAI(prompt, config, 'response');
            console.log('[AIService] Raw response generation result:', response);
            
            const parsed = this.parseResponseResult(response);
            console.log('[AIService] Parsed response result:', parsed);
            return parsed;
        } catch (error) {
            console.error('[AIService] Response generation failed:', error);
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
        const dateStr = emailData.date ? new Date(emailData.date).toLocaleString() : 'Compose Mode';
        return `Please analyze the following email and provide insights:

**Email Details:**
From: ${emailData.from}
Subject: ${emailData.subject}
Recipients: ${emailData.recipients}
Sent: ${dateStr}
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

        let prompt = `You are replying as: ${emailData.sender || 'Unknown Sender'} to: ${emailData.from}\n\n` +
            `Generate a professional email response based on the following context:\n\n` +
            `**Original Email:**\n` +
            `From: ${emailData.from}\n` +
            `Subject: ${emailData.subject}\n` +
            `Sent: ${emailData.date ? new Date(emailData.date).toLocaleString() : 'Compose Mode'}\n` +
            `Content: ${emailData.cleanBody || emailData.body}\n\n` +
            `**Analysis Summary:**\n` +
            `- Key Points: ${(analysis && analysis.keyPoints) ? analysis.keyPoints.join(', ') : 'Not analyzed'}\n` +
            `- Sentiment: ${(analysis && analysis.sentiment) || 'Not analyzed'}\n` +
            `- Recommended Strategy: ${(analysis && analysis.responseStrategy) || 'Not analyzed'}\n\n` +
            `**Response Requirements:**\n` +
            `- Length: ${lengthMap[config.length] || 'medium length'}\n` +
            `- Tone: ${toneMap[config.tone] || 'professional'}\n` +
            `- Urgency: ${urgencyMap[config.urgency] || 'appropriate'}`;

        if (config.customInstructions && config.customInstructions.trim()) {
            prompt += `\n- Special Instructions: ${config.customInstructions}`;
        }

        prompt += `\n\n**Output Requirements:**\n` +
            `Please generate an appropriate email response that:\n` +
            `1. Addresses the key points from the original email\n` +
            `2. Matches the requested tone and length\n` +
            `3. Is professional and well-structured\n` +
            `4. Includes appropriate greetings and closings\n` +
            `5. Uses proper paragraph formatting with blank lines (double newlines) between paragraphs.\n\n` +
            `Return only the body of the email response, ready to be sent. Do not include subject line, email headers, or any introductory phrases such as 'Here's a suggested email response:' or similar. Output only the email content as it should appear in the reply.`;

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
        console.log(`[AIService] Starting AI call for type: ${type}`);
        console.log('[AIService] Prompt:', prompt);
        console.log('[AIService] Config:', config);
        
        const service = config.service || 'openai';
        console.log(`[AIService] Using service: ${service}`);

        if (service === 'custom') {
            console.log('[AIService] Calling custom endpoint...');
            return this.callCustomEndpoint(prompt, config);
        }

        const serviceConfig = this.supportedServices[service];
        if (!serviceConfig) {
            console.error(`[AIService] Unsupported AI service: ${service}`);
            throw new Error(`Unsupported AI service: ${service}`);
        }
        console.log('[AIService] Service config:', serviceConfig);

        let endpoint = serviceConfig.endpoint;
        if (service === 'azure' && config.endpointUrl) {
            endpoint = config.endpointUrl;
            console.log('[AIService] Using Azure custom endpoint:', endpoint);
        }
        if (service === 'ollama' && config.baseUrl) {
            // Always use /api/chat or /api/generate, not the base URL
            const base = config.baseUrl.replace(/\/$/, '');
            endpoint = `${base}/api/chat`;
            console.log('[AIService] Using Ollama endpoint:', endpoint);
        }
        console.log('[AIService] Final endpoint:', endpoint);

        let requestBody;
        let headers;

        if (service === 'ollama') {
            requestBody = {
                model: config.model || serviceConfig.model,
                messages: [{ role: 'user', content: prompt }],
                stream: false
            };
            headers = { 'Content-Type': 'application/json' };
            console.log('[AIService] Built Ollama request body:', requestBody);
        } else {
            requestBody = this.buildRequestBody(prompt, service, config);
            headers = this.buildHeaders(service, config);
            console.log('[AIService] Built request body:', requestBody);
        }
        console.log('[AIService] Request headers:', headers);

        // Debug: Log POST request body to console
        console.log('[AIService] Making API call to endpoint:', endpoint);
        console.log('[AIService] Request Body:', requestBody);
        let response = await fetch(endpoint, {
            method: 'POST',
            headers: headers,
            body: JSON.stringify(requestBody)
        });
        console.log(`[AIService] Got response with status: ${response.status} ${response.statusText}`);

        // Fallback to /api/generate if /api/chat fails with 405
        if (service === 'ollama' && response.status === 405) {
            const base = config.baseUrl.replace(/\/$/, '');
            const fallbackEndpoint = `${base}/api/generate`;
            console.warn('[AIService] Ollama /api/chat failed with 405, retrying with /api/generate:', fallbackEndpoint);
            response = await fetch(fallbackEndpoint, {
                method: 'POST',
                headers: headers,
                body: JSON.stringify(requestBody)
            });
            console.log(`[AIService] Fallback response status: ${response.status} ${response.statusText}`);
        }

        if (!response.ok) {
            const errorText = await response.text();
            console.error(`[AIService] API request failed: ${response.status} ${response.statusText}`);
            console.error('[AIService] Error response:', errorText);
            throw new Error(`API request failed: ${response.status} ${response.statusText} - ${errorText}`);
        }

        console.log('[AIService] Response OK, parsing JSON...');
        const data = await response.json();
        console.log('[AIService] Response data:', data);
        
        const extractedText = this.extractResponseText(data, service);
        console.log('[AIService] Extracted response text:', extractedText);
        return extractedText;
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
            // Ollama and anthropic do not require Authorization header by default
        }
        return headers;
    }

    /**
     * Parses analysis response from AI
     * @param {string} responseText - Raw response text
     * @returns {Object} Parsed analysis
     */
    parseAnalysisResponse(responseText) {
        // Try to extract and parse the first JSON object from the response
        try {
            let jsonMatch = responseText.match(/\{[\s\S]*\}/);
            if (jsonMatch) {
                const parsed = JSON.parse(jsonMatch[0]);
                return {
                    keyPoints: parsed.keyPoints || [],
                    sentiment: parsed.sentiment || 'Unable to determine',
                    intent: parsed.intent || 'Unable to determine',
                    urgencyLevel: parsed.urgencyLevel || 3,
                    urgencyReason: parsed.urgencyReason || 'Standard priority',
                    actions: parsed.actions || [],
                    responseStrategy: parsed.responseStrategy || 'Respond professionally'
                };
            }
        } catch (error) {
            // Ignore and fallback
        }
        // Fallback to text parsing
        return this.parseAnalysisFromText(responseText);
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
        // Normalize newlines and clean up whitespace issues
        let text = responseText.trim();
        
        // Remove tabs and normalize them to spaces
        text = text.replace(/\t/g, ' ');
        
        // Convert \r\n and \r to \n
        text = text.replace(/\r\n?/g, '\n');
        
        // Clean up multiple spaces (but preserve intentional spacing)
        text = text.replace(/[ ]{2,}/g, ' ');
        
        // Replace 3+ newlines with exactly two
        text = text.replace(/\n{3,}/g, '\n\n');
        
        // Trim leading/trailing whitespace on each line and remove empty lines at start/end
        const lines = text.split('\n').map(line => line.trim());
        
        // Remove empty lines from the beginning and end
        while (lines.length > 0 && lines[0] === '') {
            lines.shift();
        }
        while (lines.length > 0 && lines[lines.length - 1] === '') {
            lines.pop();
        }
        
        text = lines.join('\n');
        
        return {
            text,
            generatedAt: new Date().toISOString(),
            wordCount: text.split(/\s+/).filter(word => word.length > 0).length
        };
    }
}
