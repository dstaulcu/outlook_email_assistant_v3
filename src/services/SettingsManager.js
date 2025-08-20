/**
 * Settings Manager Service
 * Handles persistent storage and retrieval of user preferences
 */

export class SettingsManager {
    constructor() {
        this.storageKey = 'promptemail_settings';
        this.defaultSettings = {
            // AI Configuration
            'model-service': 'openai',
            'api-key': '', // Legacy single API key for backwards compatibility
            'endpoint-url': '', // Legacy single endpoint for backwards compatibility
            
            // Provider-specific configurations
            'provider-configs': {
                'openai': { 'api-key': '', 'endpoint-url': 'https://api.openai.com/v1' },
                'ollama': { 'api-key': '', 'endpoint-url': 'http://localhost:11434' },
                'onsite1': { 'api-key': '', 'endpoint-url': '' }, // Empty to use baseUrl from ai-providers.json
                'onsite2': { 'api-key': '', 'endpoint-url': '' }  // Empty to use baseUrl from ai-providers.json
            },
            
            // Response Preferences
            'response-length': '1',
            'response-tone': '1',
            'response-urgency': '1',
            'custom-instructions': '',
            
            // Accessibility Settings
            'high-contrast': false,
            'screen-reader-mode': false,
            
            // UI Preferences
            'last-tab': 'analysis',
            'show-advanced-options': false,
            
            // Version tracking
            'settings-version': '1.0.0',
            'last-updated': null
        };
        
        this.settings = { ...this.defaultSettings };
        this.changeListeners = [];
        
        // Add unique instance ID for debugging
        this.instanceId = Date.now() + '-' + Math.random().toString(36).substr(2, 9);
        console.debug('SettingsManager instance created:', this.instanceId);
    }

    /**
     * Loads settings from storage
     * @returns {Promise<Object>} Loaded settings
     */
    async loadSettings() {
        console.debug('Starting settings load process...');
        try {
            console.debug('Attempting to load from Office storage...');
            // Try Office.js RoamingSettings first
            const officeSettings = await this.loadFromOfficeStorage();
            if (officeSettings) {
                console.debug('Successfully loaded from Office storage:', officeSettings);
                this.settings = { ...this.defaultSettings, ...officeSettings };
                console.debug('Merged settings with defaults:', this.settings);
                
                // Migrate legacy settings if needed
                const migrated = this.migrateLegacySettings();
                if (migrated) {
                    await this.saveSettings(this.settings);
                    console.debug('Legacy settings migrated and saved');
                }
                
                return this.settings;
            }
            console.warn('No Office storage settings found, trying localStorage...');

            // Fallback to localStorage
            const localSettings = this.loadFromLocalStorage();
            if (localSettings) {
                console.debug('Successfully loaded from localStorage:', localSettings);
                this.settings = { ...this.defaultSettings, ...localSettings };
                console.debug('Merged settings with defaults:', this.settings);
                
                // Migrate legacy settings if needed
                const migrated = this.migrateLegacySettings();
                if (migrated) {
                    await this.saveSettings(this.settings);
                    console.debug('Legacy settings migrated and saved');
                }
                
                return this.settings;
            }
            console.debug('No localStorage settings found, using defaults...');

            // No stored settings found, use defaults
            this.settings = { ...this.defaultSettings };
            console.debug('Using default settings:', this.settings);
            await this.saveSettings(this.settings);
            console.debug('Default settings saved successfully');
            
            return this.settings;

        } catch (error) {
            console.error('Failed to load settings:', error);
            console.debug('Falling back to default settings due to error');
            this.settings = { ...this.defaultSettings };
            console.debug('Using default settings after error:', this.settings);
            return this.settings;
        }
    }

    /**
     * Saves settings to storage
     * @param {Object} newSettings - Settings to save
     * @returns {Promise<boolean>} Success status
     */
    async saveSettings(newSettings = null) {
        console.debug('Starting settings save process...');
        console.debug('Settings to save:', newSettings || this.settings);
        try {
            const settingsToSave = newSettings || this.settings;
            
            // Update timestamp
            settingsToSave['last-updated'] = new Date().toISOString();
            console.debug('Added timestamp:', settingsToSave['last-updated']);
            
            // Update internal settings
            this.settings = { ...settingsToSave };
            console.debug('Updated internal settings cache');

            // Save to Office.js RoamingSettings
            console.debug('Attempting to save to Office storage...');
            const officeSaved = await this.saveToOfficeStorage(settingsToSave);
            console.debug(`[SettingsManager] Office storage save result: ${officeSaved ? 'SUCCESS' : 'FAILED'}`);
            
            // Also save to localStorage as backup
            console.debug('Saving to localStorage as backup...');
            this.saveToLocalStorage(settingsToSave);
            console.debug('localStorage backup save completed');

            // Notify listeners
            console.debug('Notifying change listeners...');
            this.notifyChangeListeners(settingsToSave);
            console.debug(`[SettingsManager] Notified ${this.changeListeners.length} listeners`);

            return officeSaved;

        } catch (error) {
            console.error('Failed to save settings:', error);
            console.log('Save operation failed, returning false');
            return false;
        }
    }

    /**
     * Loads settings from Office.js RoamingSettings
     * @returns {Promise<Object|null>} Settings object or null
     */
    async loadFromOfficeStorage() {
        console.debug('Loading from Office.js RoamingSettings...');
        return new Promise((resolve) => {
            try {
                if (typeof Office === 'undefined' || !Office.context?.roamingSettings) {
                    console.warn('Office.js or RoamingSettings not available');
                    resolve(null);
                    return;
                }

                const roamingSettings = Office.context.roamingSettings;
                console.debug('RoamingSettings object available');
                const settingsJson = roamingSettings.get(this.storageKey);
                console.debug('Raw settings from Office storage:', settingsJson);
                
                if (settingsJson) {
                    const settings = JSON.parse(settingsJson);
                    console.debug('Parsed Office settings:', settings);
                    resolve(settings);
                } else {
                    console.debug('No settings found in Office storage');
                    resolve(null);
                }

            } catch (error) {
                console.error('Failed to load from Office storage:', error);
                resolve(null);
            }
        });
    }

    /**
     * Saves settings to Office.js RoamingSettings
     * @param {Object} settings - Settings to save
     * @returns {Promise<boolean>} Success status
     */
    async saveToOfficeStorage(settings) {
        console.debug('Saving to Office.js RoamingSettings...');
        console.debug('Settings to serialize:', settings);
        return new Promise((resolve) => {
            try {
                if (typeof Office === 'undefined' || !Office.context?.roamingSettings) {
                    console.warn('Office.js or RoamingSettings not available for save');
                    resolve(false);
                    return;
                }

                const roamingSettings = Office.context.roamingSettings;
                console.debug('RoamingSettings object available for save');
                const settingsJson = JSON.stringify(settings);
                console.debug('Serialized settings JSON:', settingsJson);
                
                roamingSettings.set(this.storageKey, settingsJson);
                console.debug('Settings data set in RoamingSettings');
                
                // Save settings asynchronously
                console.debug('Initiating async save to Office...');
                roamingSettings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        console.debug('Office storage save succeeded');
                        resolve(true);
                    } else {
                        console.error('Failed to save to Office storage:', result.error);
                        resolve(false);
                    }
                });

            } catch (error) {
                console.error('Error saving to Office storage:', error);
                resolve(false);
            }
        });
    }

    /**
     * Loads settings from localStorage
     * @returns {Object|null} Settings object or null
     */
    loadFromLocalStorage() {
        console.debug('Loading from localStorage...');
        try {
            if (typeof localStorage === 'undefined') {
                console.warn('localStorage not available');
                return null;
            }

            const settingsJson = localStorage.getItem(this.storageKey);
            console.debug('Raw localStorage data:', settingsJson);
            if (settingsJson) {
                const settings = JSON.parse(settingsJson);
                console.debug('Parsed localStorage settings:', settings);
                return settings;
            }
            
            console.debug('No settings found in localStorage');
            return null;

        } catch (error) {
            console.error('Failed to load from localStorage:', error);
            return null;
        }
    }

    /**
     * Saves settings to localStorage
     * @param {Object} settings - Settings to save
     */
    saveToLocalStorage(settings) {
        console.debug('Saving to localStorage...');
        console.debug('Settings to save to localStorage:', settings);
        try {
            if (typeof localStorage === 'undefined') {
                console.warn('localStorage not available for save');
                return;
            }

            const settingsJson = JSON.stringify(settings);
            console.debug('Serialized localStorage JSON:', settingsJson);
            localStorage.setItem(this.storageKey, settingsJson);
            console.debug('localStorage save completed');

        } catch (error) {
            console.error('Failed to save to localStorage:', error);
        }
    }

    /**
     * Gets a specific setting value
     * @param {string} key - Setting key
     * @param {*} defaultValue - Default value if not found
     * @returns {*} Setting value
     */
    getSetting(key, defaultValue = null) {
        const value = this.settings[key] !== undefined ? this.settings[key] : defaultValue;
        console.log(`[SettingsManager] Getting setting '${key}':`, value);
        return value;
    }

    /**
     * Sets a specific setting value
     * @param {string} key - Setting key
     * @param {*} value - Setting value
     * @returns {Promise<boolean>} Success status
     */
    async setSetting(key, value) {
        console.log(`[SettingsManager] Setting '${key}' to:`, value);
        this.settings[key] = value;
        const result = await this.saveSettings();
        console.log(`[SettingsManager] Save result for '${key}':`, result);
        return result;
    }

    /**
     * Gets all current settings
     * @returns {Object} All settings
     */
    getSettings() {
        console.debug('Getting all settings (instance:', this.instanceId, '):', this.settings);
        return { ...this.settings };
    }

    /**
     * Resets settings to defaults
     * @returns {Promise<boolean>} Success status
     */
    async resetToDefaults() {
        this.settings = { ...this.defaultSettings };
        return await this.saveSettings();
    }

    /**
     * Exports settings as JSON
     * @returns {string} JSON string of settings
     */
    exportSettings() {
        // Create export object without sensitive data
        const exportData = { ...this.settings };
        
        // Remove sensitive fields
        delete exportData['api-key'];
        delete exportData['endpoint-url'];
        
        return JSON.stringify(exportData, null, 2);
    }

    /**
     * Imports settings from JSON
     * @param {string} jsonString - JSON string of settings
     * @returns {Promise<boolean>} Success status
     */
    async importSettings(jsonString) {
        try {
            const importedSettings = JSON.parse(jsonString);
            
            // Validate imported settings
            const validatedSettings = this.validateSettings(importedSettings);
            
            // Merge with current settings (preserve sensitive data)
            const mergedSettings = {
                ...this.settings,
                ...validatedSettings,
                // Keep current API key and endpoint
                'api-key': this.settings['api-key'],
                'endpoint-url': this.settings['endpoint-url']
            };

            return await this.saveSettings(mergedSettings);

        } catch (error) {
            console.error('Failed to import settings:', error);
            return false;
        }
    }

    /**
     * Validates settings object
     * @param {Object} settings - Settings to validate
     * @returns {Object} Validated settings
     */
    validateSettings(settings) {
        const validated = {};

        // Validate each setting against defaults
        Object.keys(this.defaultSettings).forEach(key => {
            if (settings.hasOwnProperty(key)) {
                const value = settings[key];
                const defaultValue = this.defaultSettings[key];
                
                // Type validation
                if (typeof value === typeof defaultValue) {
                    validated[key] = value;
                } else {
                    console.warn(`Invalid type for setting '${key}', using default`);
                    validated[key] = defaultValue;
                }
            } else {
                validated[key] = this.defaultSettings[key];
            }
        });

        // Validate specific setting ranges
        if (validated['response-length']) {
            const length = parseInt(validated['response-length']);
            validated['response-length'] = (length >= 1 && length <= 5) ? length.toString() : '3';
        }

        if (validated['response-tone']) {
            const tone = parseInt(validated['response-tone']);
            validated['response-tone'] = (tone >= 1 && tone <= 5) ? tone.toString() : '3';
        }

        if (validated['response-urgency']) {
            const urgency = parseInt(validated['response-urgency']);
            validated['response-urgency'] = (urgency >= 1 && urgency <= 5) ? urgency.toString() : '3';
        }

        return validated;
    }

    /**
     * Adds a change listener
     * @param {Function} listener - Callback function for setting changes
     */
    addChangeListener(listener) {
        if (typeof listener === 'function') {
            this.changeListeners.push(listener);
        }
    }

    /**
     * Removes a change listener
     * @param {Function} listener - Listener function to remove
     */
    removeChangeListener(listener) {
        const index = this.changeListeners.indexOf(listener);
        if (index > -1) {
            this.changeListeners.splice(index, 1);
        }
    }

    /**
     * Notifies all change listeners
     * @param {Object} newSettings - New settings object
     */
    notifyChangeListeners(newSettings) {
        this.changeListeners.forEach(listener => {
            try {
                listener(newSettings);
            } catch (error) {
                console.error('Settings change listener error:', error);
            }
        });
    }

    /**
     * Gets settings migration information
     * @returns {Object} Migration status
     */
    getMigrationStatus() {
        const currentVersion = this.settings['settings-version'] || '1.0.0';
        const latestVersion = this.defaultSettings['settings-version'];
        
        return {
            current: currentVersion,
            latest: latestVersion,
            needsMigration: currentVersion !== latestVersion,
            lastUpdated: this.settings['last-updated']
        };
    }

    /**
     * Migrates settings to latest version
     * @returns {Promise<boolean>} Success status
     */
    async migrateSettings() {
        const migration = this.getMigrationStatus();
        
        if (!migration.needsMigration) {
            return true;
        }

        try {
            // Perform migration logic here
            // For now, just update the version
            this.settings['settings-version'] = migration.latest;
            
            return await this.saveSettings();

        } catch (error) {
            console.error('Settings migration failed:', error);
            return false;
        }
    }

    /**
     * Clears all stored settings
     * @returns {Promise<boolean>} Success status
     */
    async clearAllSettings() {
        try {
            // Clear from Office storage
            if (typeof Office !== 'undefined' && Office.context?.roamingSettings) {
                const roamingSettings = Office.context.roamingSettings;
                roamingSettings.remove(this.storageKey);
                
                await new Promise((resolve) => {
                    roamingSettings.saveAsync((result) => {
                        resolve(result.status === Office.AsyncResultStatus.Succeeded);
                    });
                });
            }

            // Clear from localStorage
            if (typeof localStorage !== 'undefined') {
                localStorage.removeItem(this.storageKey);
            }

            // Reset to defaults
            this.settings = { ...this.defaultSettings };
            
            return true;

        } catch (error) {
            console.error('Failed to clear settings:', error);
            return false;
        }
    }

    /**
     * Get provider-specific configuration for a given service
     * @param {string} provider - The provider key (e.g., 'openai', 'ollama')
     * @returns {Object} Provider configuration with api-key and endpoint-url
     */
    getProviderConfig(provider) {
        if (!provider) return { 'api-key': '', 'endpoint-url': '' };
        
        // Try provider-specific config first
        if (this.settings['provider-configs'] && this.settings['provider-configs'][provider]) {
            return { ...this.settings['provider-configs'][provider] };
        }
        
        // Fall back to legacy single settings for backwards compatibility
        return {
            'api-key': this.settings['api-key'] || '',
            'endpoint-url': this.settings['endpoint-url'] || ''
        };
    }

    /**
     * Set provider-specific configuration for a given service
     * @param {string} provider - The provider key
     * @param {string} apiKey - The API key for this provider
     * @param {string} endpointUrl - The endpoint URL for this provider
     * @returns {Promise<boolean>} Success status
     */
    async setProviderConfig(provider, apiKey, endpointUrl) {
        if (!provider) return false;
        
        try {
            // Initialize provider-configs if it doesn't exist
            if (!this.settings['provider-configs']) {
                this.settings['provider-configs'] = { ...this.defaultSettings['provider-configs'] };
            }
            
            // Initialize this provider's config if it doesn't exist
            if (!this.settings['provider-configs'][provider]) {
                this.settings['provider-configs'][provider] = { 'api-key': '', 'endpoint-url': '' };
            }
            
            // Update the provider's configuration
            this.settings['provider-configs'][provider]['api-key'] = apiKey || '';
            this.settings['provider-configs'][provider]['endpoint-url'] = endpointUrl || '';
            
            // Save the updated settings
            return await this.saveSettings(this.settings);
            
        } catch (error) {
            console.error('Failed to set provider config:', error);
            return false;
        }
    }

    /**
     * Migrate legacy single API key/endpoint settings to provider-specific format
     * This ensures backwards compatibility when upgrading
     */
    migrateLegacySettings() {
        try {
            // Only migrate if we have legacy settings but no provider-configs
            if ((this.settings['api-key'] || this.settings['endpoint-url']) && 
                !this.settings['provider-configs']) {
                
                console.debug('Migrating legacy settings to provider-specific format');
                
                // Initialize provider configs
                this.settings['provider-configs'] = { ...this.defaultSettings['provider-configs'] };
                
                // Get the current model service or default to openai
                const currentService = this.settings['model-service'] || 'openai';
                
                // Migrate the legacy settings to the current provider
                if (this.settings['provider-configs'][currentService]) {
                    this.settings['provider-configs'][currentService]['api-key'] = this.settings['api-key'] || '';
                    this.settings['provider-configs'][currentService]['endpoint-url'] = this.settings['endpoint-url'] || '';
                }
                
                console.debug('Legacy settings migrated successfully');
                return true;
            }
            
            return false; // No migration needed
            
        } catch (error) {
            console.error('Failed to migrate legacy settings:', error);
            return false;
        }
    }
}

