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
            'api-key': '',
            'endpoint-url': '',
            
            // Response Preferences
            'response-length': '3',
            'response-tone': '3',
            'response-urgency': '3',
            'custom-instructions': '',
            
            // Accessibility Settings
            'high-contrast': false,
            'screen-reader-mode': false,
            
            // Privacy & Security
            'enable-logging': true,
            'enable-telemetry': true,
            
            // UI Preferences
            'last-tab': 'analysis',
            'show-advanced-options': false,
            
            // Version tracking
            'settings-version': '1.0.0',
            'last-updated': null
        };
        
        this.settings = { ...this.defaultSettings };
        this.changeListeners = [];
    }

    /**
     * Loads settings from storage
     * @returns {Promise<Object>} Loaded settings
     */
    async loadSettings() {
        try {
            // Try Office.js RoamingSettings first
            const officeSettings = await this.loadFromOfficeStorage();
            if (officeSettings) {
                this.settings = { ...this.defaultSettings, ...officeSettings };
                return this.settings;
            }

            // Fallback to localStorage
            const localSettings = this.loadFromLocalStorage();
            if (localSettings) {
                this.settings = { ...this.defaultSettings, ...localSettings };
                return this.settings;
            }

            // No stored settings found, use defaults
            this.settings = { ...this.defaultSettings };
            await this.saveSettings(this.settings);
            
            return this.settings;

        } catch (error) {
            console.error('Failed to load settings:', error);
            this.settings = { ...this.defaultSettings };
            return this.settings;
        }
    }

    /**
     * Saves settings to storage
     * @param {Object} newSettings - Settings to save
     * @returns {Promise<boolean>} Success status
     */
    async saveSettings(newSettings = null) {
        try {
            const settingsToSave = newSettings || this.settings;
            
            // Update timestamp
            settingsToSave['last-updated'] = new Date().toISOString();
            
            // Update internal settings
            this.settings = { ...settingsToSave };

            // Save to Office.js RoamingSettings
            const officeSaved = await this.saveToOfficeStorage(settingsToSave);
            
            // Also save to localStorage as backup
            this.saveToLocalStorage(settingsToSave);

            // Notify listeners
            this.notifyChangeListeners(settingsToSave);

            return officeSaved;

        } catch (error) {
            console.error('Failed to save settings:', error);
            return false;
        }
    }

    /**
     * Loads settings from Office.js RoamingSettings
     * @returns {Promise<Object|null>} Settings object or null
     */
    async loadFromOfficeStorage() {
        return new Promise((resolve) => {
            try {
                if (typeof Office === 'undefined' || !Office.context?.roamingSettings) {
                    resolve(null);
                    return;
                }

                const roamingSettings = Office.context.roamingSettings;
                const settingsJson = roamingSettings.get(this.storageKey);
                
                if (settingsJson) {
                    const settings = JSON.parse(settingsJson);
                    resolve(settings);
                } else {
                    resolve(null);
                }

            } catch (error) {
                console.warn('Failed to load from Office storage:', error);
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
        return new Promise((resolve) => {
            try {
                if (typeof Office === 'undefined' || !Office.context?.roamingSettings) {
                    resolve(false);
                    return;
                }

                const roamingSettings = Office.context.roamingSettings;
                const settingsJson = JSON.stringify(settings);
                
                roamingSettings.set(this.storageKey, settingsJson);
                
                // Save settings asynchronously
                roamingSettings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
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
        try {
            if (typeof localStorage === 'undefined') {
                return null;
            }

            const settingsJson = localStorage.getItem(this.storageKey);
            if (settingsJson) {
                return JSON.parse(settingsJson);
            }
            
            return null;

        } catch (error) {
            console.warn('Failed to load from localStorage:', error);
            return null;
        }
    }

    /**
     * Saves settings to localStorage
     * @param {Object} settings - Settings to save
     */
    saveToLocalStorage(settings) {
        try {
            if (typeof localStorage === 'undefined') {
                return;
            }

            const settingsJson = JSON.stringify(settings);
            localStorage.setItem(this.storageKey, settingsJson);

        } catch (error) {
            console.warn('Failed to save to localStorage:', error);
        }
    }

    /**
     * Gets a specific setting value
     * @param {string} key - Setting key
     * @param {*} defaultValue - Default value if not found
     * @returns {*} Setting value
     */
    getSetting(key, defaultValue = null) {
        return this.settings[key] !== undefined ? this.settings[key] : defaultValue;
    }

    /**
     * Sets a specific setting value
     * @param {string} key - Setting key
     * @param {*} value - Setting value
     * @returns {Promise<boolean>} Success status
     */
    async setSetting(key, value) {
        this.settings[key] = value;
        return await this.saveSettings();
    }

    /**
     * Gets all current settings
     * @returns {Object} All settings
     */
    getSettings() {
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
}
