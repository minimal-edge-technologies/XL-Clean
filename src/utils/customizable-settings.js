/*
 * Excel Data Cleaner Add-in
 * Customizable Settings
 */

/**
 * Default settings
 */
const DEFAULT_SETTINGS = {
    // General settings
    preview: {
      automaticPreviews: true,
      previewRows: 5,
      previewColumns: 3
    },
    enableUndo: true,
    showNotifications: true,
    
    // Feature-specific settings
    trimSpaces: {
      trimLeft: true,
      trimRight: true,
      reduceDuplicateSpaces: true
    },
    caseConversion: {
      defaultCase: "PROPER", // UPPER, LOWER, PROPER
      respectAcronyms: true,
      preserveCase: [] // List of words to preserve case for
    },
    dateFormat: {
      preferredFormat: "MM/DD/YYYY",
      detectExisting: true,
      preserveTimeComponents: true
    },
    findReplace: {
      useRegex: false,
      matchCase: false,
      wholeCell: false,
      recentSearches: []
    },
    oneClickCleanup: {
      enabledOperations: ["duplicates", "spaces", "case", "formatting"],
      autoDetectOperations: true
    }
  };
  
  /**
   * Settings manager for the add-in
   */
  export class SettingsManager {
    constructor() {
      this.settings = { ...DEFAULT_SETTINGS };
      this.loadSettings();
    }
    
    /**
     * Load settings from localStorage
     */
    loadSettings() {
      try {
        const savedSettings = localStorage.getItem("excelDataCleanerSettings");
        if (savedSettings) {
          const parsedSettings = JSON.parse(savedSettings);
          // Merge saved settings with defaults to ensure all properties exist
          this.settings = this._mergeDeep(this.settings, parsedSettings);
        }
      } catch (error) {
        console.error("Error loading settings:", error);
        // If loading fails, reset to defaults
        this.resetSettings();
      }
    }
    
    /**
     * Save settings to localStorage
     */
    saveSettings() {
      try {
        localStorage.setItem("excelDataCleanerSettings", JSON.stringify(this.settings));
      } catch (error) {
        console.error("Error saving settings:", error);
      }
    }
    
    /**
     * Get a specific setting by path
     * @param {string} path - Dot notation path to setting (e.g., "trimSpaces.trimLeft")
     * @param {any} defaultValue - Default value if setting doesn't exist
     * @returns {any} The setting value
     */
    getSetting(path, defaultValue) {
      const parts = path.split('.');
      let current = this.settings;
      
      for (const part of parts) {
        if (current === null || current === undefined || typeof current !== 'object') {
          return defaultValue;
        }
        current = current[part];
      }
      
      return current !== undefined ? current : defaultValue;
    }
    
    /**
     * Update a specific setting
     * @param {string} path - Dot notation path to setting
     * @param {any} value - New value for the setting
     */
    updateSetting(path, value) {
      const parts = path.split('.');
      let current = this.settings;
      
      for (let i = 0; i < parts.length - 1; i++) {
        const part = parts[i];
        if (!(part in current)) {
          current[part] = {};
        }
        current = current[part];
      }
      
      current[parts[parts.length - 1]] = value;
      this.saveSettings();
    }
    
    /**
     * Reset all settings to defaults
     */
    resetSettings() {
      this.settings = { ...DEFAULT_SETTINGS };
      this.saveSettings();
    }
    
    /**
     * Reset a specific category of settings
     * @param {string} category - The category to reset
     */
    resetCategory(category) {
      if (category in DEFAULT_SETTINGS) {
        this.settings[category] = { ...DEFAULT_SETTINGS[category] };
        this.saveSettings();
      }
    }
    
    /**
     * Add a recently used value for a specific feature
     * @param {string} feature - Feature name (e.g., "findReplace")
     * @param {string} listName - Name of the list (e.g., "recentSearches")
     * @param {any} value - Value to add
     * @param {number} maxItems - Maximum number of items to keep
     */
    addRecentItem(feature, listName, value, maxItems = 10) {
      if (!this.settings[feature]) {
        this.settings[feature] = {};
      }
      
      if (!this.settings[feature][listName]) {
        this.settings[feature][listName] = [];
      }
      
      // Remove if already exists (to move it to the top)
      const list = this.settings[feature][listName];
      const index = list.indexOf(value);
      if (index !== -1) {
        list.splice(index, 1);
      }
      
      // Add to the beginning
      list.unshift(value);
      
      // Limit to maxItems
      if (list.length > maxItems) {
        list.length = maxItems;
      }
      
      this.saveSettings();
    }
    
    /**
     * Helper method to deep merge objects
     * @private
     */
    _mergeDeep(target, source) {
      const output = { ...target };
      
      if (this._isObject(target) && this._isObject(source)) {
        Object.keys(source).forEach(key => {
          if (this._isObject(source[key])) {
            if (!(key in target)) {
              output[key] = source[key];
            } else {
              output[key] = this._mergeDeep(target[key], source[key]);
            }
          } else {
            output[key] = source[key];
          }
        });
      }
      
      return output;
    }
    
    /**
     * Helper method to check if value is an object
     * @private
     */
    _isObject(item) {
      return (item && typeof item === 'object' && !Array.isArray(item));
    }
  }
  
  // Create and export a singleton instance
  export const settingsManager = new SettingsManager();
  
  /**
 * Create and show the settings dialog
 */
export function showSettingsDialog() {
  // Create dialog if it doesn't exist
  let settingsDialog = document.getElementById("settings-dialog");
  
  if (!settingsDialog) {
    settingsDialog = document.createElement("div");
    settingsDialog.id = "settings-dialog";
    settingsDialog.className = "ms-Dialog";
    settingsDialog.setAttribute("role", "dialog");
    settingsDialog.setAttribute("aria-labelledby", "settings-dialog-title");
    
    document.body.appendChild(settingsDialog);
  }
  
  // Set up the dialog content
  settingsDialog.innerHTML = `
    <div class="ms-Dialog-main">
      <div class="ms-Dialog-title" id="settings-dialog-title">Settings</div>
      <div class="ms-Dialog-content">
        <div class="settings-tabs">
          <ul class="settings-tabs-list" role="tablist">
            <li role="presentation">
              <button id="tab-general" class="settings-tab-button selected" role="tab" aria-selected="true" aria-controls="panel-general">
                General
              </button>
            </li>
            <li role="presentation">
              <button id="tab-trim-spaces" class="settings-tab-button" role="tab" aria-selected="false" aria-controls="panel-trim-spaces">
                Trim Spaces
              </button>
            </li>
            <li role="presentation">
              <button id="tab-case" class="settings-tab-button" role="tab" aria-selected="false" aria-controls="panel-case">
                Text Case
              </button>
            </li>
            <li role="presentation">
              <button id="tab-dates" class="settings-tab-button" role="tab" aria-selected="false" aria-controls="panel-dates">
                Dates
              </button>
            </li>
            <li role="presentation">
              <button id="tab-find-replace" class="settings-tab-button" role="tab" aria-selected="false" aria-controls="panel-find-replace">
                Find & Replace
              </button>
            </li>
            <li role="presentation">
              <button id="tab-one-click" class="settings-tab-button" role="tab" aria-selected="false" aria-controls="panel-one-click">
                One-Click Cleanup
              </button>
            </li>
          </ul>
          
          <div class="settings-tab-panels">
            <!-- General Settings Panel -->
            <div id="panel-general" class="settings-tab-panel" role="tabpanel" aria-labelledby="tab-general">
              <div class="form-group checkbox-option">
                <input id="setting-show-preview" type="checkbox" 
                      ${settingsManager.getSetting("showPreviewBeforeApplying", true) ? "checked" : ""}>
                <label for="setting-show-preview">
                  Show preview before applying changes
                </label>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-enable-undo" type="checkbox"
                      ${settingsManager.getSetting("enableUndo", true) ? "checked" : ""}>
                <label for="setting-enable-undo">
                  Enable undo functionality
                </label>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-show-notifications" type="checkbox"
                      ${settingsManager.getSetting("showNotifications", true) ? "checked" : ""}>
                <label for="setting-show-notifications">
                  Show notifications
                </label>
              </div>
            </div>
            
            <!-- Trim Spaces Panel -->
            <div id="panel-trim-spaces" class="settings-tab-panel" role="tabpanel" aria-labelledby="tab-trim-spaces" hidden>
              <div class="form-group checkbox-option">
                <input id="setting-trim-left" type="checkbox"
                      ${settingsManager.getSetting("trimSpaces.trimLeft", true) ? "checked" : ""}>
                <label for="setting-trim-left">
                  Remove leading spaces
                </label>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-trim-right" type="checkbox"
                      ${settingsManager.getSetting("trimSpaces.trimRight", true) ? "checked" : ""}>
                <label for="setting-trim-right">
                  Remove trailing spaces
                </label>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-reduce-spaces" type="checkbox"
                      ${settingsManager.getSetting("trimSpaces.reduceDuplicateSpaces", true) ? "checked" : ""}>
                <label for="setting-reduce-spaces">
                  Reduce multiple spaces to a single space
                </label>
              </div>
            </div>
            
            <!-- Text Case Panel -->
            <div id="panel-case" class="settings-tab-panel" role="tabpanel" aria-labelledby="tab-case" hidden>
              <div class="form-group">
                <label class="group-label">Default Case Conversion:</label>
                <div class="radio-option">
                  <input id="setting-case-upper" type="radio" name="defaultCase"
                        value="UPPER" ${settingsManager.getSetting("caseConversion.defaultCase", "PROPER") === "UPPER" ? "checked" : ""}>
                  <label for="setting-case-upper">UPPERCASE</label>
                </div>
                
                <div class="radio-option">
                  <input id="setting-case-lower" type="radio" name="defaultCase"
                        value="LOWER" ${settingsManager.getSetting("caseConversion.defaultCase", "PROPER") === "LOWER" ? "checked" : ""}>
                  <label for="setting-case-lower">lowercase</label>
                </div>
                
                <div class="radio-option">
                  <input id="setting-case-proper" type="radio" name="defaultCase"
                        value="PROPER" ${settingsManager.getSetting("caseConversion.defaultCase", "PROPER") === "PROPER" ? "checked" : ""}>
                  <label for="setting-case-proper">Proper Case</label>
                </div>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-respect-acronyms" type="checkbox"
                      ${settingsManager.getSetting("caseConversion.respectAcronyms", true) ? "checked" : ""}>
                <label for="setting-respect-acronyms">
                  Respect common acronyms (e.g., NASA, IBM)
                </label>
              </div>
              
              <div class="form-group">
                <label for="setting-preserve-case">Preserve case for specific words (comma-separated):</label>
                <input id="setting-preserve-case" type="text" 
                      value="${settingsManager.getSetting("caseConversion.preserveCase", []).join(", ")}">
              </div>
            </div>
            
            <!-- Dates Panel -->
            <div id="panel-dates" class="settings-tab-panel" role="tabpanel" aria-labelledby="tab-dates" hidden>
              <div class="form-group">
                <label class="group-label">Preferred Date Format:</label>
                <div class="radio-option">
                  <input id="setting-date-us" type="radio" name="dateFormat"
                        value="MM/DD/YYYY" ${settingsManager.getSetting("dateFormat.preferredFormat", "MM/DD/YYYY") === "MM/DD/YYYY" ? "checked" : ""}>
                  <label for="setting-date-us">MM/DD/YYYY (US Format)</label>
                </div>
                
                <div class="radio-option">
                  <input id="setting-date-eu" type="radio" name="dateFormat"
                        value="DD/MM/YYYY" ${settingsManager.getSetting("dateFormat.preferredFormat", "MM/DD/YYYY") === "DD/MM/YYYY" ? "checked" : ""}>
                  <label for="setting-date-eu">DD/MM/YYYY (European Format)</label>
                </div>
                
                <div class="radio-option">
                  <input id="setting-date-iso" type="radio" name="dateFormat"
                        value="YYYY-MM-DD" ${settingsManager.getSetting("dateFormat.preferredFormat", "MM/DD/YYYY") === "YYYY-MM-DD" ? "checked" : ""}>
                  <label for="setting-date-iso">YYYY-MM-DD (ISO Format)</label>
                </div>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-detect-date-format" type="checkbox"
                      ${settingsManager.getSetting("dateFormat.detectExisting", true) ? "checked" : ""}>
                <label for="setting-detect-date-format">
                  Try to detect existing date format
                </label>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-preserve-time" type="checkbox"
                      ${settingsManager.getSetting("dateFormat.preserveTimeComponents", true) ? "checked" : ""}>
                <label for="setting-preserve-time">
                  Preserve time components when standardizing dates
                </label>
              </div>
            </div>
            
            <!-- Find & Replace Panel -->
            <div id="panel-find-replace" class="settings-tab-panel" role="tabpanel" aria-labelledby="tab-find-replace" hidden>
              <div class="form-group checkbox-option">
                <input id="setting-use-regex" type="checkbox"
                      ${settingsManager.getSetting("findReplace.useRegex", false) ? "checked" : ""}>
                <label for="setting-use-regex">
                  Enable regular expressions by default
                </label>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-match-case" type="checkbox"
                      ${settingsManager.getSetting("findReplace.matchCase", false) ? "checked" : ""}>
                <label for="setting-match-case">
                  Match case by default
                </label>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-whole-cell" type="checkbox"
                      ${settingsManager.getSetting("findReplace.wholeCell", false) ? "checked" : ""}>
                <label for="setting-whole-cell">
                  Match entire cell contents by default
                </label>
              </div>
              
              <div class="recent-searches">
                <label class="group-label">Recent searches:</label>
                <div id="recent-searches-list" class="recent-list">
                  ${generateRecentSearchesList()}
                </div>
                <button id="clear-recent-searches" class="secondary-button">
                  <span>Clear Recent Searches</span>
                </button>
              </div>
            </div>
            
            <!-- One-Click Cleanup Panel -->
            <div id="panel-one-click" class="settings-tab-panel" role="tabpanel" aria-labelledby="tab-one-click" hidden>
              <label class="group-label">Default Operations to Include:</label>
              
              <div class="form-group checkbox-option">
                <input id="setting-cleanup-duplicates" type="checkbox"
                      ${settingsManager.getSetting("oneClickCleanup.enabledOperations", []).includes("duplicates") ? "checked" : ""}>
                <label for="setting-cleanup-duplicates">
                  <i class="fas fa-clone"></i> Remove duplicates
                </label>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-cleanup-spaces" type="checkbox"
                      ${settingsManager.getSetting("oneClickCleanup.enabledOperations", []).includes("spaces") ? "checked" : ""}>
                <label for="setting-cleanup-spaces">
                  <i class="fas fa-compress-alt"></i> Trim spaces
                </label>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-cleanup-case" type="checkbox"
                      ${settingsManager.getSetting("oneClickCleanup.enabledOperations", []).includes("case") ? "checked" : ""}>
                <label for="setting-cleanup-case">
                  <i class="fas fa-font"></i> Fix text case
                </label>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-cleanup-formatting" type="checkbox"
                      ${settingsManager.getSetting("oneClickCleanup.enabledOperations", []).includes("formatting") ? "checked" : ""}>
                <label for="setting-cleanup-formatting">
                  <i class="fas fa-percentage"></i> Fix number formatting
                </label>
              </div>
              
              <div class="form-group checkbox-option">
                <input id="setting-auto-detect" type="checkbox"
                      ${settingsManager.getSetting("oneClickCleanup.autoDetectOperations", true) ? "checked" : ""}>
                <label for="setting-auto-detect">
                  Automatically detect necessary operations
                </label>
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="ms-Dialog-actions">
        <button id="save-settings" class="primary-button">
          <span>Save</span>
        </button>
        <button id="reset-settings" class="secondary-button">
          <span>Reset to Defaults</span>
        </button>
        <button id="close-settings" class="secondary-button">
          <span>Cancel</span>
        </button>
      </div>
    </div>
  `;
  
  // Add styles for the settings dialog
  if (!document.getElementById("settings-dialog-styles")) {
    const styleEl = document.createElement("style");
    styleEl.id = "settings-dialog-styles";
    styleEl.textContent = `
      .settings-tabs {
        display: flex;
        flex-direction: column;
        height: 100%;
      }
      
      .settings-tabs-list {
        display: flex;
        list-style: none;
        padding: 0;
        margin: 0 0 20px 0;
        border-bottom: 1px solid var(--border);
        overflow-x: auto;
      }
      
      .settings-tab-button {
        padding: 12px 15px;
        border: none;
        background: none;
        cursor: pointer;
        position: relative;
        font-size: 14px;
        color: var(--text-secondary);
        font-family: inherit;
      }
      
      .settings-tab-button.selected {
        font-weight: 600;
        color: var(--primary);
      }
      
      .settings-tab-button.selected::after {
        content: '';
        position: absolute;
        bottom: -1px;
        left: 0;
        right: 0;
        height: 2px;
        background-color: var(--primary);
      }
      
      .settings-tab-panel {
        padding: 15px 10px;
      }
      
      .recent-list {
        margin: 10px 0;
        max-height: 150px;
        overflow-y: auto;
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 5px;
      }
      
      .recent-item {
        padding: 8px;
        display: flex;
        justify-content: space-between;
        align-items: center;
      }
      
      .recent-item:not(:last-child) {
        border-bottom: 1px solid var(--border);
      }
      
      .remove-recent {
        background: none;
        border: none;
        color: var(--error);
        cursor: pointer;
        padding: 4px 8px;
        font-size: 12px;
        border-radius: var(--radius);
      }
      
      .remove-recent:hover {
        background-color: rgba(239, 68, 68, 0.1);
      }
    `;
    document.head.appendChild(styleEl);
  }
  
  // Show the dialog
  settingsDialog.style.display = "flex";
  
  // Add event listeners for tab switching
  const tabButtons = document.querySelectorAll(".settings-tab-button");
  const tabPanels = document.querySelectorAll(".settings-tab-panel");
  
  tabButtons.forEach(button => {
    button.addEventListener("click", () => {
      // Update selected state for buttons
      tabButtons.forEach(btn => {
        btn.classList.remove("selected");
        btn.setAttribute("aria-selected", "false");
      });
      button.classList.add("selected");
      button.setAttribute("aria-selected", "true");
      
      // Hide all panels and show the selected one
      tabPanels.forEach(panel => {
        panel.hidden = true;
      });
      const panelId = button.getAttribute("aria-controls");
      document.getElementById(panelId).hidden = false;
    });
  });
  
  // Add event listeners for buttons
  document.getElementById("save-settings").addEventListener("click", () => {
    saveSettingsFromDialog();
    settingsDialog.style.display = "none";
  });
  
  document.getElementById("reset-settings").addEventListener("click", () => {
    if (confirm("Are you sure you want to reset all settings to their defaults?")) {
      settingsManager.resetSettings();
      settingsDialog.style.display = "none";
      showSettingsDialog(); // Reopen with default values
    }
  });
  
  document.getElementById("close-settings").addEventListener("click", () => {
    settingsDialog.style.display = "none";
  });
  
  // Add event listener for clear recent searches
  const clearRecentButton = document.getElementById("clear-recent-searches");
  if (clearRecentButton) {
    clearRecentButton.addEventListener("click", () => {
      settingsManager.updateSetting("findReplace.recentSearches", []);
      document.getElementById("recent-searches-list").innerHTML = 
        "<div class='recent-item'>No recent searches</div>";
    });
  }
  
  // Add event listeners for remove recent search items
  document.querySelectorAll(".remove-recent").forEach(button => {
    button.addEventListener("click", (e) => {
      const searchText = e.target.getAttribute("data-search");
      const recentSearches = settingsManager.getSetting("findReplace.recentSearches", []);
      const updatedSearches = recentSearches.filter(item => item !== searchText);
      settingsManager.updateSetting("findReplace.recentSearches", updatedSearches);
      e.target.closest(".recent-item").remove();
      
      if (updatedSearches.length === 0) {
        document.getElementById("recent-searches-list").innerHTML = 
          "<div class='recent-item'>No recent searches</div>";
      }
    });
  });
}
  
  /**
   * Save settings from the dialog form
   */
  function saveSettingsFromDialog() {
    // General settings
    settingsManager.updateSetting(
      "showPreviewBeforeApplying", 
      document.getElementById("setting-show-preview").checked
    );
    
    settingsManager.updateSetting(
      "enableUndo", 
      document.getElementById("setting-enable-undo").checked
    );
    
    settingsManager.updateSetting(
      "showNotifications", 
      document.getElementById("setting-show-notifications").checked
    );
    
    // Trim spaces settings
    settingsManager.updateSetting(
      "trimSpaces.trimLeft", 
      document.getElementById("setting-trim-left").checked
    );
    
    settingsManager.updateSetting(
      "trimSpaces.trimRight", 
      document.getElementById("setting-trim-right").checked
    );
    
    settingsManager.updateSetting(
      "trimSpaces.reduceDuplicateSpaces", 
      document.getElementById("setting-reduce-spaces").checked
    );
    
    // Case conversion settings
    const defaultCaseRadios = document.getElementsByName("defaultCase");
    for (const radio of defaultCaseRadios) {
      if (radio.checked) {
        settingsManager.updateSetting("caseConversion.defaultCase", radio.value);
        break;
      }
    }
    
    settingsManager.updateSetting(
      "caseConversion.respectAcronyms", 
      document.getElementById("setting-respect-acronyms").checked
    );
    
    const preserveCaseInput = document.getElementById("setting-preserve-case").value;
    const preserveCaseWords = preserveCaseInput
      .split(",")
      .map(word => word.trim())
      .filter(word => word.length > 0);
    
    settingsManager.updateSetting("caseConversion.preserveCase", preserveCaseWords);
    
    // Date format settings
    const dateFormatRadios = document.getElementsByName("dateFormat");
    for (const radio of dateFormatRadios) {
      if (radio.checked) {
        settingsManager.updateSetting("dateFormat.preferredFormat", radio.value);
        break;
      }
    }
    
    settingsManager.updateSetting(
      "dateFormat.detectExisting", 
      document.getElementById("setting-detect-date-format").checked
    );
    
    settingsManager.updateSetting(
      "dateFormat.preserveTimeComponents", 
      document.getElementById("setting-preserve-time").checked
    );
    
    // Find & Replace settings
    settingsManager.updateSetting(
      "findReplace.useRegex", 
      document.getElementById("setting-use-regex").checked
    );
    
    settingsManager.updateSetting(
      "findReplace.matchCase", 
      document.getElementById("setting-match-case").checked
    );
    
    settingsManager.updateSetting(
      "findReplace.wholeCell", 
      document.getElementById("setting-whole-cell").checked
    );
    
    // One-Click Cleanup settings
  const enabledOperations = [];
  if (document.getElementById("setting-cleanup-duplicates").checked) enabledOperations.push("duplicates");
  if (document.getElementById("setting-cleanup-spaces").checked) enabledOperations.push("spaces");
  if (document.getElementById("setting-cleanup-case").checked) enabledOperations.push("case");
  if (document.getElementById("setting-cleanup-formatting").checked) enabledOperations.push("formatting");
  
  settingsManager.updateSetting("oneClickCleanup.enabledOperations", enabledOperations);
  
  settingsManager.updateSetting(
    "oneClickCleanup.autoDetectOperations", 
    document.getElementById("setting-auto-detect").checked
  );
}

/**
 * Generate HTML for the recent searches list
 * @returns {string} HTML string for the recent searches list
 */
function generateRecentSearchesList() {
  const recentSearches = settingsManager.getSetting("findReplace.recentSearches", []);
  
  if (recentSearches.length === 0) {
    return "<div class='recent-item'>No recent searches</div>";
  }
  
  return recentSearches.map(search => {
    // Escape HTML for safety
    const escapedSearch = search
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
    
    return `
      <div class='recent-item'>
        <span>${escapedSearch}</span>
        <button class='remove-recent' data-search="${escapedSearch}">
          <i class="fas fa-times"></i>
        </button>
      </div>
    `;
  }).join("");
}

/**
 * Add a settings button to the UI
 */
export function addSettingsButton() {
  // Check if settings button already exists
  if (document.getElementById("settings-button")) return;
  
  // Create the settings button
  const settingsButton = document.createElement("button");
  settingsButton.id = "settings-button";
  settingsButton.className = "ms-Button ms-Button--icon";
  settingsButton.title = "Settings";
  settingsButton.innerHTML = `
    <span class="ms-Button-icon"><i class="fas fa-cog" aria-hidden="true"></i></span>
    <span class="ms-Button-label">Settings</span>
  `;
  
  // Add event listener
  settingsButton.addEventListener("click", showSettingsDialog);
  
  // Add to DOM - place it in the header
  const header = document.querySelector(".ms-welcome__header");
  if (header) {
    // Create a container for the button if it doesn't exist
    let buttonContainer = header.querySelector(".header-button-container");
    if (!buttonContainer) {
      buttonContainer = document.createElement("div");
      buttonContainer.className = "header-button-container";
      header.appendChild(buttonContainer);
    }
    
    // Add the button to the container
    buttonContainer.appendChild(settingsButton);
    
    // Make sure styles are added
    if (!document.getElementById("header-button-styles")) {
      const styleEl = document.createElement("style");
      styleEl.id = "header-button-styles";
      styleEl.textContent = `
        .ms-welcome__header {
          display: flex;
          justify-content: space-between;
          align-items: center;
        }
        .header-button-container {
          display: flex;
          gap: 8px;
        }
        .ms-Button--icon {
          background: transparent;
          border: none;
          color: white;
          padding: 6px 10px;
          cursor: pointer;
          border-radius: 4px;
          display: flex;
          align-items: center;
          gap: 6px;
          transition: background-color 0.2s;
        }
        .ms-Button--icon:hover {
          background-color: rgba(255, 255, 255, 0.1);
        }
        .ms-Button--icon .ms-Button-icon {
          font-size: 14px;
        }
        .ms-Button--icon .ms-Button-label {
          font-size: 12px;
          font-weight: normal;
        }
      `;
      document.head.appendChild(styleEl);
    }
  }
}

/**
 * Apply custom trim function based on settings
 * @param {string} text - Text to trim
 * @returns {string} Trimmed text according to settings
 */
export function applyCustomTrim(text) {
  if (typeof text !== 'string') return text;
  
  let result = text;
  
  // Apply trim left if enabled
  if (settingsManager.getSetting("trimSpaces.trimLeft", true)) {
    result = result.replace(/^\s+/, '');
  }
  
  // Apply trim right if enabled
  if (settingsManager.getSetting("trimSpaces.trimRight", true)) {
    result = result.replace(/\s+$/, '');
  }
  
  // Apply multiple space reduction if enabled
  if (settingsManager.getSetting("trimSpaces.reduceDuplicateSpaces", true)) {
    result = result.replace(/\s+/g, ' ');
  }
  
  return result;
}

/**
 * Apply custom case conversion based on settings
 * @param {string} text - Text to convert
 * @param {string} caseType - Type of case conversion (overrides default if provided)
 * @returns {string} Case-converted text according to settings
 */
export function applyCustomCaseConversion(text, caseType = null) {
  if (typeof text !== 'string') return text;
  
  // Use provided case type or default from settings
  const conversionType = caseType || settingsManager.getSetting("caseConversion.defaultCase", "PROPER");
  
  // Get words that should preserve their case
  const preserveCaseWords = settingsManager.getSetting("caseConversion.preserveCase", []);
  const respectAcronyms = settingsManager.getSetting("caseConversion.respectAcronyms", true);
  
  // Common acronyms (only used if respectAcronyms is true)
  const commonAcronyms = [
    "NASA", "FBI", "CIA", "USA", "UK", "UN", "NATO", "CEO", "CFO", "CTO",
    "HR", "IT", "API", "URL", "HTTP", "HTTPS", "FTP", "HTML", "CSS", "XML",
    "JSON", "SQL", "PDF", "iOS", "ID", "PhD", "MBA", "BA", "BS", "MD"
  ];
  
  // Create a combined list of words to preserve case
  const wordsToPreserve = [...preserveCaseWords];
  if (respectAcronyms) {
    wordsToPreserve.push(...commonAcronyms);
  }
  
  // Create a dictionary for quick lookups
  const preserveCaseDict = {};
  wordsToPreserve.forEach(word => {
    preserveCaseDict[word.toLowerCase()] = word;
  });
  
  switch (conversionType) {
    case "UPPER":
      return text.toUpperCase();
      
    case "LOWER":
      return text.toLowerCase();
      
    case "PROPER":
      // Convert to proper case while preserving specific words
      return text.toLowerCase().replace(/(^|\s|\(|\[|\{|"|')(\S)/g, (match, p1, p2) => {
        return p1 + p2.toUpperCase();
      }).replace(/\b([a-z]+)\b/gi, (word) => {
        const lowerWord = word.toLowerCase();
        return preserveCaseDict[lowerWord] || word;
      });
      
    default:
      return text;
  }
}

/**
 * Get enabled operations for one-click cleanup
 * @returns {string[]} Array of enabled operation keys
 */
export function getEnabledOperations() {
  return settingsManager.getSetting("oneClickCleanup.enabledOperations", 
    ["duplicates", "spaces", "case", "formatting"]);
}

/**
 * Initialize settings with default values
 */
export function initializeSettings() {
  settingsManager.loadSettings();
  addSettingsButton();
}