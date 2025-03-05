/*
 * Excel Data Cleaner Add-in
 * Help Documentation System
 */

/**
 * Help content for each feature
 */
const HELP_CONTENT = {
    general: {
      title: "Excel Data Cleaner Help",
      content: `
        <h2>Welcome to Excel Data Cleaner</h2>
        <p>Excel Data Cleaner is a powerful tool designed to help you clean and standardize your Excel data with just a few clicks.</p>
        
        <h3>Getting Started</h3>
        <ol>
          <li>Select the range of cells you want to clean</li>
          <li>Choose the operation you want to perform from the tabs</li>
          <li>Configure any settings specific to that operation</li>
          <li>Click the action button to clean your data</li>
        </ol>
        
        <h3>Features</h3>
        <ul>
          <li><strong>Remove Duplicates</strong> - Find and remove duplicate rows in your data</li>
          <li><strong>Trim Spaces</strong> - Remove unwanted spaces from the beginning, end, or middle of text</li>
          <li><strong>Text Case</strong> - Convert text to UPPERCASE, lowercase, or Proper Case</li>
          <li><strong>Find & Replace</strong> - Search for specific text and replace it across multiple cells</li>
          <li><strong>Format Dates</strong> - Standardize dates to a consistent format</li>
          <li><strong>One-Click Cleanup</strong> - Apply multiple cleaning operations at once</li>
        </ul>
        
        <h3>Tips</h3>
        <ul>
          <li>You can preview changes before applying them</li>
          <li>Use the settings (gear icon) to customize how each feature works</li>
          <li>The undo button allows you to revert your most recent operation</li>
          <li>You can save your preferred settings for repeated use</li>
        </ul>
      `
    },
    duplicates: {
      title: "Remove Duplicates Help",
      content: `
        <h2>Remove Duplicates</h2>
        <p>This feature helps you identify and remove duplicate rows in your data, keeping only the first occurrence of each unique row.</p>
        
        <h3>How to Use</h3>
        <ol>
          <li>Select the range of cells that contains duplicates</li>
          <li>Click the "Remove Duplicates" button</li>
          <li>The add-in will automatically identify and remove duplicate rows</li>
          <li>You'll see a confirmation message showing how many duplicates were removed</li>
        </ol>
        
        <h3>Notes</h3>
        <ul>
          <li>The entire row will be considered when identifying duplicates</li>
          <li>The comparison is case-sensitive ("ABC" is different from "abc")</li>
          <li>Only the first occurrence of each row will be kept; all others will be removed</li>
          <li>The function rearranges your worksheet by removing rows; your data will shift upward</li>
        </ul>
        
        <h3>Tips</h3>
        <ul>
          <li>If you only want to check for duplicates in specific columns, select just those columns</li>
          <li>Make sure your selection includes all relevant data in the rows</li>
          <li>Consider using the preview feature to see which rows will be removed before proceeding</li>
        </ul>
      `
    },
    trimSpaces: {
      title: "Trim Spaces Help",
      content: `
        <h2>Trim Spaces</h2>
        <p>This feature removes extra spaces from the beginning, end, or middle of text in your cells.</p>
        
        <h3>How to Use</h3>
        <ol>
          <li>Select the range of cells containing text with unwanted spaces</li>
          <li>Click the "Trim Extra Spaces" button</li>
          <li>The add-in will automatically clean up spaces in your text</li>
          <li>You'll see a confirmation message showing how many cells were modified</li>
        </ol>
        
        <h3>What It Does</h3>
        <ul>
          <li><strong>Leading spaces</strong> at the beginning of text are removed</li>
          <li><strong>Trailing spaces</strong> at the end of text are removed</li>
          <li><strong>Multiple spaces</strong> between words are reduced to single spaces</li>
        </ul>
        
        <h3>Customization</h3>
        <p>You can customize which types of space trimming are applied in the Settings panel:</p>
        <ul>
          <li>Enable/disable removing leading spaces</li>
          <li>Enable/disable removing trailing spaces</li>
          <li>Enable/disable reducing multiple spaces to single spaces</li>
        </ul>
        
        <h3>Examples</h3>
        <table>
          <tr>
            <th>Before</th>
            <th>After</th>
          </tr>
          <tr>
            <td>"&nbsp;&nbsp;&nbsp;Hello&nbsp;&nbsp;"</td>
            <td>"Hello"</td>
          </tr>
          <tr>
            <td>"Hello&nbsp;&nbsp;&nbsp;&nbsp;World"</td>
            <td>"Hello World"</td>
          </tr>
          <tr>
            <td>"Multiple&nbsp;&nbsp;&nbsp;spaces&nbsp;&nbsp;here"</td>
            <td>"Multiple spaces here"</td>
          </tr>
        </table>
      `
    },
    case: {
      title: "Text Case Conversion Help",
      content: `
        <h2>Text Case Conversion</h2>
        <p>This feature allows you to change the capitalization of text in your cells to ensure consistent formatting.</p>
        
        <h3>How to Use</h3>
        <ol>
          <li>Select the range of cells containing text that needs case conversion</li>
          <li>Choose one of the three options:
            <ul>
              <li><strong>UPPERCASE</strong> - Converts all letters to capital letters</li>
              <li><strong>lowercase</strong> - Converts all letters to small letters</li>
              <li><strong>Proper Case</strong> - Capitalizes the first letter of each word</li>
            </ul>
          </li>
          <li>Click the corresponding button</li>
          <li>The add-in will convert the case of all text in the selected cells</li>
        </ol>
        
        <h3>Proper Case Details</h3>
        <p>Proper Case (also known as Title Case) applies these rules:</p>
        <ul>
          <li>The first letter of each word is capitalized</li>
          <li>All other letters are converted to lowercase</li>
          <li>Common acronyms can be preserved in their original case (configurable in settings)</li>
          <li>You can specify words to always preserve in their original case (in settings)</li>
        </ul>
        
        <h3>Examples</h3>
        <table>
          <tr>
            <th>Original</th>
            <th>UPPERCASE</th>
            <th>lowercase</th>
            <th>Proper Case</th>
          </tr>
          <tr>
            <td>john smith</td>
            <td>JOHN SMITH</td>
            <td>john smith</td>
            <td>John Smith</td>
          </tr>
          <tr>
            <td>SALES REPORT</td>
            <td>SALES REPORT</td>
            <td>sales report</td>
            <td>Sales Report</td>
          </tr>
          <tr>
            <td>mixed CASE text</td>
            <td>MIXED CASE TEXT</td>
            <td>mixed case text</td>
            <td>Mixed Case Text</td>
          </tr>
        </table>
      `
    },
    replace: {
      title: "Find & Replace Help",
      content: `
        <h2>Find & Replace</h2>
        <p>This feature allows you to search for specific text and replace it with new text across multiple cells at once.</p>
        
        <h3>How to Use</h3>
        <ol>
          <li>Select the range of cells where you want to perform find and replace</li>
          <li>Enter the text you want to find in the "Find what" field</li>
          <li>Enter the replacement text in the "Replace with" field</li>
          <li>Click the "Replace All" button</li>
          <li>The add-in will replace all occurrences and show you how many replacements were made</li>
        </ol>
        
        <h3>Advanced Options</h3>
        <p>You can customize the find & replace behavior in the Settings panel:</p>
        <ul>
          <li><strong>Regular Expressions</strong> - Use pattern matching for more complex searches</li>
          <li><strong>Match Case</strong> - Make the search case-sensitive</li>
          <li><strong>Match Entire Cell Contents</strong> - Only replace when the cell contains exactly the search text</li>
        </ul>
        
        <h3>Using Regular Expressions</h3>
        <p>If enabled in settings, you can use regular expressions for powerful pattern matching:</p>
        <ul>
          <li><code>\\d+</code> - Matches one or more digits</li>
          <li><code>^text</code> - Matches "text" at the beginning of a cell</li>
          <li><code>text$</code> - Matches "text" at the end of a cell</li>
          <li><code>[A-Z]+</code> - Matches one or more uppercase letters</li>
        </ul>
        
        <h3>Tips</h3>
        <ul>
          <li>Use the preview feature to see changes before applying them</li>
          <li>For complex replacements, consider making multiple passes with different find/replace operations</li>
          <li>Your recent searches are saved for easy access</li>
        </ul>
      `
    },
    dates: {
      title: "Format Dates Help",
      content: `
        <h2>Format Dates</h2>
        <p>This feature helps standardize dates in your spreadsheet to a consistent format for better sorting and filtering.</p>
        
        <h3>How to Use</h3>
        <ol>
          <li>Select the range of cells containing dates</li>
          <li>Choose your preferred date format:
            <ul>
              <li><strong>MM/DD/YYYY</strong> - US format (e.g., 12/31/2023)</li>
              <li><strong>DD/MM/YYYY</strong> - European format (e.g., 31/12/2023)</li>
              <li><strong>YYYY-MM-DD</strong> - ISO format (e.g., 2023-12-31)</li>
            </ul>
          </li>
          <li>Click the "Standardize Dates" button</li>
          <li>The add-in will convert all dates to your chosen format</li>
        </ol>
        
        <h3>Advanced Features</h3>
        <p>This tool can detect and convert dates in various formats:</p>
        <ul>
          <li>Text-based dates in different formats (MM/DD/YY, DD-MM-YYYY, etc.)</li>
          <li>Excel serial numbers (e.g., 44927 = December 31, 2022)</li>
          <li>Dates with month names (e.g., "31-Dec-2022", "December 31, 2022")</li>
          <li>International date formats</li>
        </ul>
        
        <h3>Settings Options</h3>
        <ul>
          <li><strong>Auto-detect existing format</strong> - Intelligently determines the current date format</li>
          <li><strong>Preserve time components</strong> - Keeps time information when standardizing dates</li>
        </ul>
        
        <h3>Notes</h3>
        <ul>
          <li>The standardization applies Excel date formatting, so the cells will contain actual date values</li>
          <li>This makes it possible to perform date calculations and sorting correctly</li>
          <li>Ambiguous dates (like 01/02/2023 which could be Jan 2 or Feb 1) will be interpreted based on the detected format</li>
        </ul>
      `
    },
    oneclick: {
      title: "One-Click Cleanup Help",
      content: `
        <h2>One-Click Cleanup</h2>
        <p>This feature allows you to apply multiple cleaning operations at once to save time and effort.</p>
        
        <h3>How to Use</h3>
        <ol>
          <li>Select the range of cells you want to clean</li>
          <li>Check the operations you want to perform:
            <ul>
              <li><strong>Remove duplicates</strong> - Eliminates duplicate rows</li>
              <li><strong>Trim spaces</strong> - Removes extra spaces from text</li>
              <li><strong>Fix text case</strong> - Applies proper case to text</li>
              <li><strong>Fix number formatting</strong> - Converts text numbers to actual number values</li>
            </ul>
          </li>
          <li>Click the "Clean Everything" button</li>
          <li>The add-in will apply all selected operations and show a summary of changes</li>
        </ol>
        
        <h3>How It Works</h3>
        <p>One-Click Cleanup performs operations in this order:</p>
        <ol>
          <li><strong>Trim spaces</strong> - First, all extra spaces are removed</li>
          <li><strong>Fix text case</strong> - Then, text case is standardized</li>
          <li><strong>Fix number formatting</strong> - Numbers stored as text are converted to actual numbers</li>
          <li><strong>Remove duplicates</strong> - Finally, any duplicate rows are removed</li>
        </ol>
        
        <h3>Smart Detection</h3>
        <p>If "Automatically detect necessary operations" is enabled in settings, the add-in will:</p>
        <ul>
          <li>Analyze your data to identify what needs cleaning</li>
          <li>Only apply operations that are needed</li>
          <li>Skip operations that wouldn't make any changes</li>
        </ul>
        
        <h3>Tips</h3>
        <ul>
          <li>Use this feature for quick data preparation before analysis</li>
          <li>You can save your preferred operation selections in settings</li>
          <li>Consider using the preview feature to see changes before applying them</li>
        </ul>
      `
    }
  };
  
  /**
   * Show help dialog for a specific feature
   * @param {string} featureId - The feature ID to show help for
   */
  export function showHelpDialog(featureId = "general") {
    // Get the help content (use general if the feature ID doesn't exist)
    const helpData = HELP_CONTENT[featureId] || HELP_CONTENT.general;
    
    // Create dialog if it doesn't exist
    let helpDialog = document.getElementById("help-dialog");
    if (!helpDialog) {
      helpDialog = document.createElement("div");
      helpDialog.id = "help-dialog";
      helpDialog.className = "ms-Dialog";
      helpDialog.setAttribute("role", "dialog");
      helpDialog.setAttribute("aria-labelledby", "help-dialog-title");
      
      document.body.appendChild(helpDialog);
    }
    
    // Set up the dialog content with proper structure
    helpDialog.innerHTML = `
      <div class="ms-Dialog-main">
        <div class="ms-Dialog-title" id="help-dialog-title">${helpData.title}</div>
        <div class="ms-Dialog-content help-content">
          ${helpData.content}
        </div>
        <div class="ms-Dialog-actions">
          <button id="close-help" class="primary-button">
            <span>Close</span>
          </button>
        </div>
      </div>
    `;
    
    // Add styles for the help dialog
    if (!document.getElementById("help-dialog-styles")) {
      const styleEl = document.createElement("style");
      styleEl.id = "help-dialog-styles";
      styleEl.textContent = `
        .help-content {
          max-height: 70vh;
          overflow-y: auto;
          padding: 15px;
        }
        
        .help-content h2 {
          color: var(--primary);
          font-size: 20px;
          margin-top: 0;
          margin-bottom: 15px;
        }
        
        .help-content h3 {
          color: var(--text-primary);
          font-size: 16px;
          margin-top: 20px;
          margin-bottom: 10px;
        }
        
        .help-content p {
          margin-bottom: 15px;
          line-height: 1.5;
        }
        
        .help-content ul, .help-content ol {
          margin-bottom: 15px;
          padding-left: 25px;
        }
        
        .help-content li {
          margin-bottom: 5px;
          line-height: 1.5;
        }
        
        .help-content table {
          width: 100%;
          border-collapse: collapse;
          margin: 15px 0;
        }
        
        .help-content th, .help-content td {
          border: 1px solid var(--border);
          padding: 8px 12px;
          text-align: left;
        }
        
        .help-content th {
          background-color: var(--background);
          font-weight: 600;
        }
        
        .help-content code {
          background-color: #f5f5f5;
          padding: 2px 4px;
          border-radius: 3px;
          font-family: monospace;
        }
      `;
      document.head.appendChild(styleEl);
    }
    
    // Show the dialog
    helpDialog.style.display = "flex";
    
    // Add event listener for close button
    document.getElementById("close-help").addEventListener("click", () => {
      helpDialog.style.display = "none";
    });
  }
  
  /**
   * Add help buttons to all feature panels
   */
  export function addHelpButtons() {
    // Map of tab IDs to feature IDs
    const tabToFeatureMap = {
      "duplicates-tab": "duplicates",
      "spaces-tab": "trimSpaces",
      "case-tab": "case",
      "replace-tab": "replace",
      "dates-tab": "dates",
      "oneclick-tab": "oneclick"
    };
    
    // Add a help button to each feature highlight
    Object.entries(tabToFeatureMap).forEach(([tabId, featureId]) => {
      const tab = document.getElementById(tabId);
      if (!tab) return;
      
      const featureHighlight = tab.querySelector(".feature-highlight");
      if (!featureHighlight) return;
      
      // Check if button already exists
      if (featureHighlight.querySelector(".help-button")) return;
      
      // Create and add the help button
      const helpButton = document.createElement("button");
      helpButton.className = "help-button";
      helpButton.setAttribute("aria-label", "Show help for this feature");
      helpButton.innerHTML = '<i class="fas fa-question-circle"></i>';
      
      // Add event listener
      helpButton.addEventListener("click", () => {
        showHelpDialog(featureId);
      });
      
      // Add to the feature highlight
      featureHighlight.appendChild(helpButton);
    });
    
    // Add styles for help buttons
    if (!document.getElementById("help-button-styles")) {
      const styleEl = document.createElement("style");
      styleEl.id = "help-button-styles";
      styleEl.textContent = `
        .feature-highlight {
          position: relative;
        }
        
        .help-button {
          position: absolute;
          top: 15px;
          right: 15px;
          background: none;
          border: none;
          color: #0078d4;
          font-size: 18px;
          cursor: pointer;
          padding: 0;
          width: 24px;
          height: 24px;
          display: flex;
          align-items: center;
          justify-content: center;
        }
        
        .help-button:hover {
          color: #106ebe;
        }
      `;
      document.head.appendChild(styleEl);
    }
  }
  
  /**
   * Add a main help button to the header
   */
  export function addMainHelpButton() {
    // Check if help button already exists
    if (document.getElementById("main-help-button")) return;
    
    // Create the help button
    const helpButton = document.createElement("button");
    helpButton.id = "main-help-button";
    helpButton.className = "ms-Button ms-Button--icon";
    helpButton.title = "Help";
    helpButton.innerHTML = `
      <span class="ms-Button-icon"><i class="fas fa-question-circle" aria-hidden="true"></i></span>
      <span class="ms-Button-label">Help</span>
    `;
    
    // Add event listener
    helpButton.addEventListener("click", () => {
      showHelpDialog("general");
    });
    
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
      buttonContainer.appendChild(helpButton);
    }
  }
  
  /**
   * Add tooltips to UI elements
   */
  export function addTooltips() {
    // Define tooltips for various elements
    const tooltips = [
      {
        selector: "#remove-duplicates-button",
        text: "Remove duplicate rows from the selected range"
      },
      {
        selector: "#trim-spaces-button",
        text: "Remove extra spaces from the beginning, end, or middle of text"
      },
      {
        selector: "#uppercase-button",
        text: "Convert text to UPPERCASE"
      },
      {
        selector: "#lowercase-button",
        text: "Convert text to lowercase"
      },
      {
        selector: "#propercase-button",
        text: "Convert Text to Proper Case (capitalize first letter of each word)"
    },
    {
      selector: "#find-replace-button",
      text: "Replace all occurrences of text in the selected range"
    },
    {
      selector: "#standardize-dates-button",
      text: "Convert dates to a consistent format"
    },
    {
      selector: "#one-click-cleanup-button",
      text: "Apply multiple cleaning operations at once"
    },
    {
      selector: "#setting-show-preview",
      text: "Show a preview of changes before applying them"
    },
    {
      selector: "#undo-button",
      text: "Revert the last operation"
    }
  ];
  
  // Create tooltip container if it doesn't exist
  let tooltipContainer = document.getElementById("tooltip-container");
  if (!tooltipContainer) {
    tooltipContainer = document.createElement("div");
    tooltipContainer.id = "tooltip-container";
    tooltipContainer.className = "tooltip-container";
    document.body.appendChild(tooltipContainer);
  }
  
  // Add tooltips to each element
  tooltips.forEach(tooltip => {
    const element = document.querySelector(tooltip.selector);
    if (!element) return;
    
    // Skip if tooltip already applied
    if (element.getAttribute("data-tooltip-applied")) return;
    
    // Mark as having tooltip
    element.setAttribute("data-tooltip-applied", "true");
    element.setAttribute("aria-label", tooltip.text);
    element.setAttribute("title", tooltip.text);
    
    // Add event listeners for custom tooltip (optional enhancement)
    element.addEventListener("mouseenter", (e) => {
      showTooltip(e.target, tooltip.text);
    });
    
    element.addEventListener("mouseleave", () => {
      hideTooltip();
    });
  });
  
  // Add styles for tooltips
  if (!document.getElementById("tooltip-styles")) {
    const styleEl = document.createElement("style");
    styleEl.id = "tooltip-styles";
    styleEl.textContent = `
      .tooltip-container {
        position: absolute;
        z-index: 9999;
        background-color: #333;
        color: white;
        padding: 8px 12px;
        border-radius: 4px;
        font-size: 12px;
        max-width: 250px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
        opacity: 0;
        pointer-events: none;
        transition: opacity 0.2s;
      }
      
      .tooltip-container.visible {
        opacity: 1;
      }
      
      .tooltip-container::after {
        content: '';
        position: absolute;
        top: 100%;
        left: 50%;
        margin-left: -6px;
        border-width: 6px;
        border-style: solid;
        border-color: #333 transparent transparent transparent;
      }
    `;
    document.head.appendChild(styleEl);
  }
}

/**
 * Show a custom tooltip
 * @param {HTMLElement} targetElement - The element to show tooltip for
 * @param {string} text - Tooltip text
 */
function showTooltip(targetElement, text) {
  const tooltipContainer = document.getElementById("tooltip-container");
  if (!tooltipContainer) return;
  
  // Set text
  tooltipContainer.textContent = text;
  
  // Position tooltip above element
  const rect = targetElement.getBoundingClientRect();
  const tooltipHeight = tooltipContainer.offsetHeight;
  
  tooltipContainer.style.left = rect.left + (rect.width / 2) - (tooltipContainer.offsetWidth / 2) + "px";
  tooltipContainer.style.top = rect.top - tooltipHeight - 10 + "px";
  
  // Show tooltip
  tooltipContainer.classList.add("visible");
}

/**
 * Hide the custom tooltip
 */
function hideTooltip() {
  const tooltipContainer = document.getElementById("tooltip-container");
  if (tooltipContainer) {
    tooltipContainer.classList.remove("visible");
  }
}

/**
 * Add contextual help to input fields
 */
export function addInputHelp() {
  // Input field help definitions
  const inputHelp = [
    {
      selector: "#find-text",
      helpText: "Enter the text you want to find in your data"
    },
    {
      selector: "#replace-text",
      helpText: "Enter the text that will replace what you're looking for"
    },
    {
      selector: "input[name='dateFormat']",
      helpText: "Select your preferred date format"
    },
    {
      selector: "#setting-preserve-case",
      helpText: "List words that should keep their original capitalization (e.g., USA, iPhone)"
    }
  ];
  
  // Add help text to each input
  inputHelp.forEach(item => {
    const elements = document.querySelectorAll(item.selector);
    
    elements.forEach(element => {
      // Skip if help already applied
      if (element.getAttribute("data-help-applied")) return;
      
      // Mark as having help
      element.setAttribute("data-help-applied", "true");
      
      // Create help element
      const helpEl = document.createElement("div");
      helpEl.className = "input-help";
      helpEl.innerHTML = `<i class="fas fa-info-circle"></i> ${item.helpText}`;
      
      // Add after the input field
      if (element.parentNode) {
        element.parentNode.appendChild(helpEl);
      }
    });
  });
  
  // Add styles for input help
  if (!document.getElementById("input-help-styles")) {
    const styleEl = document.createElement("style");
    styleEl.id = "input-help-styles";
    styleEl.textContent = `
      .input-help {
        font-size: 12px;
        color: #666;
        margin-top: 4px;
        margin-left: 4px;
      }
      
      .input-help i {
        color: #0078d4;
        margin-right: 4px;
      }
    `;
    document.head.appendChild(styleEl);
  }
}

/**
 * Add a "What's New" dialog to show recent changes
 */
export function showWhatsNew() {
  // Version and changes information
  const versionInfo = {
    version: "1.2.0", // Current version
    lastShownVersion: localStorage.getItem("lastShownVersion") || "0.0.0",
    changes: [
      {
        version: "1.2.0",
        date: "February 2025",
        features: [
          "Added data preview feature to see changes before applying them",
          "Enhanced date format detection for international formats",
          "Implemented undo functionality for all operations",
          "Added customizable settings for all features",
          "Improved performance for large data sets",
          "Added comprehensive help documentation"
        ]
      },
      {
        version: "1.1.0",
        date: "January 2025",
        features: [
          "Added one-click cleanup feature",
          "Enhanced text case conversion with acronym detection",
          "Improved find and replace with regular expression support",
          "Added date standardization feature",
          "UI improvements and bug fixes"
        ]
      },
      {
        version: "1.0.0",
        date: "December 2024",
        features: [
          "Initial release",
          "Core features: remove duplicates, trim spaces, case conversion, find & replace"
        ]
      }
    ]
  };
  
  // Only show if the stored version is older than current version
  if (compareVersions(versionInfo.lastShownVersion, versionInfo.version) >= 0) {
    return;
  }
  
  // Create dialog if it doesn't exist
  let whatsNewDialog = document.getElementById("whats-new-dialog");
  if (!whatsNewDialog) {
    whatsNewDialog = document.createElement("div");
    whatsNewDialog.id = "whats-new-dialog";
    whatsNewDialog.className = "ms-Dialog ms-Dialog--lgHeader";
    whatsNewDialog.setAttribute("role", "dialog");
    whatsNewDialog.setAttribute("aria-labelledby", "whats-new-dialog-title");
    
    document.body.appendChild(whatsNewDialog);
  }
  
  // Generate content for releases since last shown
  const releaseNotes = versionInfo.changes
    .filter(change => compareVersions(versionInfo.lastShownVersion, change.version) < 0)
    .map(change => `
      <div class="release-note">
        <h3>Version ${change.version} <span class="release-date">(${change.date})</span></h3>
        <ul>
          ${change.features.map(feature => `<li>${feature}</li>`).join('')}
        </ul>
      </div>
    `)
    .join('');
  
  // Set up the dialog content
  whatsNewDialog.innerHTML = `
    <div class="ms-Dialog-title" id="whats-new-dialog-title">
      <i class="fas fa-gift" style="margin-right: 8px;"></i> What's New
    </div>
    <div class="ms-Dialog-content whats-new-content">
      <p>Excel Data Cleaner has been updated with new features and improvements!</p>
      ${releaseNotes}
    </div>
    <div class="ms-Dialog-actions">
      <button id="close-whats-new" class="ms-Button ms-Button--primary">
        <span class="ms-Button-label">Got It</span>
      </button>
    </div>
  `;
  
  // Add styles for the what's new dialog
  if (!document.getElementById("whats-new-styles")) {
    const styleEl = document.createElement("style");
    styleEl.id = "whats-new-styles";
    styleEl.textContent = `
      .whats-new-content {
        max-height: 70vh;
        overflow-y: auto;
        padding: 15px;
      }
      
      .release-note {
        margin-bottom: 20px;
      }
      
      .release-note h3 {
        color: #0078d4;
        font-size: 16px;
        margin-top: 0;
        margin-bottom: 10px;
      }
      
      .release-date {
        color: #666;
        font-size: 14px;
        font-weight: normal;
      }
      
      .release-note ul {
        margin-bottom: 0;
        padding-left: 25px;
      }
      
      .release-note li {
        margin-bottom: 5px;
        line-height: 1.5;
      }
    `;
    document.head.appendChild(styleEl);
  }
  
  // Show the dialog
  whatsNewDialog.style.display = "block";
  
  // Add event listener for close button
  document.getElementById("close-whats-new").addEventListener("click", () => {
    // Store the current version as last shown
    localStorage.setItem("lastShownVersion", versionInfo.version);
    whatsNewDialog.style.display = "none";
  });
}

/**
 * Compare two version strings
 * @param {string} v1 - First version
 * @param {string} v2 - Second version
 * @returns {number} -1 if v1 < v2, 0 if v1 == v2, 1 if v1 > v2
 */
function compareVersions(v1, v2) {
  const v1Parts = v1.split('.').map(Number);
  const v2Parts = v2.split('.').map(Number);
  
  for (let i = 0; i < Math.max(v1Parts.length, v2Parts.length); i++) {
    const v1Part = v1Parts[i] || 0;
    const v2Part = v2Parts[i] || 0;
    
    if (v1Part < v2Part) return -1;
    if (v1Part > v2Part) return 1;
  }
  
  return 0;
}

/**
 * Initialize help documentation system
 */
export function initializeHelpSystem() {
  addHelpButtons();
  addMainHelpButton();
  addTooltips();
  addInputHelp();
  
  // Show what's new dialog on startup (if there's a new version)
  setTimeout(() => {
    showWhatsNew();
  }, 1000);
}