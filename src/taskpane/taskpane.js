/*
 * Excel Data Cleaner Add-in
 * Main entry point for the taskpane
 */

// Import basic features
import { removeDuplicates } from '../features/basic/duplicates.js';
import { trimSpaces } from '../features/basic/trim.js';
import { convertCase } from '../features/basic/case.js';
import { findAndReplace } from '../features/basic/replace.js';

// Import advanced features (formerly premium)
import { standardizeDates } from '../features/advanced/dates.js';
import { oneClickCleanup } from '../features/advanced/oneclick.js';

// Import utils
import { showMessage, showError, showLoading, hideLoading } from '../utils/ui-utils.js';

// Import new feature utilities
import { initializeUndoFeature, performUndo, addToUndoStack } from '../utils/undo-functionality.js';
import { detectMostCommonDateFormat } from '../utils/enhanced-date-detection.js';
import { initializeSettings, showSettingsDialog } from '../utils/customizable-settings.js';
import { initializeHelpSystem, showHelpDialog } from '../utils/help-documentation.js';
import { initializeAccessibility } from '../utils/accessibility-helpers.js';

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Initialize all systems
    initializeSettings();
    initializeUndoFeature();
    initializeHelpSystem();
    initializeAccessibility();
    initializeTabs();
  
    // Connect settings and help buttons directly
    connectButtons();

    // Register event handlers for basic features
    document.getElementById("remove-duplicates-button").onclick = function() {
      executeOperation(removeDuplicates, "Remove Duplicates");
    };
    document.getElementById("trim-spaces-button").onclick = function() {
      executeOperation(trimSpaces, "Trim Spaces");
    };
    document.getElementById("uppercase-button").onclick = function() {
      executeOperation(() => convertCase("UPPER"), "Convert to UPPERCASE");
    };
    document.getElementById("lowercase-button").onclick = function() {
      executeOperation(() => convertCase("LOWER"), "Convert to lowercase");
    };
    document.getElementById("propercase-button").onclick = function() {
      executeOperation(() => convertCase("PROPER"), "Convert to Proper Case");
    };
    document.getElementById("find-replace-button").onclick = function() {
      executeOperation(findAndReplace, "Find and Replace");
    };
    document.getElementById("standardize-dates-button").onclick = function() {
      executeOperation(standardizeDates, "Standardize Dates");
    };
    document.getElementById("one-click-cleanup-button").onclick = function() {
      executeOperation(oneClickCleanup, "One-Click Cleanup");
    };

    // Register date detection
    document.getElementById("detect-format-button").onclick = detectDateFormat;
    
    // Show user message that the add-in is ready
    showMessage("Excel Data Cleaner is ready to use!", "info");
  }
});

/**
 * Connect settings and help buttons to their functions
 */
function connectButtons() {
  // Simple direct connection for settings button
  const settingsButton = document.getElementById("settings-button");
  if (settingsButton) {
    // Make sure there's only one event listener
    settingsButton.onclick = function() {
      showSettingsDialog();
    };
  }
  
  // Simple direct connection for help button
  const helpButton = document.getElementById("help-button");
  if (helpButton) {
    // Make sure there's only one event listener
    helpButton.onclick = function() {
      showHelpDialog("general");
    };
  }
  
  console.log("Settings and help buttons connected");
}

/**
 * Execute operation with fast response and undo support
 * This separates the undo state saving from the actual operation
 * for better performance while maintaining undo functionality
 */
async function executeOperation(func, operationName) {
  try {
    // Show loading immediately for better responsiveness
    showLoading(`Running ${operationName}...`);
    
    // Check if a range is selected first
    let hasSelection = false;
    
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("cellCount");
      await context.sync();
      
      if (range.cellCount === 0) {
        showMessage("Please select a range first.", "error");
        hideLoading();
        return;
      }
      
      hasSelection = true;
    });
    
    if (!hasSelection) {
      return;
    }
    
    // First, we need to save the state for undo
    let saveSuccessful = await saveStateForUndo(operationName);
    
    if (saveSuccessful) {
      // Now execute the operation
      await func();
    } else {
      hideLoading();
    }
  } catch (error) {
    console.error(`Error in ${operationName}:`, error);
    hideLoading();
    showError(error);
  }
}

/**
 * Save the current state for undo
 */
async function saveStateForUndo(operationName) {
  try {
    let state = null;
    
    await Excel.run(async (context) => {
      // Get the selected range
      const range = context.workbook.getSelectedRange();
      
      // Load necessary properties
      range.load(["address", "values", "formulas", "cellCount", "worksheet"]);
      await context.sync();
      
      // Check if there's anything selected
      if (range.cellCount === 0) {
        showMessage("Please select a range first.", "error");
        return;
      }
      
      // Get worksheet name
      const worksheetName = range.worksheet.name;
      
      // Save state
      state = {
        timestamp: new Date(),
        operation: operationName,
        worksheetName: worksheetName,
        address: range.address,
        values: JSON.parse(JSON.stringify(range.values)),
        formulas: JSON.parse(JSON.stringify(range.formulas))
      };
    });
    
    // If we have a state, add it to the undo stack
    if (state) {
      addToUndoStack(state);
      return true;
    }
    
    return false;
  } catch (error) {
    console.error("Error saving undo state:", error);
    return false;
  }
}

/**
 * Initialize tab navigation
 */
function initializeTabs() {
  const tabs = document.querySelectorAll(".tab-button");
  
  tabs.forEach(tab => {
    tab.addEventListener("click", function() {
      // Remove active class from all tabs
      tabs.forEach(t => {
        t.classList.remove("active");
      });
      
      // Add active class to clicked tab
      this.classList.add("active");
      
      // Hide all tab content
      const tabContents = document.querySelectorAll(".tab-content");
      tabContents.forEach(content => {
        content.classList.remove("active");
      });
      
      // Show the corresponding tab content
      const tabId = this.getAttribute("data-tab");
      document.getElementById(`${tabId}-tab`).classList.add("active");
    });
  });
}

async function detectDateFormat() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("cellCount");
    await context.sync();
    
    if (range.cellCount === 0) {
      showMessage("Please select cells containing dates first.", "error");
      return;
    }
    
    showLoading("Detecting date format...");
    
    try {
      const detectedFormat = await detectMostCommonDateFormat(context, range);
      
      // Select the appropriate radio button
      const formatRadios = document.getElementsByName("dateFormat");
      for (const radio of formatRadios) {
        if (radio.value === detectedFormat) {
          radio.checked = true;
          break;
        }
      }
      
      hideLoading();
      showMessage(`Detected date format: ${detectedFormat}`, "success");
    } catch (error) {
      hideLoading();
      showError(error);
    }
  });
}