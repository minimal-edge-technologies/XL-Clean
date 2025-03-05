/*
 * Excel Data Cleaner Add-in
 * Undo Functionality
 */

import { showMessage, showError, showLoading, hideLoading } from './ui-utils.js';

// Track undo history
const undoStack = [];
const MAX_UNDO_STACK_SIZE = 10;

/**
 * Add an operation to the undo stack
 * @param {Object} state - The state to save for undo
 */
export function addToUndoStack(state) {
  undoStack.push(state);
  
  // Limit stack size
  if (undoStack.length > MAX_UNDO_STACK_SIZE) {
    undoStack.shift(); // Remove oldest entry
  }
  
  // Update the UI
  updateUndoUI();
  
  console.log(`Added "${state.operation}" to undo stack`);
}

/**
 * Perform an undo operation
 * @returns {Promise<boolean>} True if undo was successful
 */
export async function performUndo() {
  if (undoStack.length === 0) {
    console.log("Nothing to undo");
    showMessage("Nothing to undo", "info");
    return false;
  }
  
  try {
    // Show loading indicator
    showLoading("Undoing last operation...");
    
    // Get the last operation from the stack
    const lastOperation = undoStack.pop();
    console.log(`Performing undo for "${lastOperation.operation}" on range ${lastOperation.address}`);
    
    await Excel.run(async (context) => {
      try {
        // Get the worksheet and range
        const worksheet = context.workbook.worksheets.getItem(lastOperation.worksheetName);
        const range = worksheet.getRange(lastOperation.address);
        
        // Prepare restored values
        const restoreValues = [];
        
        for (let i = 0; i < lastOperation.values.length; i++) {
          restoreValues[i] = [];
          for (let j = 0; j < lastOperation.values[i].length; j++) {
            if (lastOperation.formulas[i][j] && 
                String(lastOperation.formulas[i][j]).startsWith('=')) {
              restoreValues[i][j] = lastOperation.formulas[i][j];
            } else {
              restoreValues[i][j] = lastOperation.values[i][j];
            }
          }
        }
        
        // Restore the values/formulas
        range.values = restoreValues;
        
        await context.sync();
        console.log("Undo completed successfully");
      } catch (innerError) {
        console.error("Error during undo restore:", innerError);
        throw innerError;
      }
    });
    
    // Update the UI to reflect undo stack state
    updateUndoUI();
    
    // Hide loading indicator
    hideLoading();
    
    // Show success message
    showMessage(`Successfully undid "${lastOperation.operation}"`, "success");
    
    return true;
  } catch (error) {
    console.error("Undo failed:", error);
    
    // Hide loading indicator
    hideLoading();
    
    // Show error message
    showError(error);
    
    return false;
  }
}

/**
 * Check if undo is available
 * @returns {boolean} True if undo is available
 */
export function canUndo() {
  return undoStack.length > 0;
}

/**
 * Get the last operation name
 * @returns {string|null} Name of the last operation or null if none
 */
export function getLastOperationName() {
  if (undoStack.length === 0) return null;
  return undoStack[undoStack.length - 1].operation;
}

/**
 * Clear the undo stack
 */
export function clearUndoStack() {
  undoStack.length = 0;
  updateUndoUI();
}

/**
 * Update the undo UI elements
 */
function updateUndoUI() {
  const undoButton = document.getElementById("undo-button");
  if (!undoButton) return;
  
  if (canUndo()) {
    undoButton.removeAttribute("disabled");
    const lastOp = getLastOperationName();
    undoButton.title = `Undo ${lastOp}`;
  } else {
    undoButton.setAttribute("disabled", "disabled");
    undoButton.title = "Nothing to undo";
  }
}

/**
 * Initialize undo feature and create/connect the undo button
 */
export function initializeUndoFeature() {
  console.log("Initializing undo feature");
  
  // Connect existing undo button if present
  const existingButton = document.getElementById("undo-button");
  if (existingButton) {
    connectUndoButton(existingButton);
    return;
  }
  
  // Otherwise, create the undo button
  const undoButton = document.createElement("button");
  undoButton.id = "undo-button";
  undoButton.className = "icon-button";
  undoButton.setAttribute("disabled", "disabled");
  undoButton.title = "Nothing to undo";
  undoButton.innerHTML = `
    <i class="fas fa-undo"></i>
  `;
  
  // Connect the button
  connectUndoButton(undoButton);
  
  // Add to DOM - place it in the header actions div
  const headerActions = document.querySelector(".header-actions");
  if (headerActions) {
    headerActions.prepend(undoButton);
  }
}

/**
 * Connect event handlers to an undo button
 * @param {HTMLElement} button - The undo button element
 */
function connectUndoButton(button) {
  // Remove any existing listeners
  const newButton = button.cloneNode(true);
  if (button.parentNode) {
    button.parentNode.replaceChild(newButton, button);
  }
  
  // Add event listener
  newButton.addEventListener("click", async () => {
    if (canUndo()) {
      await performUndo();
    }
  });
}