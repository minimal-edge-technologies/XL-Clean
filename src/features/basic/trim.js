/*
 * Excel Data Cleaner Add-in
 * Trim spaces functionality
 */

import { showMessage, showError, showLoading, hideLoading } from '../../utils/ui-utils.js';
import { checkTrialLimits, incrementOperationCount } from '../../utils/trial.js';

/**
 * Trim extra spaces from selected cells
 */
export async function trimSpaces() {
  try {
    // Check trial limits
    if (!checkTrialLimits()) return;
    
    // Show loading overlay
    showLoading("Trimming extra spaces...");
    
    await Excel.run(async (context) => {
      // Get the selected range
      const range = context.workbook.getSelectedRange();
      range.load(["values", "cellCount"]);
      await context.sync();
      
      // Check if a range is selected
      if (range.cellCount === 0) {
        hideLoading();
        showMessage("Please select cells containing text first.", "error");
        return;
      }
      
      // Get the current values
      const values = range.values;
      let trimCount = 0;
      
      // Trim spaces from each cell value if it's a string
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
          if (typeof values[i][j] === "string") {
            const original = values[i][j];
            // Trim spaces and reduce multiple spaces to single spaces
            values[i][j] = original.trim().replace(/\s+/g, ' ');
            
            if (values[i][j] !== original) {
              trimCount++;
            }
          }
        }
      }
      
      // Set the trimmed values back to the range
      range.values = values;
      
      await context.sync();
      
      // Only count as an operation if changes were made
      if (trimCount > 0) {
        incrementOperationCount();
        hideLoading();
        showMessage(`Success! Trimmed extra spaces in ${trimCount} cells.`, "success");
      } else {
        hideLoading();
        showMessage("No spaces to trim in the selected range.", "info");
      }
    });
  } catch (error) {
    hideLoading();
    showError(error);
  }
}

/**
 * Helper function to trim spaces for use in one-click cleanup
 * @param {any[][]} values - The current cell values
 * @returns {Object} Object containing new values and count of trimmed cells
 */
export function trimSpacesHelper(values) {
  // Create a copy of values to modify
  const newValues = [...values.map(row => [...row])];
  let trimCount = 0;
  
  for (let i = 0; i < newValues.length; i++) {
    for (let j = 0; j < newValues[i].length; j++) {
      if (typeof newValues[i][j] === "string") {
        const original = newValues[i][j];
        newValues[i][j] = original.trim().replace(/\s+/g, ' ');
        
        if (newValues[i][j] !== original) {
          trimCount++;
        }
      }
    }
  }
  
  return {
    values: newValues,
    trimmedCells: trimCount
  };
}