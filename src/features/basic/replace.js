/*
 * Excel Data Cleaner Add-in
 * Find and replace functionality
 */

import { showMessage, showError, showLoading, hideLoading } from '../../utils/ui-utils.js';
import { checkTrialLimits, incrementOperationCount } from '../../utils/trial.js';

/**
 * Find and replace text in selected range
 */
export async function findAndReplace() {
  try {
    // Check trial limits
    if (!checkTrialLimits()) return;
    
    const findText = document.getElementById("find-text").value;
    const replaceText = document.getElementById("replace-text").value;
    
    if (!findText) {
      showMessage("Please enter text to find.", "error");
      return;
    }
    
    // Show loading overlay
    showLoading(`Replacing "${findText}" with "${replaceText}"...`);
    
    await Excel.run(async (context) => {
      // Get the selected range
      const range = context.workbook.getSelectedRange();
      range.load(["values", "cellCount"]);
      await context.sync();
      
      // Check if a range is selected
      if (range.cellCount === 0) {
        hideLoading();
        showMessage("Please select a range with data first.", "error");
        return;
      }
      
      // Get the current values
      const values = range.values;
      let replacementCount = 0;
      let cellsAffected = 0;
      
      // Replace text in each cell
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
          if (typeof values[i][j] === "string") {
            // Use regular expression with global flag to replace all occurrences
            const originalText = values[i][j];
            
            try {
              // Create a RegExp with 'g' flag for global replacement
              const regex = new RegExp(escapeRegExp(findText), "g");
              const newText = originalText.replace(regex, replaceText);
              
              // Count replacements
              if (newText !== originalText) {
                // Count how many replacements were made
                const matchCount = (originalText.match(regex) || []).length;
                replacementCount += matchCount;
                cellsAffected++;
                values[i][j] = newText;
              }
            } catch (regexError) {
              // If there's an error with the regex, fall back to simple replace
              const newText = originalText.split(findText).join(replaceText);
              if (newText !== originalText) {
                // Estimate replacements based on length changes
                const matchCount = (originalText.split(findText).length - 1);
                replacementCount += matchCount;
                cellsAffected++;
                values[i][j] = newText;
              }
            }
          }
        }
      }
      
      // Set the updated values back to the range
      range.values = values;
      
      await context.sync();
      
      if (replacementCount > 0) {
        incrementOperationCount();
        hideLoading();
        showMessage(`Success! Replaced ${replacementCount} occurrences in ${cellsAffected} cells.`, "success");
      } else {
        hideLoading();
        showMessage(`"${findText}" not found in the selected range.`, "info");
      }
    });
  } catch (error) {
    hideLoading();
    showError(error);
  }
}

/**
 * Escape regular expression special characters in a string
 * @param {string} string - The input string
 * @returns {string} The escaped string
 */
export function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}