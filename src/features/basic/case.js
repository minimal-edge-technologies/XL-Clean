/*
 * Excel Data Cleaner Add-in
 * Text case conversion functionality
 */

import { showMessage, showError, showLoading, hideLoading } from '../../utils/ui-utils.js';
import { checkTrialLimits, isPremiumUser, incrementOperationCount } from '../../utils/trial.js';

/**
 * Convert text case in selected cells
 * @param {string} caseType - The type of case conversion (UPPER, LOWER, PROPER)
 */
export async function convertCase(caseType) {
  try {
    // Check trial limits
    if (!checkTrialLimits()) return;
    
    // Show loading overlay with case-specific message
    let message = "Converting text case...";
    if (caseType === "UPPER") message = "Converting to UPPERCASE...";
    if (caseType === "LOWER") message = "Converting to lowercase...";
    if (caseType === "PROPER") message = "Converting to Proper Case...";
    
    showLoading(message);
    
    await Excel.run(async (context) => {
      // Get the selected range
      const range = context.workbook.getSelectedRange();
      range.load(["values", "columnCount", "rowCount", "cellCount"]);
      await context.sync();
      
      // Check if a range is selected
      if (range.cellCount === 0) {
        hideLoading();
        showMessage("Please select cells containing text first.", "error");
        return;
      }
      
      // Trial version column limitation check
      if (!isPremiumUser() && range.columnCount > 2) {
        hideLoading();
        showMessage("Trial version is limited to 2 columns. Upgrade to process more data.", "error");
        return;
      }
      
      // Get the current values
      const values = range.values;
      let conversionCount = 0;
      
      // Convert case for each cell value if it's a string
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
          // Handle null or undefined cells
          if (values[i][j] === null || values[i][j] === undefined) {
            continue;
          }
          
          // Convert to string if it's not already
          let cellValue = String(values[i][j]);
          let convertedValue = cellValue;
          
          switch (caseType) {
            case "UPPER":
              convertedValue = cellValue.toUpperCase();
              break;
            case "LOWER":
              convertedValue = cellValue.toLowerCase();
              break;
            case "PROPER":
              // This handles proper case conversion more accurately
              convertedValue = convertToProperCase(cellValue);
              break;
          }
          
          // Only update if there's a change
          if (convertedValue !== cellValue) {
            values[i][j] = convertedValue;
            conversionCount++;
          }
        }
      }
      
      // Set the converted values back to the range
      range.values = values;
      
      await context.sync();
      
      // Only increment operation count if changes were made
      if (conversionCount > 0) {
        incrementOperationCount();
        hideLoading();
        showMessage(`Success! Converted case for ${conversionCount} cells.`, "success");
      } else {
        hideLoading();
        showMessage("No text values found to convert.", "info");
      }
    });
  } catch (error) {
    hideLoading();
    showError(error);
  }
}

/**
 * Convert text to proper case
 * @param {string} text - The text to convert
 * @returns {string} The converted text
 */
export function convertToProperCase(text) {
  if (!text) return text;
  
  return text.toLowerCase()
    .replace(/(^|\s|\(|\[|\{|"|')(\S)/g, function(match, p1, p2) {
      return p1 + p2.toUpperCase();
    });
}

/**
 * Convert text to sentence case
 * @param {string} text - The text to convert
 * @returns {string} The converted text
 */
export function convertToSentenceCase(text) {
  if (!text) return text;
  
  return text.toLowerCase()
    .replace(/(^\s*|\.\s*|\?\s*|\!\s*)([a-z])/g, function(match, p1, p2) {
      return p1 + p2.toUpperCase();
    });
}

/**
 * Intelligently convert text case based on content
 * @param {string} text - The text to convert
 * @returns {string} The converted text
 */
export function convertTextCaseIntelligently(text) {
  if (!text) return text;
  
  // Detect if it's likely to be a sentence or a title
  if (text.length > 20 || text.includes('.') || text.includes('?') || text.includes('!')) {
    // Sentence case for longer text or text with sentence terminators
    return convertToSentenceCase(text);
  } else {
    // Title case for shorter text (likely titles, names, etc.)
    return convertToProperCase(text);
  }
}