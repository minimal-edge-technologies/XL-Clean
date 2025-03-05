/*
 * Excel Data Cleaner Add-in
 * One-click cleanup functionality
 */

import { showMessage, showError, showLoading, hideLoading } from '../../utils/ui-utils.js';
import { removeDuplicatesHelper } from '../basic/duplicates.js';
import { convertToProperCase, convertToSentenceCase } from '../basic/case.js';
/**
 * Perform one-click cleanup on the selected range
 */
export async function oneClickCleanup() {
  try {
    // Show loading overlay
    showLoading("Running one-click cleanup...");
    
    await Excel.run(async (context) => {
      // Start tracking the changes
      const results = {
        trimmedCells: 0,
        caseFixedCells: 0,
        numberFixedCells: 0,
        duplicatesRemoved: 0
      };
      
      // Get the selected range
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount", "address", "cellCount"]);
      await context.sync();
      
      // Check if a range is selected
      if (range.cellCount === 0) {
        hideLoading();
        showMessage("Please select a range with data first.", "error");
        return;
      }
      
      // Check which cleanup operations are selected
      const cleanupDuplicates = document.getElementById("cleanup-duplicates").checked;
      const cleanupSpaces = document.getElementById("cleanup-spaces").checked;
      const cleanupCase = document.getElementById("cleanup-case").checked;
      const cleanupFormatting = document.getElementById("cleanup-formatting").checked;
      
      // Update loading message based on selected operations
      let operations = [];
      if (cleanupSpaces) operations.push("trimming spaces");
      if (cleanupCase) operations.push("fixing text case");
      if (cleanupFormatting) operations.push("formatting numbers");
      if (cleanupDuplicates) operations.push("removing duplicates");
      
      if (operations.length > 0) {
        showLoading(`Cleaning data: ${operations.join(", ")}...`);
      }
      
      // Get the current values
      let values = range.values;
      
      // Make a backup for comparison
      const originalValues = JSON.parse(JSON.stringify(values));
      
      // Step 1: Trim spaces if selected
      if (cleanupSpaces) {
        values = await cleanupExtraSpaces(values, results);
      }
      
      // Step 2: Fix text case if selected
      if (cleanupCase) {
        values = await fixTextCase(values, results);
      }
      
      // Step 3: Fix number formatting if selected
      if (cleanupFormatting) {
        values = await fixNumberFormatting(values, results);
      }
      
      // Update the range with cleaned values
      range.values = values;
      await context.sync();
      
      // Step 4: Remove duplicates if selected
      if (cleanupDuplicates && range.rowCount > 1) {
        showLoading("Removing duplicate rows...");
        const duplicateResult = await removeDuplicatesHelper(context, range);
        results.duplicatesRemoved = duplicateResult.duplicatesRemoved;
      }
      
      // Show summary of changes
      hideLoading();
      showCleanupResults(results);
    });
  } catch (error) {
    hideLoading();
    showError(error);
  }
}

/**
 * Clean up extra spaces in cell values
 * @param {any[][]} values - The current values
 * @param {object} results - The results tracking object
 * @returns {any[][]} The updated values
 */
async function cleanupExtraSpaces(values, results) {
  // Create a copy of values to modify
  const newValues = [...values.map(row => [...row])];
  
  for (let i = 0; i < newValues.length; i++) {
    for (let j = 0; j < newValues[i].length; j++) {
      if (typeof newValues[i][j] === "string") {
        const original = newValues[i][j];
        
        // Trim leading/trailing spaces and reduce multiple spaces to one
        newValues[i][j] = original.trim().replace(/\s+/g, ' ');
        
        if (newValues[i][j] !== original) {
          results.trimmedCells++;
        }
      }
    }
  }
  
  return newValues;
}

/**
 * Fix text case in cell values
 * @param {any[][]} values - The current values
 * @param {object} results - The results tracking object
 * @returns {any[][]} The updated values
 */
async function fixTextCase(values, results) {
  // Create a copy of values to modify
  const newValues = [...values.map(row => [...row])];
  
  for (let i = 0; i < newValues.length; i++) {
    for (let j = 0; j < newValues[i].length; j++) {
      if (typeof newValues[i][j] === "string" && newValues[i][j].length > 0) {
        const original = newValues[i][j];
        
        // Determine the appropriate case transformation
        if (original.length > 20 || original.includes('.') || original.includes('?') || original.includes('!')) {
          // Sentence case for longer text or text with punctuation
          newValues[i][j] = convertToSentenceCase(original);
        } else {
          // Title case for shorter text (likely titles, names, etc.)
          newValues[i][j] = convertToProperCase(original);
        }
        
        if (newValues[i][j] !== original) {
          results.caseFixedCells++;
        }
      }
    }
  }
  
  return newValues;
}

/**
 * Fix number formatting in cell values
 * @param {any[][]} values - The current values
 * @param {object} results - The results tracking object
 * @returns {any[][]} The updated values
 */
async function fixNumberFormatting(values, results) {
  // Create a copy of values to modify
  const newValues = [...values.map(row => [...row])];
  
  for (let i = 0; i < newValues.length; i++) {
    for (let j = 0; j < newValues[i].length; j++) {
      const cellValue = newValues[i][j];
      
      // Check for numbers stored as text
      if (typeof cellValue === "string") {
        // Remove extra spaces and commas
        const cleanedStr = cellValue.trim().replace(/,/g, '');
        
        // Try to convert to number if it looks like a number (not alphanumeric)
        if (/^-?\d+(\.\d+)?$/.test(cleanedStr)) {
          const numericValue = parseFloat(cleanedStr);
          if (!isNaN(numericValue)) {
            newValues[i][j] = numericValue;
            results.numberFixedCells++;
          }
        }
      }
    }
  }
  
  return newValues;
}

/**
 * Show the results of the cleanup operation
 * @param {object} results - The results of the cleanup
 */
function showCleanupResults(results) {
  let message = "One-click cleanup results:\n";
  
  if (results.trimmedCells > 0) {
    message += `• Trimmed spaces in ${results.trimmedCells} cells\n`;
  }
  
  if (results.caseFixedCells > 0) {
    message += `• Fixed text case in ${results.caseFixedCells} cells\n`;
  }
  
  if (results.numberFixedCells > 0) {
    message += `• Fixed number formatting in ${results.numberFixedCells} cells\n`;
  }
  
  if (results.duplicatesRemoved > 0) {
    message += `• Removed ${results.duplicatesRemoved} duplicate rows\n`;
  }
  
  if (results.trimmedCells === 0 && 
      results.caseFixedCells === 0 && 
      results.numberFixedCells === 0 && 
      results.duplicatesRemoved === 0) {
    message = "No changes were needed in the selected range.";
  }
  
  showMessage(message, "success");
}