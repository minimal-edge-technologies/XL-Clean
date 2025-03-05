/*
 * Excel Data Cleaner Add-in
 * Duplicate removal functionality
 */

import { showMessage, showError, showLoading, hideLoading } from '../../utils/ui-utils.js';
import { checkTrialLimits, isPremiumUser, incrementOperationCount } from '../../utils/trial.js';

/**
 * Remove duplicates from selected range
 */
export async function removeDuplicates() {
  try {
    // Check trial limits
    if (!checkTrialLimits()) return;

     // Show loading overlay with custom message
     showLoading("Removing duplicates...");
    
    await Excel.run(async (context) => {
      // Get the selected range
      const range = context.workbook.getSelectedRange();
      range.load(["rowCount", "values", "address", "cellCount"]);
      await context.sync();
      
      // Check if a range is selected
      if (range.cellCount === 0) {
        hideLoading();
        showMessage("Please select a range with data first.", "error");
        return;
      }
      
      // Trial version row limitation check
      if (!isPremiumUser() && range.rowCount > 100) {
        hideLoading();
        showMessage("Trial version is limited to 100 rows. Upgrade to process more data.", "error");
        return;
      }
      
      // Get the current values
      const values = range.values;
      if (values.length === 0) {
        hideLoading();
        showMessage("Please select a range with data first.", "error");
        return;
      }
      
      // Create a map to track unique rows (as strings)
      const uniqueRows = new Map();
      const duplicateRowIndices = [];
      
      // Identify duplicate rows
      for (let i = 0; i < values.length; i++) {
        // Convert row to string for comparison
        const rowString = JSON.stringify(values[i]);
        
        if (uniqueRows.has(rowString)) {
          // This is a duplicate row
          duplicateRowIndices.push(i);
        } else {
          // This is a unique row so far
          uniqueRows.set(rowString, i);
        }
      }
      
      // If no duplicates found
      if (duplicateRowIndices.length === 0) {
        hideLoading();
        showMessage("No duplicates found in the selected range.", "info");
        return;
      }
      
      // Get the worksheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const originalRangeAddress = range.address;
      
      // Filter out the duplicate rows
      const uniqueValues = values.filter((_, index) => !duplicateRowIndices.includes(index));
      
      // Write the unique values back to the range
      // First, clear the original range
      range.clear();
      await context.sync();
      
      // Get the top-left cell of the original range
      const topLeftCell = sheet.getRange(originalRangeAddress.split(':')[0]);
      
      // Calculate the new range size
      const newRange = topLeftCell.getResizedRange(uniqueValues.length - 1, values[0].length - 1);
      
      // Set the unique values
      newRange.values = uniqueValues;
      
      await context.sync();
      incrementOperationCount();

      hideLoading();
      showMessage(`Success! Removed ${duplicateRowIndices.length} duplicate rows.`, "success");
    });
  } catch (error) {
    hideLoading();
    showError(error);
  }
}

/**
 * Helper function to remove duplicates (used by one-click cleanup)
 * @param {Excel.RequestContext} context - The Excel request context
 * @param {Excel.Range} range - The range to process
 * @returns {Promise<{duplicatesRemoved: number}>} The number of duplicates removed
 */
export async function removeDuplicatesHelper(context, range) {
  range.load(["values", "address"]);
  await context.sync();
  
  // Get the current values
  const values = range.values;
  if (values.length <= 1) {
    return { duplicatesRemoved: 0 }; // Not enough rows to have duplicates
  }
  
  // Create a map to track unique rows (as strings)
  const uniqueRows = new Map();
  const duplicateRowIndices = [];
  
  // Identify duplicate rows
  for (let i = 0; i < values.length; i++) {
    // Convert row to string for comparison
    const rowString = JSON.stringify(values[i]);
    
    if (uniqueRows.has(rowString)) {
      // This is a duplicate row
      duplicateRowIndices.push(i);
    } else {
      // This is a unique row so far
      uniqueRows.set(rowString, i);
    }
  }
  
  // If no duplicates found
  if (duplicateRowIndices.length === 0) {
    return { duplicatesRemoved: 0 };
  }
  
  // Get the worksheet
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const originalRangeAddress = range.address;
  
  // Filter out the duplicate rows
  const uniqueValues = values.filter((_, index) => !duplicateRowIndices.includes(index));
  
  // Write the unique values back to the range
  // First, clear the original range
  range.clear();
  await context.sync();
  
  // Get the top-left cell of the original range
  const topLeftCell = sheet.getRange(originalRangeAddress.split(':')[0]);
  
  // Calculate the new range size
  if (uniqueValues.length > 0) {
    const newRange = topLeftCell.getResizedRange(uniqueValues.length - 1, values[0].length - 1);
    newRange.values = uniqueValues;
    await context.sync();
  }
  
  return { duplicatesRemoved: duplicateRowIndices.length };
}