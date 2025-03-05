/*
 * Excel Data Cleaner Add-in
 * Excel helper utilities
 */

/**
 * Check if a range is selected
 * @param {Excel.Range} range - The range to check
 * @returns {boolean} True if a range is selected, false otherwise
 */
export function isRangeSelected(range) {
    return range && range.cellCount > 0;
  }
  
  /**
   * Check if a value is likely a date
   * @param {any} value - The value to check
   * @returns {boolean} True if the value is likely a date, false otherwise
   */
  export function isLikelyDate(value) {
    // If it's a date object
    if (value instanceof Date) return true;
    
    // If it's a number that could be an Excel date serial number
    if (typeof value === "number" && value > 1000 && value < 50000) return true;
    
    // If it's a string with date-like format
    if (typeof value === "string") {
      // Check for common date formats
      if (/^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}$/.test(value)) return true;
      if (/^\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}$/.test(value)) return true;
      
      // Try parsing with Date constructor
      const date = new Date(value);
      if (!isNaN(date) && date.getFullYear() > 1900) return true;
    }
    
    return false;
  }
  
  /**
   * Get a summary of range contents
   * @param {Excel.Range} range - The range to analyze
   * @returns {Promise<Object>} A summary of the range contents
   */
  export async function getRangeSummary(context, range) {
    range.load(["values", "rowCount", "columnCount", "cellCount"]);
    await context.sync();
    
    const values = range.values;
    const summary = {
      totalCells: range.cellCount,
      textCells: 0,
      numberCells: 0,
      dateCells: 0,
      emptyCells: 0,
      duplicates: 0
    };
    
    const uniqueValues = new Set();
    
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        const cellValue = values[i][j];
        
        // Check cell type
        if (cellValue === null || cellValue === undefined || cellValue === "") {
          summary.emptyCells++;
        } else if (typeof cellValue === "number") {
          summary.numberCells++;
          
          // Check if it might be a date
          if (isLikelyDate(cellValue)) {
            summary.dateCells++;
          }
        } else if (typeof cellValue === "string") {
          summary.textCells++;
          
          // Check if string might be a date
          if (isLikelyDate(cellValue)) {
            summary.dateCells++;
          }
        }
        
        // Track unique values
        const valueString = JSON.stringify(cellValue);
        if (uniqueValues.has(valueString)) {
          summary.duplicates++;
        } else {
          uniqueValues.add(valueString);
        }
      }
    }
    
    return summary;
  }
  
  /**
   * Get column names from a range
   * @param {Excel.Range} range - The range to analyze
   * @returns {Promise<string[]>} An array of column names
   */
  export async function getColumnNames(context, range) {
    // Get the actual range that has data
    const rangeWithData = range.getUsedRange();
    rangeWithData.load(["values", "rowCount", "columnCount"]);
    await context.sync();
    
    // If there's no data, return empty array
    if (rangeWithData.rowCount === 0 || rangeWithData.columnCount === 0) {
      return [];
    }
    
    // Get the first row which usually contains headers
    const headerRow = rangeWithData.getRow(0);
    headerRow.load("values");
    await context.sync();
    
    // Convert the 2D array to 1D array of column names
    const columnNames = headerRow.values[0].map(header => 
      header ? header.toString() : "");
    
    return columnNames;
  }
  
  /**
   * Apply formatting to a range based on content type
   * @param {Excel.Range} range - The range to format
   */
  export async function applySmartFormatting(context, range) {
    range.load(["values", "rowCount", "columnCount"]);
    await context.sync();
    
    const values = range.values;
    
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        const cell = range.getCell(i, j);
        const cellValue = values[i][j];
        
        if (isLikelyDate(cellValue)) {
          // Format as date
          cell.numberFormat = "mm/dd/yyyy";
        } else if (typeof cellValue === "number") {
          // Format as number with commas
          cell.numberFormat = "#,##0.00";
        } else if (typeof cellValue === "string" && cellValue.trim().startsWith("$")) {
          // Format as currency
          cell.numberFormat = "$#,##0.00";
        }
      }
    }
    
    await context.sync();
  }