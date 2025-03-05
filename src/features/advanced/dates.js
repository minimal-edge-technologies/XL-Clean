/*
 * Excel Data Cleaner Add-in
 * Date standardization functionality
 */

import { showMessage, showError, showLoading, hideLoading } from '../../utils/ui-utils.js';

/**
 * Standardize dates in the selected range
 */
export async function standardizeDates() {
  try {
    // Get selected date format
    const formatRadios = document.getElementsByName("dateFormat");
    let selectedFormat = "MM/DD/YYYY"; // Default
    for (const radio of formatRadios) {
      if (radio.checked) {
        selectedFormat = radio.value;
        break;
      }
    }
    
    // Show loading overlay with format-specific message
    showLoading(`Standardizing dates to ${selectedFormat} format...`);
    
    await Excel.run(async (context) => {
      // Get the selected range
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "cellCount"]);
      await context.sync();
      
      // Check if a range is selected
      if (range.cellCount === 0) {
        hideLoading();
        showMessage("Please select cells containing dates first.", "error");
        return;
      }
      
      // Get the Excel date format based on user selection
      const excelDateFormat = getExcelDateFormat(selectedFormat);
      
      // Get the current values
      const values = range.values;
      let dateCount = 0;
      
      // Process each cell
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
          const cellValue = values[i][j];
          
          // Skip empty cells
          if (cellValue === null || cellValue === undefined || cellValue === "") {
            continue;
          }
          
          // Check if the cell actually contains a date
          const dateValue = tryParseDate(cellValue);
          if (dateValue) {
            // Get a reference to this specific cell
            const cell = range.getCell(i, j);
            // Format the cell as a date with the chosen format
            cell.numberFormat = excelDateFormat;
            dateCount++;
          }
        }
      }
      
      await context.sync();
      
      if (dateCount > 0) {
        hideLoading();
        showMessage(`Success! Standardized ${dateCount} dates to ${selectedFormat} format.`, "success");
      } else {
        hideLoading();
        showMessage("No valid dates found in the selected range. Try selecting cells that contain dates.", "info");
      }
    });
  } catch (error) {
    hideLoading();
    showError(error);
  }
}

/**
 * Convert user-friendly date format to Excel number format
 * @param {string} format - User-friendly format
 * @returns {string} Excel number format
 */
export function getExcelDateFormat(format) {
  switch (format) {
    case "MM/DD/YYYY":
      return "mm/dd/yyyy";
    case "DD/MM/YYYY":
      return "dd/mm/yyyy";
    case "YYYY-MM-DD":
      return "yyyy-mm-dd";
    default:
      return "mm/dd/yyyy";
  }
}

/**
 * Try to parse a value as a date
 * @param {any} value - The value to try to parse
 * @returns {Date|null} A Date object if parsing was successful, null otherwise
 */
export function tryParseDate(value) {
  // If it's already a Date object
  if (value instanceof Date && !isNaN(value)) {
    return value;
  }
  
  // If it's a number, it might be an Excel date serial number
  if (typeof value === "number") {
    try {
      // Excel dates are stored as days since 1/1/1900
      // We need to adjust for Excel's bug where it thinks 1900 was a leap year
      const date = new Date((value - 1) * 86400000);
      if (!isNaN(date) && date.getFullYear() > 1900 && date.getFullYear() < 2100) {
        return date;
      }
    } catch (e) {
      console.error("Error parsing Excel date serial number:", e);
    }
  }
  
  // If it's a string, try to parse it
  if (typeof value === "string") {
    // Remove extra spaces
    const trimmed = value.trim();
    
    // Try parsing with Date constructor
    try {
      const date = new Date(trimmed);
      if (!isNaN(date) && date.getFullYear() > 1900 && date.getFullYear() < 2100) {
        return date;
      }
    } catch (e) {
      console.error("Error parsing date string with Date constructor:", e);
    }
    
    // Try common date formats with regex
    const formats = [
      // MM/DD/YYYY
      { regex: /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, fn: (m) => new Date(parseInt(m[3]), parseInt(m[1]) - 1, parseInt(m[2])) },
      // DD/MM/YYYY
      { regex: /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, fn: (m) => new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])) },
      // YYYY-MM-DD
      { regex: /^(\d{4})-(\d{1,2})-(\d{1,2})$/, fn: (m) => new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3])) },
      // MM-DD-YYYY
      { regex: /^(\d{1,2})-(\d{1,2})-(\d{4})$/, fn: (m) => new Date(parseInt(m[3]), parseInt(m[1]) - 1, parseInt(m[2])) },
      // DD-MM-YYYY
      { regex: /^(\d{1,2})-(\d{1,2})-(\d{4})$/, fn: (m) => new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])) },
      // Short date formats
      { regex: /^(\d{1,2})\/(\d{1,2})\/(\d{2})$/, fn: (m) => {
          // Interpret yy as 19yy or 20yy
          const year = parseInt(m[3]);
          const fullYear = year < 50 ? 2000 + year : 1900 + year;
          return new Date(fullYear, parseInt(m[1]) - 1, parseInt(m[2]));
        }
      },
      { regex: /^(\d{1,2})-(\d{1,2})-(\d{2})$/, fn: (m) => {
          // Interpret yy as 19yy or 20yy
          const year = parseInt(m[3]);
          const fullYear = year < 50 ? 2000 + year : 1900 + year;
          return new Date(fullYear, parseInt(m[1]) - 1, parseInt(m[2]));
        }
      }
    ];
    
    for (const format of formats) {
      const match = trimmed.match(format.regex);
      if (match) {
        try {
          const date = format.fn(match);
          if (!isNaN(date) && date.getFullYear() > 1900 && date.getFullYear() < 2100) {
            return date;
          }
        } catch (e) {
          console.error("Error parsing date with regex pattern:", e);
        }
      }
    }
  }
  
  return null;
}

/**
 * Format a date according to the specified format
 * @param {Date} date - The date to format
 * @param {string} format - The format to use
 * @returns {string} The formatted date
 */
export function formatDate(date, format) {
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();
  
  switch (format) {
    case "MM/DD/YYYY":
      return `${month}/${day}/${year}`;
    case "DD/MM/YYYY":
      return `${day}/${month}/${year}`;
    case "YYYY-MM-DD":
      return `${year}-${month}-${day}`;
    default:
      return `${month}/${day}/${year}`;
  }
}