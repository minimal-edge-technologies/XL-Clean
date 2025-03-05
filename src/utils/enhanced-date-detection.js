/*
 * Excel Data Cleaner Add-in
 * Enhanced Date Format Detection
 */

/**
 * Comprehensive date format detection for international formats
 * @param {any} value - The value to check
 * @returns {Object|null} Date information object or null if not a date
 */
export function detectDateFormat(value) {
    // If it's already a Date object
    if (value instanceof Date && !isNaN(value)) {
      return {
        date: value,
        format: "object",
        formatName: "Date Object"
      };
    }
    
    // If it's a number that might be an Excel date serial number
    if (typeof value === "number" && value > 1000 && value < 50000) {
      try {
        // Excel dates are stored as days since 1/1/1900
        // We need to adjust for Excel's bug where it thinks 1900 was a leap year
        const date = new Date((value - 1) * 86400000);
        if (!isNaN(date) && date.getFullYear() > 1900 && date.getFullYear() < 2100) {
          return {
            date,
            format: "excel",
            formatName: "Excel Serial Number"
          };
        }
      } catch (e) {
        console.error("Error parsing Excel date serial number:", e);
      }
    }
    
    // If it's a string, try multiple formats
    if (typeof value === "string" && value.trim().length > 0) {
      const trimmed = value.trim();
      
      // Create a collection of date format patterns with parsing functions
      const datePatterns = [
        // ISO 8601 formats (YYYY-MM-DD, YYYY/MM/DD)
        {
          regex: /^(\d{4})[\-\/](\d{1,2})[\-\/](\d{1,2})$/,
          parse: (m) => new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3])),
          format: "yyyy-mm-dd",
          formatName: "ISO 8601 (Year-Month-Day)"
        },
        
        // US format (MM/DD/YYYY, MM-DD-YYYY)
        {
          regex: /^(\d{1,2})[\-\/](\d{1,2})[\-\/](\d{4})$/,
          parse: (m) => new Date(parseInt(m[3]), parseInt(m[1]) - 1, parseInt(m[2])),
          format: "mm/dd/yyyy",
          formatName: "US Format (Month/Day/Year)"
        },
        
        // European format (DD/MM/YYYY, DD.MM.YYYY)
        {
          regex: /^(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{4})$/,
          parse: (m) => new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])),
          format: "dd/mm/yyyy",
          formatName: "European Format (Day/Month/Year)"
        },
        
        // Short year formats (MM/DD/YY, DD/MM/YY)
        {
          regex: /^(\d{1,2})[\-\/](\d{1,2})[\-\/](\d{2})$/,
          parse: (m) => {
            const year = parseInt(m[3]);
            const fullYear = year < 50 ? 2000 + year : 1900 + year;
            return new Date(fullYear, parseInt(m[1]) - 1, parseInt(m[2]));
          },
          format: "mm/dd/yy",
          formatName: "Short US Format (Month/Day/Year)"
        },
        
        // European short year format (DD/MM/YY)
        {
          regex: /^(\d{1,2})[\-\/\.](\d{1,2})[\-\/\.](\d{2})$/,
          parse: (m) => {
            const year = parseInt(m[3]);
            const fullYear = year < 50 ? 2000 + year : 1900 + year;
            return new Date(fullYear, parseInt(m[2]) - 1, parseInt(m[1]));
          },
          format: "dd/mm/yy",
          formatName: "Short European Format (Day/Month/Year)"
        },
        
        // Month name formats (13-Jan-2021, 13 January 2021)
        {
          regex: /^(\d{1,2})[\-\s]([a-zA-Z]{3,9})[\-\s](\d{4})$/,
          parse: (m) => {
            const monthNames = ["january", "february", "march", "april", "may", "june", 
                               "july", "august", "september", "october", "november", "december"];
            const shortMonthNames = ["jan", "feb", "mar", "apr", "may", "jun", 
                                    "jul", "aug", "sep", "oct", "nov", "dec"];
            
            let monthIndex = -1;
            const monthName = m[2].toLowerCase();
            
            // Try to match full month name first
            monthIndex = monthNames.indexOf(monthName);
            
            // If that fails, try short month name
            if (monthIndex === -1) {
              monthIndex = shortMonthNames.indexOf(monthName);
            }
            
            // If that fails, try to match just the first 3 letters
            if (monthIndex === -1 && monthName.length >= 3) {
              monthIndex = shortMonthNames.indexOf(monthName.substring(0, 3));
            }
            
            if (monthIndex !== -1) {
              return new Date(parseInt(m[3]), monthIndex, parseInt(m[1]));
            }
            
            return null;
          },
          format: "dd-mmm-yyyy",
          formatName: "Day-Month Name-Year"
        },
        
        // Year first with month name (2021-Jan-13, 2021 January 13)
        {
          regex: /^(\d{4})[\-\s]([a-zA-Z]{3,9})[\-\s](\d{1,2})$/,
          parse: (m) => {
            const monthNames = ["january", "february", "march", "april", "may", "june", 
                               "july", "august", "september", "october", "november", "december"];
            const shortMonthNames = ["jan", "feb", "mar", "apr", "may", "jun", 
                                    "jul", "aug", "sep", "oct", "nov", "dec"];
            
            let monthIndex = -1;
            const monthName = m[2].toLowerCase();
            
            // Try to match full month name first
            monthIndex = monthNames.indexOf(monthName);
            
            // If that fails, try short month name
            if (monthIndex === -1) {
              monthIndex = shortMonthNames.indexOf(monthName);
            }
            
            // If that fails, try to match just the first 3 letters
            if (monthIndex === -1 && monthName.length >= 3) {
              monthIndex = shortMonthNames.indexOf(monthName.substring(0, 3));
            }
            
            if (monthIndex !== -1) {
              return new Date(parseInt(m[1]), monthIndex, parseInt(m[3]));
            }
            
            return null;
          },
          format: "yyyy-mmm-dd",
          formatName: "Year-Month Name-Day"
        },
        
        // Other international formats
        
        // Chinese/Japanese/Korean (YYYY年MM月DD日)
        {
          regex: /^(\d{4})[年\s](\d{1,2})[月\s](\d{1,2})[日\s]?$/,
          parse: (m) => new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3])),
          format: "yyyy年mm月dd日",
          formatName: "East Asian Format (Year Month Day)"
        },
        
        // Date with time component (YYYY-MM-DD HH:MM:SS)
        {
          regex: /^(\d{4})[\-\/](\d{1,2})[\-\/](\d{1,2})[T\s](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/,
          parse: (m) => new Date(
            parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]), 
            parseInt(m[4]), parseInt(m[5]), m[6] ? parseInt(m[6]) : 0
          ),
          format: "yyyy-mm-dd hh:mm:ss",
          formatName: "ISO 8601 with Time"
        }
      ];
      
      // Try to match each pattern
      for (const pattern of datePatterns) {
        const match = trimmed.match(pattern.regex);
        if (match) {
          try {
            const date = pattern.parse(match);
            if (date && !isNaN(date) && date.getFullYear() > 1900 && date.getFullYear() < 2100) {
              return {
                date,
                format: pattern.format,
                formatName: pattern.formatName
              };
            }
          } catch (e) {
            console.error(`Error parsing date with pattern ${pattern.format}:`, e);
          }
        }
      }
      
      // Last resort: try the built-in Date parser
      try {
        const date = new Date(trimmed);
        if (!isNaN(date) && date.getFullYear() > 1900 && date.getFullYear() < 2100) {
          return {
            date,
            format: "auto",
            formatName: "Automatically Detected"
          };
        }
      } catch (e) {
        console.error("Error parsing date with built-in parser:", e);
      }
    }
    
    // Not a recognized date
    return null;
  }
  
  /**
   * Format a date according to a specified format
   * @param {Date} date - The date to format
   * @param {string} format - The format to use (e.g., "MM/DD/YYYY")
   * @returns {string} The formatted date string
   */
  export function formatDate(date, format) {
    if (!date || isNaN(date)) {
      return "";
    }
    
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    const shortYear = String(year).substring(2);
    
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    
    const monthNames = ["January", "February", "March", "April", "May", "June",
                        "July", "August", "September", "October", "November", "December"];
    const shortMonthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    
    switch (format.toUpperCase()) {
      case "MM/DD/YYYY":
        return `${month}/${day}/${year}`;
      case "DD/MM/YYYY":
        return `${day}/${month}/${year}`;
      case "YYYY-MM-DD":
        return `${year}-${month}-${day}`;
      case "MM/DD/YY":
        return `${month}/${day}/${shortYear}`;
      case "DD/MM/YY":
        return `${day}/${month}/${shortYear}`;
      case "DD-MMM-YYYY":
        return `${day}-${shortMonthNames[date.getMonth()]}-${year}`;
      case "YYYY-MMM-DD":
        return `${year}-${shortMonthNames[date.getMonth()]}-${day}`;
      case "DD MMMM YYYY":
        return `${day} ${monthNames[date.getMonth()]} ${year}`;
      case "YYYY年MM月DD日":
        return `${year}年${month}月${day}日`;
      case "YYYY-MM-DD HH:MM:SS":
        return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
      default:
        return `${month}/${day}/${year}`;
    }
  }
  
  /**
   * Get Excel date format string from user-friendly format
   * @param {string} userFormat - User-friendly format string (e.g., "MM/DD/YYYY")
   * @returns {string} Excel format string
   */
  export function getExcelDateFormat(userFormat) {
    const formatMap = {
      "MM/DD/YYYY": "mm/dd/yyyy",
      "DD/MM/YYYY": "dd/mm/yyyy",
      "YYYY-MM-DD": "yyyy-mm-dd",
      "MM/DD/YY": "mm/dd/yy",
      "DD/MM/YY": "dd/mm/yy",
      "DD-MMM-YYYY": "dd-mmm-yyyy",
      "YYYY-MMM-DD": "yyyy-mmm-dd",
      "DD MMMM YYYY": "dd mmmm yyyy",
      "YYYY年MM月DD日": "yyyy\"年\"mm\"月\"dd\"日\"",
      "YYYY-MM-DD HH:MM:SS": "yyyy-mm-dd hh:mm:ss"
    };
    
    return formatMap[userFormat] || "mm/dd/yyyy";
  }
  
  /**
   * Detect the most likely format of dates in a range
   * @param {Excel.Range} range - Range to analyze
   * @returns {Promise<string>} The most common date format
   */
  export async function detectMostCommonDateFormat(context, range) {
    // Load range values
    range.load("values");
    await context.sync();
    
    const values = range.values;
    const detectedFormats = {};
    
    // Analyze each cell
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        const value = values[i][j];
        const dateInfo = detectDateFormat(value);
        
        if (dateInfo) {
          detectedFormats[dateInfo.format] = (detectedFormats[dateInfo.format] || 0) + 1;
        }
      }
    }
    
    // Find the most common format
    let mostCommonFormat = "MM/DD/YYYY"; // Default
    let maxCount = 0;
    
    for (const format in detectedFormats) {
      if (detectedFormats[format] > maxCount) {
        maxCount = detectedFormats[format];
        mostCommonFormat = format;
      }
    }
    
    // Map internal format to user-friendly format
    const formatMapping = {
      "yyyy-mm-dd": "YYYY-MM-DD",
      "mm/dd/yyyy": "MM/DD/YYYY",
      "dd/mm/yyyy": "DD/MM/YYYY",
      "mm/dd/yy": "MM/DD/YY",
      "dd/mm/yy": "DD/MM/YY",
      "dd-mmm-yyyy": "DD-MMM-YYYY",
      "yyyy-mmm-dd": "YYYY-MMM-DD"
    };
    
    return formatMapping[mostCommonFormat] || mostCommonFormat;
  }