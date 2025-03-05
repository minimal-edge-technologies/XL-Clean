/*
 * Performance optimization for large data sets
 * This utility processes data in chunks to prevent UI freezing
 */

/**
 * Process a large range in chunks
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Range} range - The range to process
 * @param {Function} processChunk - Function to process each chunk (async)
 * @param {Object} options - Processing options
 * @param {number} options.chunkSize - Number of rows to process in each chunk (default: 1000)
 * @param {Function} options.onProgress - Callback for progress updates (optional)
 * @returns {Promise<Object>} Results of the processing
 */
export async function processRangeInChunks(context, range, processChunk, options = {}) {
    // Default options
    const chunkSize = options.chunkSize || 1000;
    const onProgress = options.onProgress || (() => {});
    
    // Load range properties
    range.load(["rowCount", "columnCount", "address"]);
    await context.sync();
    
    const totalRows = range.rowCount;
    let processedRows = 0;
    const results = { processedRows: 0, changedItems: 0 };
    
    // If the range is small enough, process it all at once
    if (totalRows <= chunkSize) {
      const chunkResult = await processChunk(range);
      Object.assign(results, chunkResult);
      results.processedRows = totalRows;
      return results;
    }
    
    // Get the worksheet and the starting cell address
    const worksheet = range.worksheet;
    const startAddress = range.address.split(':')[0];
    
    // Process in chunks
    for (let rowStart = 0; rowStart < totalRows; rowStart += chunkSize) {
      // Calculate the current chunk size (might be smaller for the last chunk)
      const currentChunkSize = Math.min(chunkSize, totalRows - rowStart);
      
      // Get the range for the current chunk
      const chunkStartCell = worksheet.getRange(startAddress).getOffsetRange(rowStart, 0);
      const chunkRange = chunkStartCell.getResizedRange(currentChunkSize - 1, range.columnCount - 1);
      
      // Process this chunk
      const chunkResult = await processChunk(chunkRange);
      
      // Update results
      results.changedItems += chunkResult.changedItems || 0;
      processedRows += currentChunkSize;
      results.processedRows = processedRows;
      
      // Report progress
      onProgress({
        processedRows,
        totalRows,
        percentComplete: Math.round((processedRows / totalRows) * 100),
        changedItems: results.changedItems
      });
      
      // Allow UI to update by yielding execution briefly
      await new Promise(resolve => setTimeout(resolve, 0));
    }
    
    return results;
  }
  
  /**
   * Example usage for trimming spaces in large ranges
   * @param {Excel.RequestContext} context - Excel context
   * @param {Excel.Range} range - Range to process
   * @returns {Promise<Object>} Processing results
   */
  export async function optimizedTrimSpaces(context, range) {
    // Function to process each chunk
    const processChunk = async (chunkRange) => {
      chunkRange.load("values");
      await context.sync();
      
      const values = chunkRange.values;
      let changedItems = 0;
      
      // Process values in this chunk
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
          if (typeof values[i][j] === "string") {
            const original = values[i][j];
            values[i][j] = original.trim().replace(/\s+/g, ' ');
            
            if (values[i][j] !== original) {
              changedItems++;
            }
          }
        }
      }
      
      // Update the range with processed values
      chunkRange.values = values;
      await context.sync();
      
      return { changedItems };
    };
    
    // Progress indicator update function
    const updateProgress = (progress) => {
      const loadingText = document.getElementById("loading-overlay").querySelector(".loading-text");
      if (loadingText) {
        loadingText.textContent = `Trimming spaces... ${progress.percentComplete}% complete`;
      }
    };
    
    // Process the range in chunks
    return processRangeInChunks(context, range, processChunk, {
      chunkSize: 1000,
      onProgress: updateProgress
    });
  }
  
  /**
   * Example usage for finding and replacing in large ranges
   * @param {Excel.RequestContext} context - Excel context
   * @param {Excel.Range} range - Range to process
   * @param {string} findText - Text to find
   * @param {string} replaceText - Replacement text
   * @returns {Promise<Object>} Processing results
   */
  export async function optimizedFindAndReplace(context, range, findText, replaceText) {
    // Create a regex for the find operation
    const regex = new RegExp(escapeRegExp(findText), "g");
    
    // Function to process each chunk
    const processChunk = async (chunkRange) => {
      chunkRange.load("values");
      await context.sync();
      
      const values = chunkRange.values;
      let changedItems = 0;
      let replacementCount = 0;
      
      // Process values in this chunk
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
          if (typeof values[i][j] === "string") {
            const original = values[i][j];
            
            try {
              const newText = original.replace(regex, replaceText);
              
              if (newText !== original) {
                const matchCount = (original.match(regex) || []).length;
                replacementCount += matchCount;
                changedItems++;
                values[i][j] = newText;
              }
            } catch (regexError) {
              // Fallback to simple replace if regex fails
              const newText = original.split(findText).join(replaceText);
              if (newText !== original) {
                const matchCount = (original.split(findText).length - 1);
                replacementCount += matchCount;
                changedItems++;
                values[i][j] = newText;
              }
            }
          }
        }
      }
      
      // Update the range with processed values
      chunkRange.values = values;
      await context.sync();
      
      return { changedItems, replacementCount };
    };
    
    // Progress indicator update function
    const updateProgress = (progress) => {
      const loadingText = document.getElementById("loading-overlay").querySelector(".loading-text");
      if (loadingText) {
        loadingText.textContent = `Replacing text... ${progress.percentComplete}% complete`;
      }
    };
    
    // Process the range in chunks
    return processRangeInChunks(context, range, processChunk, {
      chunkSize: 1000,
      onProgress: updateProgress
    });
  }
  
  /**
   * Escape regular expression special characters
   * @param {string} string - The string to escape
   * @returns {string} Escaped string
   */
  function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }