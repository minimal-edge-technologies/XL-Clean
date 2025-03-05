/*
 * Excel Data Cleaner Add-in
 * Data Preview Feature
 */

/**
 * Generate a preview of a data transformation
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Range} range - The range to preview
 * @param {Function} transformFunc - Function to transform a single value
 * @param {Object} options - Preview options
 * @param {number} options.maxRows - Maximum number of rows to preview (default: 5)
 * @param {number} options.maxCols - Maximum number of columns to preview (default: 3)
 * @returns {Promise<Object>} The preview data for before/after comparison
 */
export async function generatePreview(context, range, transformFunc, options = {}) {
    // Default options
    const maxRows = options.maxRows || 5;
    const maxCols = options.maxCols || 3;
    
    // Load the range properties
    range.load(["values", "rowCount", "columnCount", "address"]);
    await context.sync();
    
    // Determine the preview size
    const previewRows = Math.min(range.rowCount, maxRows);
    const previewCols = Math.min(range.columnCount, maxCols);
    
    // Create a smaller range for the preview
    let previewRange;
    if (previewRows < range.rowCount || previewCols < range.columnCount) {
      previewRange = range.getResizedRange(previewRows - 1, previewCols - 1);
      previewRange.load("values");
      await context.sync();
    } else {
      previewRange = range;
    }
    
    // Create before/after preview
    const before = previewRange.values;
    const after = [];
    
    // Apply the transformation to create the "after" preview
    let changedCells = 0;
    let totalCells = 0;
    
    for (let i = 0; i < before.length; i++) {
      after[i] = [];
      for (let j = 0; j < before[i].length; j++) {
        const original = before[i][j];
        const transformed = transformFunc(original);
        after[i][j] = transformed;
        
        if (transformed !== original) {
          changedCells++;
        }
        totalCells++;
      }
    }
    
    // Create preview metadata
    const previewMeta = {
      totalRange: {
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        address: range.address
      },
      previewRange: {
        rowCount: previewRows,
        columnCount: previewCols
      },
      hasMoreData: previewRows < range.rowCount || previewCols < range.columnCount,
      estimatedChanges: Math.round((changedCells / totalCells) * range.rowCount * range.columnCount)
    };
    
    return {
      before,
      after,
      meta: previewMeta
    };
  }

  // Add to data-preview.js
export async function generatePreviewForLargeData(context, range, transformFunc, options = {}) {
  // Default options
  const maxRows = options.maxRows || 5;
  const maxCols = options.maxCols || 3;
  const maxCellsToProcess = options.maxCellsToProcess || 10000;
  
  // Load the range properties
  range.load(["values", "rowCount", "columnCount", "address"]);
  await context.sync();
  
  // Check if we need to switch to sampling for large ranges
  const totalCells = range.rowCount * range.columnCount;
  
  if (totalCells > maxCellsToProcess) {
    // For large ranges, use sampling instead of processing everything
    return generateSampledPreview(context, range, transformFunc, options);
  } else {
    // For smaller ranges, use the standard preview
    return generatePreview(context, range, transformFunc, options);
  }
}

// Helper function to generate a sampled preview for large data sets
async function generateSampledPreview(context, range, transformFunc, options) {
  const maxRows = options.maxRows || 5;
  const maxCols = options.maxCols || 3;
  
  // Load range properties
  range.load(["rowCount", "columnCount", "address"]);
  await context.sync();
  
  // Calculate sample positions
  const rowPositions = [];
  const colPositions = [];
  
  // Add start, middle and end positions
  if (range.rowCount > 0) rowPositions.push(0);
  if (range.rowCount > 2) rowPositions.push(Math.floor(range.rowCount / 2));
  if (range.rowCount > 1) rowPositions.push(range.rowCount - 1);
  
  if (range.columnCount > 0) colPositions.push(0);
  if (range.columnCount > 2) colPositions.push(Math.floor(range.columnCount / 2));
  if (range.columnCount > 1) colPositions.push(range.columnCount - 1);
  
  // Get samples from the range
  const sampleCells = [];
  
  for (const row of rowPositions.slice(0, maxRows)) {
    for (const col of colPositions.slice(0, maxCols)) {
      const cell = range.getCell(row, col);
      cell.load(["values", "formulas"]);
      sampleCells.push({ cell, row, col });
    }
  }
  
  await context.sync();
  
  // Create before/after preview
  const before = Array(rowPositions.length).fill().map(() => Array(colPositions.length).fill(null));
  const after = Array(rowPositions.length).fill().map(() => Array(colPositions.length).fill(null));
  
  // Apply transformation to samples
  let changedCells = 0;
  
  sampleCells.forEach((sample, index) => {
    const rowIndex = rowPositions.indexOf(sample.row);
    const colIndex = colPositions.indexOf(sample.col);
    
    if (rowIndex >= 0 && colIndex >= 0) {
      const value = sample.cell.values[0][0];
      before[rowIndex][colIndex] = value;
      
      const transformed = transformFunc(value);
      after[rowIndex][colIndex] = transformed;
      
      if (transformed !== value) {
        changedCells++;
      }
    }
  });
  
  // Estimate changes based on the sample
  const estimatedChanges = Math.round((changedCells / sampleCells.length) * range.rowCount * range.columnCount);
  
  // Create preview metadata
  const previewMeta = {
    totalRange: {
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      address: range.address
    },
    previewRange: {
      rowCount: rowPositions.length,
      columnCount: colPositions.length
    },
    hasMoreData: true,
    isSampled: true,
    estimatedChanges: estimatedChanges
  };
  
  return {
    before,
    after,
    meta: previewMeta
  };
}
  
  /**
   * Show a preview dialog for data transformation
   * @param {Object} previewData - Preview data from generatePreview()
   * @param {string} title - Dialog title
   * @param {Function} onApply - Callback when user applies the transformation
   * @param {Function} onCancel - Callback when user cancels the transformation
   */
  export function showPreviewDialog(previewData, title, onApply, onCancel) {
    // Create dialog if it doesn't exist
    let previewDialog = document.getElementById("preview-dialog");
    if (!previewDialog) {
      previewDialog = document.createElement("div");
      previewDialog.id = "preview-dialog";
      previewDialog.className = "ms-Dialog";
      previewDialog.setAttribute("role", "dialog");
      previewDialog.setAttribute("aria-labelledby", "preview-dialog-title");
      
      document.body.appendChild(previewDialog);
    }
    
    // Clear any existing content
    previewDialog.innerHTML = '';
    
    // Set up the dialog content
    previewDialog.innerHTML = `
      <div class="ms-Dialog-main">
        <div class="ms-Dialog-title" id="preview-dialog-title">${title}</div>
        <div class="ms-Dialog-content">
          <p>Below is a preview of how your data will change:</p>
          <div class="preview-container">
            <div class="preview-before">
              <h3>Before</h3>
              <div class="preview-table-container" id="preview-before"></div>
            </div>
            <div class="preview-after">
              <h3>After</h3>
              <div class="preview-table-container" id="preview-after"></div>
            </div>
          </div>
          <div class="preview-meta">
            <p>
              Showing ${previewData.meta.previewRange.rowCount} × ${previewData.meta.previewRange.columnCount} 
              cells from a total of ${previewData.meta.totalRange.rowCount} × ${previewData.meta.totalRange.columnCount}
              ${previewData.meta.hasMoreData ? ' (preview only)' : ''}
            </p>
            <p>Estimated cells to be changed: <strong>${previewData.meta.estimatedChanges}</strong></p>
          </div>
        </div>
        <div class="ms-Dialog-actions">
          <button id="apply-preview" class="primary-button">
            <span>Apply Changes</span>
          </button>
          <button id="cancel-preview" class="secondary-button">
            <span>Cancel</span>
          </button>
        </div>
      </div>
    `;
    
    // Add the before/after tables
    const beforeContainer = document.getElementById("preview-before");
    const afterContainer = document.getElementById("preview-after");
    
    // Create the tables
    const renderTable = (container, data) => {
      const table = document.createElement("table");
      table.className = "preview-table";
      
      // Create table rows and cells
      for (let i = 0; i < data.length; i++) {
        const row = document.createElement("tr");
        
        for (let j = 0; j < data[i].length; j++) {
          const cell = document.createElement("td");
          // Handle different data types appropriately
          if (data[i][j] === null || data[i][j] === undefined) {
            cell.innerHTML = '<span class="empty-cell">(empty)</span>';
          } else {
            cell.textContent = String(data[i][j]);
          }
          
          // Highlight differences in the after table
          if (container === afterContainer && 
              JSON.stringify(data[i][j]) !== JSON.stringify(previewData.before[i][j])) {
            cell.className = "changed-cell";
          }
          
          row.appendChild(cell);
        }
        
        table.appendChild(row);
      }
      
      container.appendChild(table);
    };
    
    renderTable(beforeContainer, previewData.before);
    renderTable(afterContainer, previewData.after);
    
    // Set up event handlers
    document.getElementById("apply-preview").addEventListener("click", () => {
      previewDialog.style.display = "none";
      if (onApply) onApply();
    });
    
    document.getElementById("cancel-preview").addEventListener("click", () => {
      previewDialog.style.display = "none";
      if (onCancel) onCancel();
    });
    
    // Show the dialog
    previewDialog.style.display = "block";
    
    // Add some basic styles if they don't exist
    if (!document.getElementById("preview-dialog-styles")) {
      const styleEl = document.createElement("style");
      styleEl.id = "preview-dialog-styles";
      styleEl.textContent = `
        .preview-container {
          display: flex;
          gap: 20px;
          margin-bottom: 20px;
        }
        .preview-before, .preview-after {
          flex: 1;
        }
        .preview-table-container {
          max-height: 300px;
          overflow-y: auto;
          border: 1px solid #e0e0e0;
        }
        .preview-table {
          width: 100%;
          border-collapse: collapse;
        }
        .preview-table td {
          padding: 6px 8px;
          border: 1px solid #e0e0e0;
          max-width: 150px;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }
        .changed-cell {
          background-color: #e6f7ff;
          font-weight: 500;
          position: relative;
        }
        .empty-cell {
          color: #999;
          font-style: italic;
        }
      `;
      document.head.appendChild(styleEl);
    }
  }
  
  /**
   * Example usage with Find & Replace
   * @param {Excel.Range} range - Range to preview
   * @param {string} findText - Text to find
   * @param {string} replaceText - Text to replace with
   */
  export async function previewFindAndReplace(context, range, findText, replaceText) {
    // Create the transformation function
    const findReplaceTransform = (value) => {
      if (typeof value !== 'string') return value;
      
      try {
        // Create a RegExp for the find operation
        const regex = new RegExp(escapeRegExp(findText), "g");
        return value.replace(regex, replaceText);
      } catch (error) {
        // Fallback to simple replace if regex fails
        return value.split(findText).join(replaceText);
      }
    };
    
    // Generate the preview
    const previewData = await generatePreview(context, range, findReplaceTransform);
    
    return previewData;
  }
  
  /**
   * Example usage with Case Conversion
   * @param {Excel.Range} range - Range to preview
   * @param {string} caseType - Type of case conversion (UPPER, LOWER, PROPER)
   */
  export async function previewCaseOperation() {
    try {
      // Determine which case type is being previewed based on which button was clicked
      let caseType = "PROPER"; // Default
      const clickedButtonId = document.activeElement.id;
      
      if (clickedButtonId === "uppercase-button" || clickedButtonId.includes("uppercase")) {
        caseType = "UPPER";
      } else if (clickedButtonId === "lowercase-button" || clickedButtonId.includes("lowercase")) {
        caseType = "LOWER";
      }
      
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("cellCount");
        await context.sync();
        
        if (range.cellCount === 0) {
          showMessage("Please select cells containing text first.", "error");
          return;
        }
        
        const previewData = await previewCaseConversion(context, range, caseType);
        showPreviewDialog(previewData, `${caseType.charAt(0) + caseType.slice(1).toLowerCase()} Case Preview`, 
          () => convertCase(caseType), // onApply
          () => {} // onCancel
        );
      });
    } catch (error) {
      showError(error);
    }
  }
  
  /**
   * Example usage with Trim Spaces
   * @param {Excel.Range} range - Range to preview
   */
  export async function previewTrimSpaces(context, range) {
    // Create the transformation function
    const trimTransform = (value) => {
      if (typeof value !== 'string') return value;
      return value.trim().replace(/\s+/g, ' ');
    };
    
    // Generate the preview
    const previewData = await generatePreview(context, range, trimTransform);
    
    return previewData;
  }
  
  /**
   * Escape regular expression special characters
   * @param {string} string - The string to escape
   * @returns {string} Escaped string
   */
  function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  /**
 * Preview duplicate removal
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Range} range - Range to preview
 */
export async function previewDuplicateRemoval() {
  try {
    showLoading("Generating duplicate removal preview...");
    
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("cellCount");
      await context.sync();
      
      if (range.cellCount === 0) {
        hideLoading();
        showMessage("Please select a range with data first.", "error");
        return;
      }
      
      // Generate the preview data
      const previewData = await previewDuplicateRemoval(context, range);
      
      hideLoading();
      showPreviewDialog(previewData, "Remove Duplicates Preview", 
        () => removeDuplicates(), // onApply
        () => {} // onCancel
      );
    });
  } catch (error) {
    hideLoading();
    showError(error);
  }
}

/**
 * Preview date standardization
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Range} range - Range to preview
 */
export async function previewDateStandardization() {
  try {
    showLoading("Generating date format preview...");
    
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("cellCount");
      await context.sync();
      
      if (range.cellCount === 0) {
        hideLoading();
        showMessage("Please select cells containing dates first.", "error");
        return;
      }
      
      // Get the selected date format
      const formatRadios = document.getElementsByName("dateFormat");
      let selectedFormat = "MM/DD/YYYY"; // Default
      for (const radio of formatRadios) {
        if (radio.checked) {
          selectedFormat = radio.value;
          break;
        }
      }
      
      // Generate the preview data
      const previewData = await previewDateStandardization(context, range);
      
      hideLoading();
      showPreviewDialog(previewData, `Date Format Preview (${selectedFormat})`, 
        () => standardizeDates(), // onApply
        () => {} // onCancel
      );
    });
  } catch (error) {
    hideLoading();
    showError(error);
  }
}

/**
 * Preview one-click cleanup
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Range} range - Range to preview
 */
export async function previewOneClickCleanup() {
  try {
    showLoading("Generating one-click cleanup preview...");
    
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("cellCount");
      await context.sync();
      
      if (range.cellCount === 0) {
        hideLoading();
        showMessage("Please select a range with data first.", "error");
        return;
      }
      
      // Generate the preview data
      const previewData = await previewOneClickCleanup(context, range);
      
      // Build a description of what operations will be performed
      const operations = [];
      if (document.getElementById("cleanup-duplicates").checked) operations.push("remove duplicates");
      if (document.getElementById("cleanup-spaces").checked) operations.push("trim spaces");
      if (document.getElementById("cleanup-case").checked) operations.push("fix text case");
      if (document.getElementById("cleanup-formatting").checked) operations.push("fix number formatting");
      
      const operationsText = operations.join(", ");
      
      hideLoading();
      showPreviewDialog(previewData, `One-Click Cleanup Preview (${operationsText})`, 
        () => oneClickCleanup(), // onApply
        () => {} // onCancel
      );
    });
  } catch (error) {
    hideLoading();
    showError(error);
  }
}