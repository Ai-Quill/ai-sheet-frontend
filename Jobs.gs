/**
 * @file Jobs.gs
 * @version 2.2.0
 * @updated 2026-01-10
 * @author AISheeter Team
 * 
 * CHANGELOG:
 * - 2.2.0 (2026-01-10): Header reading & column inference for smart prompts
 * - 2.1.0 (2026-01-10): Best-in-class UX (auto-detect, preview, incremental, retry)
 * - 2.0.0 (2026-01-10): Initial bulk processing
 * 
 * ============================================
 * JOBS.gs - Best-in-Class Bulk Processing
 * ============================================
 * 
 * Handles bulk AI processing without timeouts.
 * Features:
 *   - Smart auto-detection of selected range
 *   - Preview mode (test first 3 rows)
 *   - Incremental writes (results appear live)
 *   - Resilient error handling (retry failed only)
 *   - Job persistence (survives sidebar close)
 * 
 * Usage:
 *   var selection = getSelectedRange();
 *   var preview = previewBulkJob(range, prompt, model);
 *   var job = createBulkJob(inputs, prompt, model);
 *   var status = getJobStatus(job.id);
 *   writeIncrementalResults(job.id, results, column, row);
 *   retryFailedRows(job.id);
 */

// ============================================
// JOB CREATION
// ============================================

/**
 * Create a bulk processing job
 * 
 * @param {Array<string>} inputData - Array of inputs to process
 * @param {string} prompt - Prompt template with {{input}} placeholder
 * @param {string} model - Provider (CHATGPT, CLAUDE, GROQ, GEMINI)
 * @param {string} [specificModel] - Optional specific model ID
 * @return {Object} Job info with id, status, totalRows
 * @throws {Error} If job creation fails
 */
function createBulkJob(inputData, prompt, model, specificModel) {
  Logger.log('[createBulkJob] === CALLED ===');
  Logger.log('[createBulkJob] inputData length: ' + (inputData ? inputData.length : 'null'));
  Logger.log('[createBulkJob] inputData type: ' + (Array.isArray(inputData) ? 'array' : typeof inputData));
  Logger.log('[createBulkJob] prompt length: ' + (prompt ? prompt.length : 'null'));
  Logger.log('[createBulkJob] model: ' + model);
  
  // Validate inputs first (before any API calls)
  if (!inputData || !Array.isArray(inputData) || inputData.length === 0) {
    Logger.log('[createBulkJob] ❌ Invalid inputData!');
    throw new Error('Input data must be a non-empty array');
  }
  
  if (!prompt || typeof prompt !== 'string' || prompt.trim() === '') {
    throw new Error('Prompt must be a non-empty string');
  }
  
  // Filter out empty values from input data
  inputData = inputData.filter(function(item) {
    return item !== null && item !== undefined && String(item).trim() !== '';
  });
  
  if (inputData.length === 0) {
    throw new Error('All input values are empty');
  }
  
  if (!Config.getSupportedProviders().includes(model)) {
    throw new Error('Unsupported model: ' + model);
  }
  
  var userId = getUserId();
  
  // Use default model if not specified
  if (!specificModel) {
    specificModel = Config.getDefaultModel(model);
  }
  
  // Get encrypted API key using SecureRequest (centralized handling)
  var encryptedApiKey = SecureRequest.getEncryptedKey(model);
  
  var payload = {
    userId: userId,
    inputs: inputData,  // Backend API expects 'inputs' field
    config: {
      prompt: prompt,
      model: model,
      specificModel: specificModel,
      useBYOK: true,
      encryptedApiKey: encryptedApiKey
    }
  };
  
  // Log detailed job creation info for debugging
  Logger.log('📤 Creating bulk job:');
  Logger.log('  Rows: ' + inputData.length);
  Logger.log('  Prompt (first 200 chars): ' + prompt.substring(0, 200) + (prompt.length > 200 ? '...' : ''));
  if (inputData.length > 0) {
    var firstInputPreview = String(inputData[0]).substring(0, 150).replace(/\n/g, ' | ');
    Logger.log('  First input: "' + firstInputPreview + '"');
  }
  
  try {
    var result = ApiClient.post('JOBS', payload);
    
    // Backend returns 'jobId', not 'id'
    var jobId = result.jobId || result.id;
    
    if (Config.isDebug()) {
      Logger.log('Bulk job created: ' + jobId);
    }
    
    if (!jobId) {
      throw new Error('No job ID returned from server');
    }
    
    return {
      id: jobId,
      status: result.status || 'pending',
      totalRows: inputData.length,
      createdAt: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('Error creating bulk job:', error);
    throw new Error('Failed to create bulk job: ' + error.message);
  }
}

// ============================================
// JOB STATUS & RESULTS
// ============================================

/**
 * Get status and results of a bulk job
 * 
 * @param {string} jobId - The job ID
 * @return {Object} Job status with progress, results, errors
 */
function getJobStatus(jobId) {
  var userId = getUserId();
  
  try {
    var result = ApiClient.get('JOBS', { id: jobId, userId: userId });
    
    return {
      id: result.id,
      status: result.status,
      progress: result.progress || 0,
      processedRows: result.processedRows || 0,
      totalRows: result.totalRows || 0,
      results: result.results || [],
      errors: result.errors || [],
      createdAt: result.createdAt,
      completedAt: result.completedAt
    };
    
  } catch (error) {
    console.error('Error getting job status:', error);
    throw new Error('Failed to get job status: ' + error.message);
  }
}

/**
 * Get all jobs for current user
 * 
 * @param {number} [limit] - Max jobs to return (default 10)
 * @return {Array<Object>} List of jobs
 */
function getJobHistory(limit) {
  var userId = getUserId();
  
  try {
    var params = { userId: userId };
    if (limit) params.limit = limit;
    
    var result = ApiClient.get('JOBS', params);
    return result.jobs || result;
    
  } catch (error) {
    console.error('Error getting job history:', error);
    throw new Error('Failed to get job history: ' + error.message);
  }
}

/**
 * Cancel a running bulk job
 * 
 * @param {string} jobId - The job ID
 * @return {Object} Cancellation result
 */
function cancelJob(jobId) {
  var userId = getUserId();
  
  try {
    var result = ApiClient.delete('JOBS', { id: jobId, userId: userId });
    
    if (Config.isDebug()) {
      Logger.log('Job cancelled: ' + jobId);
    }
    
    return {
      success: true,
      message: result.message || 'Job cancelled'
    };
    
  } catch (error) {
    console.error('Error cancelling job:', error);
    throw new Error('Failed to cancel job: ' + error.message);
  }
}

// ============================================
// SHEET HELPERS
// ============================================

/**
 * Get values from a spreadsheet range
 * 
 * @param {string} rangeNotation - Range in A1 notation (e.g., "A2:A100")
 * @return {Array<string>} Array of cell values (non-empty)
 */
function getRangeValues(rangeNotation) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getRange(rangeNotation);
    var values = range.getValues();
    
    // Flatten and filter empty values
    return values
      .flat()
      .filter(function(v) { 
        return v !== '' && v !== null && v !== undefined; 
      })
      .map(function(v) {
        return String(v);
      });
      
  } catch (error) {
    throw new Error('Invalid range: ' + rangeNotation + ' - ' + error.message);
  }
}

/**
 * Get range values as 2D array (preserves multi-column structure)
 * Used for pattern analysis to understand data relationships
 * 
 * @param {string} rangeNotation - Range like "B6:D15"
 * @return {Array<Array>} 2D array of values
 */
function getRangeValuesMultiColumn(rangeNotation) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getRange(rangeNotation);
    var values = range.getValues();
    
    // Return 2D array, filtering out completely empty rows
    return values.filter(function(row) {
      return row.some(function(cell) {
        return cell !== '' && cell !== null && cell !== undefined;
      });
    });
    
  } catch (error) {
    throw new Error('Invalid range: ' + rangeNotation + ' - ' + error.message);
  }
}

/**
 * Get input data from multiple columns, structured for AI processing
 * Each row returns a formatted string with labeled values from all input columns
 * 
 * @param {string} inputRange - Range notation (e.g., "A6:C10")
 * @param {Array<string>} inputColumns - Array of column letters (e.g., ["A", "B", "C"])
 * @param {Object} columnHeaders - Map of column letter to header name (e.g., {A: "Spare parts", B: "Car Brand"})
 * @return {Array<string>} Array of structured input strings, one per row
 */
function getMultiColumnInputData(inputRange, inputColumns, columnHeaders) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    
    // Parse the input range to get start and end rows
    var rangeMatch = inputRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
    if (!rangeMatch) {
      throw new Error('Invalid range format: ' + inputRange);
    }
    
    var startRow = parseInt(rangeMatch[2]);
    var endRow = parseInt(rangeMatch[4]);
    var rowCount = endRow - startRow + 1;
    
    Logger.log('📊 Multi-column input extraction:');
    Logger.log('  Range: ' + inputRange + ' (rows ' + startRow + '-' + endRow + ')');
    Logger.log('  Columns: ' + inputColumns.join(', '));
    Logger.log('  Headers: ' + JSON.stringify(columnHeaders));
    
    // Get data from each input column
    var columnData = {};
    inputColumns.forEach(function(col) {
      var colRange = col + startRow + ':' + col + endRow;
      var values = sheet.getRange(colRange).getValues();
      columnData[col] = values.map(function(row) { return String(row[0] || '').trim(); });
      Logger.log('  Column ' + col + ' values: ' + JSON.stringify(columnData[col].slice(0, 3)) + (columnData[col].length > 3 ? '...' : ''));
    });
    
    // Build structured input for each row
    var structuredInputs = [];
    for (var rowIdx = 0; rowIdx < rowCount; rowIdx++) {
      var parts = [];
      var hasData = false;
      
      inputColumns.forEach(function(col) {
        var value = columnData[col][rowIdx];
        if (value) {
          hasData = true;
        }
        var header = columnHeaders[col] || ('Column ' + col);
        parts.push(header + ': ' + value);
      });
      
      if (hasData) {
        // Create a structured string that the AI can parse
        var structuredRow = parts.join('\n');
        structuredInputs.push(structuredRow);
        
        // Log each structured row for debugging
        if (rowIdx < 3) {
          Logger.log('  Row ' + rowIdx + ' structured input:\n    ' + structuredRow.replace(/\n/g, '\n    '));
        }
      } else {
        Logger.log('  Row ' + rowIdx + ' SKIPPED (no data)');
      }
    }
    
    Logger.log('📊 Generated ' + structuredInputs.length + ' structured inputs from ' + rowCount + ' rows');
    
    return structuredInputs;
    
  } catch (error) {
    Logger.log('getMultiColumnInputData error: ' + error.message);
    throw new Error('Failed to get multi-column input: ' + error.message);
  }
}

/**
 * Write job results to a column using batch writing for performance
 * Now handles data validation errors gracefully by falling back to cell-by-cell writes
 * 
 * @param {Array<Object>} results - Array of {index, output, error} objects
 * @param {string} outputColumn - Column letter (e.g., "B")
 * @param {number} [startRow] - Starting row number (default 2)
 * @return {Object} Write result with successCount, errorCount, errors
 */
function writeJobResults(results, outputColumn, startRow) {
  startRow = startRow || 2;
  
  if (!results || results.length === 0) {
    Logger.log('writeJobResults: No results to write for column ' + outputColumn);
    return { successCount: 0, errorCount: 0, errors: [] };
  }
  
  // Debug: log results structure with actual values
  Logger.log('📝 writeJobResults: Writing ' + results.length + ' results to column ' + outputColumn + ' starting at row ' + startRow);
  Logger.log('First 3 values:');
  results.slice(0, 3).forEach(function(r) {
    var preview = (r.output || '').toString().substring(0, 60);
    Logger.log('  [' + r.index + '] → row ' + (startRow + r.index) + ': "' + preview + (preview.length < (r.output || '').toString().length ? '...' : '') + '"');
  });
  
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Sort by index to ensure correct order
  results.sort(function(a, b) { return a.index - b.index; });
  
  // Find min and max indices to determine range
  var minIndex = results[0].index;
  var maxIndex = results[results.length - 1].index;
  var rangeLength = maxIndex - minIndex + 1;
  
  // Check if indices are contiguous (no gaps)
  var isContiguous = results.length === rangeLength;
  
  // Try batch write first (fast path)
  // Clear validation first to avoid silent rejection of values that don't match dropdowns
  try {
    var range = sheet.getRange(outputColumn + (startRow + minIndex) + ':' + outputColumn + (startRow + maxIndex));
    
    // Store existing validation to potentially restore later (optional)
    var existingValidation = null;
    try {
      existingValidation = range.getDataValidation();
    } catch (e) { /* no validation */ }
    
    // Clear validation to ensure all values are written
    if (existingValidation) {
      Logger.log('Clearing data validation on ' + outputColumn + ' before writing');
      range.clearDataValidations();
    }
    
    if (isContiguous) {
      var values = results.map(function(result) {
        return [result.output || result.error || ''];
      });
      
      range.setValues(values);
      
    } else {
      // range was already defined above
      var existingValues = range.getValues();
      
      results.forEach(function(result) {
        var arrayIndex = result.index - minIndex;
        existingValues[arrayIndex][0] = result.output || result.error || '';
      });
      
      range.setValues(existingValues);
    }
    
    SpreadsheetApp.flush();
    
    if (Config.isDebug()) {
      Logger.log('Wrote ' + results.length + ' results to column ' + outputColumn + ' (batch mode)');
    }
    
    return { successCount: results.length, errorCount: 0, errors: [] };
    
  } catch (batchError) {
    // Batch write failed (likely data validation) - fall back to cell-by-cell with validation bypass
    Logger.log('Batch write failed, falling back to cell-by-cell with validation bypass: ' + batchError.message);
    
    var successCount = 0;
    var errorCount = 0;
    var errors = [];
    
    results.forEach(function(result) {
      var row = startRow + result.index;
      var value = result.output || result.error || '';
      
      try {
        var cell = sheet.getRange(outputColumn + row);
        cell.setValue(value);
        successCount++;
      } catch (cellError) {
        // First failure - try to bypass data validation
        Logger.log('Cell ' + outputColumn + row + ' failed, attempting validation bypass: ' + cellError.message);
        
        try {
          var cell = sheet.getRange(outputColumn + row);
          
          // Store existing validation (if any)
          var existingValidation = null;
          try {
            existingValidation = cell.getDataValidation();
          } catch (e) {
            // No validation or can't read it
          }
          
          // Clear validation, write value, then restore validation
          cell.clearDataValidations();
          cell.setValue(value);
          
          // Optionally restore validation (commented out - let user decide)
          // if (existingValidation) {
          //   cell.setDataValidation(existingValidation);
          // }
          
          successCount++;
          Logger.log('Successfully wrote cell ' + outputColumn + row + ' after bypassing validation');
        } catch (bypassError) {
          // Even with bypass, still failed
          errorCount++;
          errors.push({
            row: row,
            column: outputColumn,
            value: value,
            error: bypassError.message
          });
          Logger.log('Failed to write cell ' + outputColumn + row + ' even with bypass: ' + bypassError.message);
        }
      }
    });
    
    // Flush what we could write
    try {
      SpreadsheetApp.flush();
    } catch (e) {
      // Ignore flush errors
    }
    
    Logger.log('Cell-by-cell write complete: ' + successCount + ' success, ' + errorCount + ' errors');
    
    // If ALL cells failed, throw an error so the UI shows something went wrong
    if (successCount === 0 && errorCount > 0) {
      throw new Error('Failed to write results: ' + errors[0].error);
    }
    
    return { successCount: successCount, errorCount: errorCount, errors: errors };
  }
}

/**
 * Get the number of non-empty rows in a range
 * 
 * @param {string} rangeNotation - Range in A1 notation
 * @return {number} Count of non-empty cells
 */
function countNonEmptyRows(rangeNotation) {
  return getRangeValues(rangeNotation).length;
}

// ============================================
// COST ESTIMATION
// ============================================

/**
 * Estimate cost for a bulk job
 * Based on average tokens per request
 * 
 * @param {number} rowCount - Number of rows to process
 * @param {string} model - Provider name
 * @return {Object} Cost estimate with tokens and price
 */
function estimateBulkCost(rowCount, model) {
  // Use centralized pricing from PricingConfig.gs
  // This function is a wrapper for backward compatibility
  return estimateJobCost(model, rowCount, {
    avgInputTokens: 150,
    avgOutputTokens: 100
  });
}

// ============================================
// SMART AUTO-DETECTION
// ============================================

/**
 * Get info about the currently selected range in the sheet.
 * Used for auto-filling the bulk job form.
 * 
 * @return {Object} Selection info with range, rowCount, suggestedOutput
 */
function getSelectedRange() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var selection = sheet.getActiveRange();
    
    if (!selection) {
      return {
        hasSelection: false,
        message: 'No cells selected'
      };
    }
    
    var values = selection.getValues();
    var nonEmptyCount = 0;
    
    // Count non-empty cells
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (values[i][j] !== '' && values[i][j] !== null) {
          nonEmptyCount++;
        }
      }
    }
    
    // Get suggested output column (first empty column after selection)
    var suggestedOutput = getNextEmptyColumn_(sheet, selection);
    
    return {
      hasSelection: true,
      range: selection.getA1Notation(),
      column: columnToLetter_(selection.getColumn()),
      startRow: selection.getRow(),
      endRow: selection.getLastRow(),
      rowCount: nonEmptyCount,
      totalCells: selection.getNumRows() * selection.getNumColumns(),
      suggestedOutput: suggestedOutput,
      hasData: nonEmptyCount > 0
    };
    
  } catch (error) {
    return {
      hasSelection: false,
      message: 'Error reading selection: ' + error.message
    };
  }
}

/**
 * Find the next empty column after the given selection
 * @private
 */
function getNextEmptyColumn_(sheet, selection) {
  var startCol = selection.getColumn() + selection.getNumColumns();
  var startRow = selection.getRow();
  var endRow = selection.getLastRow();
  var maxCols = sheet.getMaxColumns();
  
  for (var col = startCol; col <= Math.min(startCol + 10, maxCols); col++) {
    var range = sheet.getRange(startRow, col, endRow - startRow + 1, 1);
    var values = range.getValues();
    var isEmpty = values.every(function(row) {
      return row[0] === '' || row[0] === null;
    });
    
    if (isEmpty) {
      return columnToLetter_(col);
    }
  }
  
  // Default to next column if no empty found within 10 columns
  return columnToLetter_(startCol);
}

/**
 * Convert column number to letter (1 = A, 27 = AA)
 * @private
 */
function columnToLetter_(column) {
  var letter = '';
  while (column > 0) {
    var temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = Math.floor((column - temp - 1) / 26);
  }
  return letter;
}

// ============================================
// PREVIEW MODE
// ============================================

/**
 * Preview bulk job with first N rows before full execution.
 * Allows users to validate their prompt works correctly.
 * 
 * @param {string} inputRange - Range in A1 notation
 * @param {string} prompt - Prompt template with {{input}} placeholder
 * @param {string} model - Provider (CHATGPT, CLAUDE, GROQ, GEMINI)
 * @param {string} [specificModel] - Optional specific model ID
 * @param {number} [previewCount] - Number of rows to preview (default 3)
 * @return {Object} Preview results with success/error for each row
 */
function previewBulkJob(inputRange, prompt, model, specificModel, previewCount) {
  previewCount = previewCount || 3;
  
  // Validate inputs
  if (!prompt || typeof prompt !== 'string' || prompt.trim() === '') {
    throw new Error('Prompt must be a non-empty string');
  }
  
  var inputs = getRangeValues(inputRange);
  
  if (inputs.length === 0) {
    throw new Error('No data found in selected range');
  }
  
  // Take only first N inputs for preview
  var previewInputs = inputs.slice(0, previewCount);
  var results = [];
  
  for (var i = 0; i < previewInputs.length; i++) {
    var input = previewInputs[i];
    var expandedPrompt = prompt.replace(/\{\{input\}\}/g, input);
    
    try {
      // Call the AI function directly
      var output = AIQuery(model, expandedPrompt, null, specificModel);
      
      results.push({
        index: i,
        input: truncateText_(input, 50),
        output: truncateText_(output, 200),
        fullOutput: output,
        success: true
      });
      
    } catch (error) {
      results.push({
        index: i,
        input: truncateText_(input, 50),
        error: error.message,
        success: false
      });
    }
  }
  
  var successCount = results.filter(function(r) { return r.success; }).length;
  
  return {
    previews: results,
    previewCount: previewInputs.length,
    successCount: successCount,
    errorCount: previewInputs.length - successCount,
    totalRows: inputs.length,
    estimatedCost: estimateBulkCost(inputs.length, model),
    allSuccessful: successCount === previewInputs.length
  };
}

/**
 * Truncate text to max length with ellipsis
 * @private
 */
function truncateText_(text, maxLength) {
  if (!text) return '';
  text = String(text);
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength - 3) + '...';
}

// ============================================
// INCREMENTAL WRITES
// ============================================

// Track last written index per job (session-scoped)
var _lastWrittenIndex = {};

/**
 * Write job results incrementally as they arrive.
 * Only writes NEW results since last call.
 * 
 * @param {string} jobId - The job ID
 * @param {Array<Object>} results - Array of {index, output, error} objects
 * @param {string} outputColumn - Column letter (e.g., "B")
 * @param {number} [startRow] - Starting row number (default 2)
 * @param {boolean} [highlight] - Whether to highlight new cells (default true)
 * @return {Object} Write stats with count and lastIndex
 */
function writeIncrementalResults(jobId, results, outputColumn, startRow, highlight) {
  startRow = startRow || 2;
  highlight = highlight !== false; // Default to true
  
  if (!results || results.length === 0) {
    return { written: 0, lastIndex: _lastWrittenIndex[jobId] || -1 };
  }
  
  var lastIndex = _lastWrittenIndex[jobId] || -1;
  
  // Filter to only NEW results
  var newResults = results.filter(function(r) {
    return r.index > lastIndex;
  });
  
  if (newResults.length === 0) {
    return { written: 0, lastIndex: lastIndex };
  }
  
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    
    // Sort by index to write in order
    newResults.sort(function(a, b) { return a.index - b.index; });
    
    // Batch write for performance (up to 50 at a time)
    var batchSize = 50;
    for (var batch = 0; batch < newResults.length; batch += batchSize) {
      var batchResults = newResults.slice(batch, batch + batchSize);
      
      batchResults.forEach(function(result) {
        var row = startRow + result.index;
        var value = result.output || result.error || '';
        var cell = sheet.getRange(outputColumn + row);
        
        cell.setValue(value);
        
        // Highlight based on success/error
        if (highlight) {
          if (result.error) {
            cell.setBackground('#fee2e2'); // Light red for errors
          } else {
            cell.setBackground('#dcfce7'); // Light green for success
          }
        }
      });
    }
    
    // Update tracker
    var maxIndex = Math.max.apply(null, newResults.map(function(r) { return r.index; }));
    _lastWrittenIndex[jobId] = maxIndex;
    
    if (Config.isDebug()) {
      Logger.log('Incremental write: ' + newResults.length + ' results to column ' + outputColumn);
    }
    
    return { 
      written: newResults.length,
      lastIndex: maxIndex,
      totalWritten: maxIndex + 1
    };
    
  } catch (error) {
    throw new Error('Failed to write incremental results: ' + error.message);
  }
}

/**
 * Clear highlighting from results column
 * 
 * @param {string} outputColumn - Column letter
 * @param {number} startRow - Starting row
 * @param {number} rowCount - Number of rows
 */
function clearResultHighlights(outputColumn, startRow, rowCount) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getRange(outputColumn + startRow + ':' + outputColumn + (startRow + rowCount - 1));
    range.setBackground(null);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Reset incremental write tracker for a job
 * Call this when starting a new job or retrying
 * 
 * @param {string} jobId - The job ID
 */
function resetIncrementalTracker(jobId) {
  delete _lastWrittenIndex[jobId];
}

// ============================================
// ERROR HANDLING & RETRY
// ============================================

/**
 * Retry only the failed rows from a completed job
 * 
 * @param {string} jobId - The job ID
 * @return {Object} Result with retriedCount and newJobId
 */
function retryFailedRows(jobId) {
  var userId = getUserId();
  
  try {
    var result = ApiClient.put('JOBS', {
      id: jobId,
      userId: userId,
      action: 'retry_failed'
    });
    
    // Reset tracker since we're retrying
    resetIncrementalTracker(jobId);
    
    if (Config.isDebug()) {
      Logger.log('Retrying failed rows for job: ' + jobId);
    }
    
    return {
      success: true,
      retriedCount: result.retriedCount || 0,
      message: 'Retrying ' + (result.retriedCount || 0) + ' failed rows'
    };
    
  } catch (error) {
    throw new Error('Failed to retry: ' + error.message);
  }
}

/**
 * Get detailed error information for a job
 * 
 * @param {string} jobId - The job ID
 * @return {Array<Object>} Array of error details
 */
function getJobErrors(jobId) {
  var userId = getUserId();
  
  try {
    var result = ApiClient.get('JOBS', { 
      id: jobId, 
      userId: userId,
      includeErrors: true 
    });
    
    return (result.errors || []).map(function(err) {
      return {
        rowIndex: err.row_index,
        input: err.input,
        error: err.error_message,
        retryable: err.retryable !== false,
        retryCount: err.retry_count || 0
      };
    });
    
  } catch (error) {
    throw new Error('Failed to get job errors: ' + error.message);
  }
}

// ============================================
// JOB PERSISTENCE
// ============================================

/**
 * Check if user has any active (in-progress) jobs.
 * Called when sidebar opens to resume watching.
 * 
 * @return {Object} Active job info or null
 */
function checkActiveJobs() {
  var userId = getUserId();
  
  try {
    var result = ApiClient.get('JOBS', { 
      userId: userId, 
      status: 'processing',
      limit: 1 
    });
    
    var jobs = result.jobs || result;
    
    if (jobs && jobs.length > 0) {
      var activeJob = jobs[0];
      return {
        hasActiveJob: true,
        job: {
          id: activeJob.id,
          status: activeJob.status,
          progress: activeJob.progress || 0,
          processedRows: activeJob.processedRows || 0,
          totalRows: activeJob.totalRows || 0,
          createdAt: activeJob.createdAt,
          config: activeJob.config
        }
      };
    }
    
    return { hasActiveJob: false };
    
  } catch (error) {
    // Fail silently - don't block sidebar from opening
    if (Config.isDebug()) {
      Logger.log('Error checking active jobs: ' + error.message);
    }
    return { hasActiveJob: false };
  }
}

// ============================================
// DATA VALIDATION DETECTION
// ============================================

/**
 * Get data validation rules for a specific column
 * Used to constrain AI outputs to valid values
 * 
 * @param {string} columnLetter - Column letter (e.g., "D")
 * @param {number} startRow - Starting row to check
 * @param {number} rowCount - Number of rows to check
 * @return {Object|null} Validation info { type, values, formula } or null if none
 */
function getColumnValidation(columnLetter, startRow, rowCount) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var validation = null;
    var foundAt = null;
    
    // Check multiple cells in the column (header row, first data row, second data row)
    // Validation might be applied to different ranges
    var rowsToCheck = [startRow, startRow - 1, startRow + 1, startRow + 2];
    
    for (var i = 0; i < rowsToCheck.length && !validation; i++) {
      var row = rowsToCheck[i];
      if (row < 1) continue;
      
      var cellRef = columnLetter + row;
      var cell = sheet.getRange(cellRef);
      validation = cell.getDataValidation();
      
      if (validation) {
        foundAt = cellRef;
        Logger.log('Validation found at ' + cellRef + ' (checked rows: ' + rowsToCheck.slice(0, i+1).join(', ') + ')');
      }
    }
    
    if (!validation) {
      Logger.log('No validation found for column ' + columnLetter + ' (checked rows: ' + rowsToCheck.filter(r => r >= 1).join(', ') + ')');
      return null;
    }
    
    Logger.log('Using validation from ' + foundAt);
    
    var criteriaType = validation.getCriteriaType();
    var criteriaValues = validation.getCriteriaValues();
    
    Logger.log('Validation type: ' + criteriaType);
    Logger.log('Criteria values: ' + JSON.stringify(criteriaValues));
    
    // Handle different validation types
    if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      // Dropdown list with specific values
      var values = criteriaValues[0] || [];
      Logger.log('Dropdown values found: [' + values.join(', ') + ']');
      return {
        type: 'list',
        values: values,
        helpOnInvalid: validation.getHelpText() || null
      };
    } else if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      // Values from a range
      var range = criteriaValues[0];
      if (range) {
        var rangeValues = range.getValues().flat().filter(function(v) {
          return v !== '' && v !== null && v !== undefined;
        });
        return {
          type: 'list',
          values: rangeValues,
          helpOnInvalid: validation.getHelpText() || null
        };
      }
    } else if (criteriaType === SpreadsheetApp.DataValidationCriteria.NUMBER_BETWEEN ||
               criteriaType === SpreadsheetApp.DataValidationCriteria.NUMBER_GREATER_THAN ||
               criteriaType === SpreadsheetApp.DataValidationCriteria.NUMBER_LESS_THAN) {
      return {
        type: 'number',
        min: criteriaValues[0],
        max: criteriaValues[1],
        helpOnInvalid: validation.getHelpText() || null
      };
    } else if (criteriaType === SpreadsheetApp.DataValidationCriteria.DATE_IS_VALID_DATE ||
               criteriaType === SpreadsheetApp.DataValidationCriteria.DATE_AFTER ||
               criteriaType === SpreadsheetApp.DataValidationCriteria.DATE_BEFORE) {
      return {
        type: 'date',
        helpOnInvalid: validation.getHelpText() || null
      };
    }
    
    // Other validation types - return generic info
    return {
      type: 'other',
      criteriaType: criteriaType.toString(),
      helpOnInvalid: validation.getHelpText() || null
    };
    
  } catch (e) {
    Logger.log('Error getting column validation: ' + e.message);
    return null;
  }
}

/**
 * Get data validation for multiple columns in a selection
 * 
 * @param {Array<string>} columns - Array of column letters
 * @param {number} startRow - Starting row
 * @param {number} rowCount - Number of rows
 * @return {Object} Map of column letter to validation info
 */
function getColumnsValidation(columns, startRow, rowCount) {
  var validationMap = {};
  
  columns.forEach(function(col) {
    var validation = getColumnValidation(col, startRow, rowCount);
    if (validation) {
      validationMap[col] = validation;
    }
  });
  
  return validationMap;
}

// ============================================
// HEADER READING
// ============================================

/**
 * Get headers from the first row of the active sheet
 * Used for smart prompt suggestions
 * @return {Array<string>} Array of header values
 */
function getSheetHeaders() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastColumn = sheet.getLastColumn();
  
  if (lastColumn < 1) {
    return [];
  }
  
  var headerRange = sheet.getRange(1, 1, 1, lastColumn);
  var headerValues = headerRange.getValues()[0];
  
  // Filter out empty headers and convert to strings
  var headers = headerValues
    .filter(function(h) { return h !== null && h !== undefined && String(h).trim() !== ''; })
    .map(function(h) { return String(h).trim(); });
  
  return headers;
}

/**
 * Infer column purpose from header name
 * Used for smart prompt generation
 * @param {string} header - Column header text
 * @return {Object} Inferred column info { type, suggestedAction }
 */
function inferColumnPurpose(header) {
  var h = header.toLowerCase();
  
  // Common patterns
  if (h.includes('email')) return { type: 'email', suggestedAction: 'extract' };
  if (h.includes('name') || h.includes('company')) return { type: 'name', suggestedAction: 'clean' };
  if (h.includes('description') || h.includes('text') || h.includes('content')) return { type: 'text', suggestedAction: 'summarize' };
  if (h.includes('category') || h.includes('type') || h.includes('status')) return { type: 'category', suggestedAction: 'classify' };
  if (h.includes('url') || h.includes('link') || h.includes('website')) return { type: 'url', suggestedAction: 'extract' };
  if (h.includes('phone') || h.includes('mobile') || h.includes('tel')) return { type: 'phone', suggestedAction: 'extract' };
  if (h.includes('address') || h.includes('location')) return { type: 'address', suggestedAction: 'clean' };
  if (h.includes('date') || h.includes('time')) return { type: 'date', suggestedAction: 'format' };
  if (h.includes('price') || h.includes('amount') || h.includes('cost')) return { type: 'number', suggestedAction: 'format' };
  if (h.includes('summary') || h.includes('tldr')) return { type: 'summary', suggestedAction: 'summarize' };
  if (h.includes('translation') || h.includes('translate')) return { type: 'translation', suggestedAction: 'translate' };
  
  return { type: 'unknown', suggestedAction: 'custom' };
}
