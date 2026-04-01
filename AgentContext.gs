/**
 * @file AgentContext.gs
 * @description Unified Context System for AI Sheet Agent
 * 
 * VISION: Every AI interaction should have FULL context.
 * This is our edge - we don't just send one column to AI,
 * we send the complete picture so AI can make smart decisions.
 * 
 * ARCHITECTURE:
 * 1. Build rich context from user's selection
 * 2. Pass context to FormulaFirst for evaluation
 * 3. If formula not possible, pass SAME context to AI
 * 4. AI always sees: all columns, headers, data types, samples
 */

// ============================================
// UNIFIED CONTEXT BUILDER
// ============================================

/**
 * Build comprehensive context from the current selection
 * This is the SINGLE SOURCE OF TRUTH for all context needs
 * 
 * @param {Object} options - Optional overrides
 * @param {string} options.command - User's command (for intent analysis)
 * @return {Object} Rich context object
 */
function buildUnifiedContext(options) {
  options = options || {};
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var selection = sheet.getActiveRange();
  
  if (!selection) {
    return { error: 'No selection', columns: [], rows: 0 };
  }
  
  var startRow = selection.getRow();
  var endRow = selection.getLastRow();
  var startCol = selection.getColumn();
  var endCol = selection.getLastColumn();
  var numRows = endRow - startRow + 1;
  var numCols = endCol - startCol + 1;
  
  // Detect if first row is headers
  var headerRow = startRow;
  var dataStartRow = startRow + 1;
  var values = selection.getValues();
  var firstRowValues = values[0] || [];
  
  // Heuristic: if first row has text and subsequent rows have mixed types, it's a header
  var hasHeaders = firstRowValues.every(function(v) {
    return typeof v === 'string' && v.trim().length > 0 && v.length < 50;
  });
  
  if (!hasHeaders && numRows > 1) {
    // No clear header row - use column letters
    headerRow = null;
    dataStartRow = startRow;
  }
  
  // Build column information
  var columns = [];
  var dataColumns = [];
  var emptyColumns = [];
  var columnMap = {};  // For quick lookup
  
  for (var c = 0; c < numCols; c++) {
    var colIndex = startCol + c;
    var colLetter = columnToLetter(colIndex);
    var header = hasHeaders ? String(firstRowValues[c] || '').trim() : 'Column ' + colLetter;
    
    // Analyze column data
    var colData = [];
    var numericCount = 0;
    var textCount = 0;
    var emptyCount = 0;
    
    for (var r = hasHeaders ? 1 : 0; r < values.length; r++) {
      var val = values[r][c];
      if (val === null || val === undefined || String(val).trim() === '') {
        emptyCount++;
      } else {
        colData.push(val);
        if (typeof val === 'number' || !isNaN(parseFloat(val))) {
          numericCount++;
        } else {
          textCount++;
        }
      }
    }
    
    var dataType = 'unknown';
    var totalNonEmpty = numericCount + textCount;
    if (totalNonEmpty > 0) {
      if (numericCount > totalNonEmpty * 0.7) {
        dataType = 'numeric';
      } else if (textCount > totalNonEmpty * 0.7) {
        dataType = 'text';
      } else {
        dataType = 'mixed';
      }
    }
    
    var isEmpty = emptyCount === (values.length - (hasHeaders ? 1 : 0));
    var samples = colData.slice(0, 5);  // First 5 non-empty values
    
    var colInfo = {
      column: colLetter,
      index: colIndex,
      header: header,
      dataType: dataType,
      isEmpty: isEmpty,
      sampleValues: samples,
      rowCount: colData.length,
      numericRatio: totalNonEmpty > 0 ? numericCount / totalNonEmpty : 0
    };
    
    columns.push(colInfo);
    columnMap[colLetter] = colInfo;
    
    if (isEmpty) {
      emptyColumns.push(colInfo);
    } else {
      dataColumns.push(colInfo);
    }
  }
  
  // Build the unified context object
  var context = {
    // Selection info
    sheetName: sheet.getName(),
    sheetId: sheet.getSheetId(),
    range: selection.getA1Notation(),
    startRow: startRow,
    endRow: endRow,
    startCol: startCol,
    endCol: endCol,
    
    // Row info
    hasHeaders: hasHeaders,
    headerRow: headerRow,
    dataStartRow: dataStartRow,
    dataRowCount: hasHeaders ? numRows - 1 : numRows,
    
    // Column info (the key differentiator!)
    columns: columns,
    columnMap: columnMap,
    dataColumns: dataColumns,
    emptyColumns: emptyColumns,
    
    // Quick access
    inputColumns: dataColumns.map(function(c) { return c.column; }),
    outputColumns: emptyColumns.map(function(c) { return c.column; }),
    
    // Headers map for easy access
    headers: {},
    
    // Data summary (for AI context)
    dataSummary: buildDataSummary(dataColumns),
    
    // Original command (if provided)
    command: options.command || null,
    
    // Timestamp
    capturedAt: new Date().toISOString()
  };
  
  // Build headers map
  columns.forEach(function(col) {
    context.headers[col.column] = col.header;
  });
  
  Logger.log('[AgentContext] Built unified context:');
  Logger.log('  Range: ' + context.range);
  Logger.log('  Data columns: ' + context.inputColumns.join(', '));
  Logger.log('  Empty columns: ' + context.outputColumns.join(', '));
  Logger.log('  Data rows: ' + context.dataRowCount);
  
  return context;
}

/**
 * Build a human-readable data summary for AI context
 */
function buildDataSummary(dataColumns) {
  if (!dataColumns || dataColumns.length === 0) {
    return 'No data columns detected.';
  }
  
  var lines = ['Available data columns:'];
  
  dataColumns.forEach(function(col) {
    var typeInfo = col.dataType;
    if (col.dataType === 'numeric' && col.sampleValues.length > 0) {
      var nums = col.sampleValues.filter(function(v) { return !isNaN(parseFloat(v)); });
      if (nums.length > 0) {
        var min = Math.min.apply(null, nums);
        var max = Math.max.apply(null, nums);
        typeInfo = 'numeric (range: ' + min + ' to ' + max + ')';
      }
    }
    
    var samples = col.sampleValues.slice(0, 3).map(function(v) {
      return '"' + String(v).substring(0, 25) + '"';
    }).join(', ');
    
    lines.push('- ' + col.column + ' ("' + col.header + '"): ' + typeInfo);
    if (samples) {
      lines.push('    Samples: ' + samples);
    }
  });
  
  return lines.join('\n');
}

/**
 * Convert column index to letter
 */
function columnToLetter(colIndex) {
  var letter = '';
  while (colIndex > 0) {
    var mod = (colIndex - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    colIndex = Math.floor((colIndex - 1) / 26);
  }
  return letter;
}

// ============================================
// CONTEXT-AWARE FORMULA-FIRST CHECK
// ============================================

/**
 * Check if task can be solved with formula, using FULL context
 * This is the main entry point for the formula-first flow
 * 
 * @param {Object} taskInfo - Task information
 * @param {string} taskInfo.command - User's command
 * @param {string} taskInfo.taskType - Detected task type
 * @param {string} taskInfo.categories - Target categories (for classification)
 * @param {Object} context - Unified context from buildUnifiedContext
 * @return {Object} Formula evaluation result
 */
function checkFormulaFirstWithContext(taskInfo, context) {
  Logger.log('[AgentContext] Formula-first check with full context');
  
  if (!context || !context.columns || context.columns.length === 0) {
    return {
      canUseFormula: false,
      explanation: 'No valid context available'
    };
  }
  
  // Build the rich context for FormulaFirst
  var formulaContext = {
    command: taskInfo.command,
    taskType: taskInfo.taskType,
    categories: taskInfo.categories,
    
    // Full column context
    allColumns: context.inputColumns,
    columnDetails: context.dataColumns,
    columnHeaders: context.headers,
    emptyColumns: context.outputColumns,
    
    // Suggested input/output based on context
    suggestedInputColumn: context.inputColumns[0],
    suggestedOutputColumn: context.outputColumns[0] || nextColumn(context.inputColumns[context.inputColumns.length - 1]),
    
    // Range info
    dataStartRow: context.dataStartRow,
    dataEndRow: context.endRow,
    inputRange: context.inputColumns[0] + context.dataStartRow + ':' + 
                context.inputColumns[context.inputColumns.length - 1] + context.endRow,
    
    // Data summary for AI prompt
    dataSummary: context.dataSummary,
    
    // Is this multi-column input?
    isMultiColumn: context.inputColumns.length > 1
  };
  
  // Call the formula evaluator with full context
  return evaluateFormulaFirst(formulaContext);
}

/**
 * Get next column letter
 */
function nextColumn(col) {
  if (!col) return 'B';
  return String.fromCharCode(col.charCodeAt(0) + 1);
}

// ============================================
// AI PROMPT BUILDER WITH FULL CONTEXT
// ============================================

/**
 * Build an AI prompt that includes FULL data context
 * This ensures AI always has complete information to make smart decisions
 * 
 * @param {Object} taskInfo - Task information
 * @param {Object} context - Unified context
 * @return {string} Enhanced prompt with full context
 */
function buildContextAwarePrompt(taskInfo, context) {
  var basePrompt = taskInfo.prompt || 'Process this data:';
  
  if (!context || !context.dataColumns || context.dataColumns.length === 0) {
    return basePrompt;
  }
  
  var contextSection = [
    '## DATA CONTEXT',
    context.dataSummary,
    '',
    '## YOUR TASK',
    taskInfo.command || basePrompt
  ].join('\n');
  
  // For classification tasks, be more specific
  if (taskInfo.taskType === 'classify' || taskInfo.taskType === 'categorize') {
    var categories = taskInfo.categories || 'appropriate category';
    return [
      contextSection,
      '',
      '## INSTRUCTIONS',
      'Based on ALL the data columns above, classify each row into: ' + categories,
      'Consider all available information (not just one column) to make your decision.',
      'Respond with ONLY the category name, nothing else.'
    ].join('\n');
  }
  
  return contextSection + '\n\n' + basePrompt;
}

// ============================================
// ROW DATA BUILDER FOR AI PROCESSING
// ============================================

/**
 * Build row data for AI that includes ALL column values
 * Each row sent to AI has the complete picture
 * 
 * @param {number} rowIndex - Row index in the data
 * @param {Object} context - Unified context
 * @return {Object} Row data with all column values
 */
function buildRowDataForAI(rowIndex, context) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var actualRow = context.dataStartRow + rowIndex;
  
  var rowData = {
    rowNumber: actualRow,
    values: {},
    formattedText: ''
  };
  
  var textParts = [];
  
  context.dataColumns.forEach(function(col) {
    var cell = sheet.getRange(actualRow, col.index);
    var value = cell.getValue();
    
    rowData.values[col.column] = value;
    textParts.push(col.header + ': ' + value);
  });
  
  rowData.formattedText = textParts.join('\n');
  
  return rowData;
}

// ============================================
// MULTI-SHEET SUPPORT
// ============================================

/**
 * Get list of all sheets in the spreadsheet
 * @return {Array<Object>} Array of sheet info objects
 */
function getAllSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var activeSheet = spreadsheet.getActiveSheet();
  
  return sheets.map(function(sheet) {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    return {
      name: sheet.getName(),
      sheetId: sheet.getSheetId(),
      isActive: sheet.getSheetId() === activeSheet.getSheetId(),
      rowCount: lastRow,
      colCount: lastCol,
      isEmpty: lastRow === 0 || lastCol === 0,
      // Get headers if sheet has data
      headers: lastRow > 0 && lastCol > 0 
        ? sheet.getRange(1, 1, 1, Math.min(lastCol, 10)).getValues()[0].filter(function(h) { return h; })
        : []
    };
  });
}

/**
 * Get data from a specific sheet
 * @param {string} sheetName - Name of the sheet
 * @param {string} range - Optional A1 notation range (e.g., "A1:C10")
 * @return {Object} Sheet data context
 */
function getSheetData(sheetName, range) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    return { error: 'Sheet "' + sheetName + '" not found' };
  }
  
  var dataRange;
  if (range) {
    try {
      dataRange = sheet.getRange(range);
    } catch (e) {
      return { error: 'Invalid range: ' + range };
    }
  } else {
    // Get entire data range
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) {
      return { error: 'Sheet "' + sheetName + '" is empty' };
    }
    dataRange = sheet.getRange(1, 1, lastRow, lastCol);
  }
  
  var values = dataRange.getValues();
  var headers = values[0] || [];
  
  return {
    sheetName: sheetName,
    sheetId: sheet.getSheetId(),
    range: dataRange.getA1Notation(),
    headers: headers,
    rowCount: values.length - 1, // Exclude header
    colCount: headers.length,
    sampleData: values.slice(1, 4) // First 3 data rows
  };
}

/**
 * Write data to a specific sheet
 * @param {string} sheetName - Name of the sheet
 * @param {Array} data - 2D array of data to write
 * @param {string} startCell - Starting cell (e.g., "A1")
 * @return {Object} Result with success status
 */
function writeToSheet(sheetName, data, startCell) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    // Create sheet if it doesn't exist
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  startCell = startCell || 'A1';
  
  try {
    var range = sheet.getRange(startCell);
    var startRow = range.getRow();
    var startCol = range.getColumn();
    
    if (Array.isArray(data) && data.length > 0) {
      // Ensure data is 2D array
      var dataToWrite = data.map(function(row) {
        return Array.isArray(row) ? row : [row];
      });
      
      var numRows = dataToWrite.length;
      var numCols = Math.max.apply(null, dataToWrite.map(function(r) { return r.length; }));
      
      // Pad rows to same length
      dataToWrite = dataToWrite.map(function(row) {
        while (row.length < numCols) row.push('');
        return row;
      });
      
      sheet.getRange(startRow, startCol, numRows, numCols).setValues(dataToWrite);
      
      return {
        success: true,
        sheetName: sheetName,
        range: sheet.getRange(startRow, startCol, numRows, numCols).getA1Notation(),
        rowsWritten: numRows
      };
    }
    
    return { success: false, error: 'No data to write' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Get context with multi-sheet awareness
 * Includes info about other sheets for AI to reference
 * 
 * @param {string} command - User's command (may reference other sheets)
 * @return {Object} Enhanced context with sheet info
 */
function getMultiSheetContext(command) {
  var baseContext = buildUnifiedContext({ command: command });
  
  // Add multi-sheet info
  baseContext.availableSheets = getAllSheets();
  baseContext.activeSheetName = SpreadsheetApp.getActiveSheet().getName();
  
  // Check if command references another sheet
  var sheetMatch = (command || '').match(/(?:from|in|on|to)\s+(?:sheet\s+)?["']?([^"'\s,]+)["']?/i);
  if (sheetMatch) {
    var referencedSheet = sheetMatch[1];
    var matchingSheet = baseContext.availableSheets.find(function(s) {
      return s.name.toLowerCase() === referencedSheet.toLowerCase();
    });
    if (matchingSheet) {
      baseContext.referencedSheet = matchingSheet;
      baseContext.crossSheetOperation = true;
    }
  }
  
  return baseContext;
}

// ============================================
// EXPORTS
// ============================================

/**
 * Get unified context - exposed for frontend
 */
function getUnifiedContext(command) {
  return buildUnifiedContext({ command: command });
}

/**
 * Get multi-sheet aware context - exposed for frontend
 */
function getContextWithSheets(command) {
  return getMultiSheetContext(command);
}
