/**
 * @file SheetActions_Data.gs
 * @version 1.0.0
 * @created 2026-01-30
 * @author AISheeter Team
 * 
 * ============================================
 * SHEET ACTIONS - Data Writing
 * ============================================
 * 
 * Writes data to sheets and creates tables including:
 * - Write parsed table data (markdown, CSV, etc.)
 * - Create native Google Sheets Tables
 * - Apply table formatting (fallback)
 */

var SheetActions_Data = (function() {
  
  // ============================================
  // WRITE DATA
  // ============================================
  
  /**
   * Write parsed table data to the sheet
   * Used when user pastes table data (markdown, CSV, etc.) in their command
   * Also supports copying existing data to a new sheet
   * @param {Object} step - { config: { data: [[]], startCell: "A1" or range: "A1:D10" } }
   *                        OR { config: { dataRange: "A1:G51", newSheet: "SheetName" } }
   * @return {Object} Result with write details
   */
  function writeData(step) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var config = step.config || {};
    
    // Get data array - should be 2D array [[row1], [row2], ...]
    var data = config.data || step.data;
    
    // SPECIAL CASE: Copy existing data to a new sheet
    // If no data array but we have dataRange and newSheet, copy from existing range
    if ((!data || !Array.isArray(data) || data.length === 0) && config.dataRange && config.newSheet) {
      Logger.log('[SheetActions_Data] Copy to new sheet mode - dataRange: ' + config.dataRange + ', newSheet: ' + config.newSheet);
      return _copyDataToNewSheet(sheet, config);
    }
    
    if (!data || !Array.isArray(data) || data.length === 0) {
      throw new Error('No data provided to write. Expected 2D array.');
    }
    
    Logger.log('[SheetActions_Data] Received ' + data.length + ' rows of data');
    
    // Determine target location
    var startCell = config.startCell || config.range || step.startCell || 'A1';
    
    // If range is provided as A1:D10 format, extract just the start cell
    if (startCell.indexOf(':') !== -1) {
      startCell = startCell.split(':')[0];
    }
    
    // Parse start cell to get row and column
    var targetRange = sheet.getRange(startCell);
    var startRow = targetRange.getRow();
    var startCol = targetRange.getColumn();
    
    // Normalize data
    var normalizedData = _normalizeData(data);
    
    var numRows = normalizedData.length;
    var numCols = normalizedData[0] ? normalizedData[0].length : 0;
    
    if (numRows === 0 || numCols === 0) {
      throw new Error('No valid data to write after parsing.');
    }
    
    Logger.log('[SheetActions_Data] Writing ' + numRows + ' rows x ' + numCols + ' columns to ' + startCell);
    
    // Write data to sheet
    var writeRange = sheet.getRange(startRow, startCol, numRows, numCols);
    
    // Clear any existing formatting first to prevent auto-date conversion
    writeRange.clearFormat();
    
    // Write the data
    writeRange.setValues(normalizedData);
    
    // Apply appropriate number format to numeric columns
    _applyNumericFormats(sheet, normalizedData, startRow, startCol, numRows, numCols);
    
    // Get the written range in A1 notation
    var writtenRange = writeRange.getA1Notation();
    
    Logger.log('[SheetActions_Data] Wrote data to ' + writtenRange);
    
    return {
      range: writtenRange,
      rowsWritten: numRows,
      columnsWritten: numCols,
      startCell: startCell
    };
  }
  
  /**
   * Copy data from existing range to a new sheet
   * Handles the case where AI wants to copy existing data to a new sheet
   * @param {Sheet} sourceSheet - The source sheet to copy from
   * @param {Object} config - { dataRange: "A1:G51", newSheet: "SheetName", includeHeaders: true }
   * @return {Object} Result with copy details
   */
  function _copyDataToNewSheet(sourceSheet, config) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var dataRange = config.dataRange || 'A1';
    var newSheetName = config.newSheet;
    
    // Get source data
    var sourceRange = sourceSheet.getRange(dataRange);
    var sourceData = sourceRange.getValues();
    
    if (!sourceData || sourceData.length === 0) {
      throw new Error('No data found in range ' + dataRange);
    }
    
    Logger.log('[SheetActions_Data] Copying ' + sourceData.length + ' rows to new sheet: ' + newSheetName);
    
    // Check if sheet already exists
    var targetSheet = spreadsheet.getSheetByName(newSheetName);
    if (targetSheet) {
      // Sheet exists - append data or clear and write
      Logger.log('[SheetActions_Data] Sheet "' + newSheetName + '" already exists, clearing and writing');
      targetSheet.clear();
    } else {
      // Create new sheet
      targetSheet = spreadsheet.insertSheet(newSheetName);
      Logger.log('[SheetActions_Data] Created new sheet: ' + newSheetName);
    }
    
    // Write data to new sheet
    var numRows = sourceData.length;
    var numCols = sourceData[0].length;
    var targetRange = targetSheet.getRange(1, 1, numRows, numCols);
    targetRange.setValues(sourceData);
    
    // Copy formatting from source
    sourceRange.copyFormatToRange(targetSheet, 1, numCols, 1, numRows);
    
    // Auto-resize columns
    for (var col = 1; col <= numCols; col++) {
      targetSheet.autoResizeColumn(col);
    }
    
    // Optionally freeze header row
    if (config.includeHeaders !== false && numRows > 1) {
      targetSheet.setFrozenRows(1);
    }
    
    // Switch to the new sheet
    spreadsheet.setActiveSheet(targetSheet);
    
    Logger.log('[SheetActions_Data] Successfully copied data to ' + newSheetName);
    
    return {
      range: 'A1:' + _columnToLetter(numCols) + numRows,
      rowsWritten: numRows,
      columnsWritten: numCols,
      startCell: 'A1',
      newSheet: newSheetName,
      copiedFrom: dataRange
    };
  }
  
  /**
   * Convert column number to letter (1 -> A, 2 -> B, 27 -> AA, etc.)
   */
  function _columnToLetter(column) {
    var letter = '';
    while (column > 0) {
      var temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = Math.floor((column - temp - 1) / 26);
    }
    return letter;
  }
  
  /**
   * Normalize data - ensure all rows are arrays and pad to max width
   */
  function _normalizeData(data) {
    var normalizedData = [];
    var maxCols = 0;
    
    // First pass: convert to arrays and find max column count
    data.forEach(function(row) {
      var rowArray;
      if (Array.isArray(row)) {
        rowArray = row;
      } else if (typeof row === 'string') {
        // Handle string rows (might be comma or pipe separated)
        if (row.indexOf('|') !== -1) {
          rowArray = row.split('|').map(function(cell) { return cell.trim(); });
        } else if (row.indexOf('\t') !== -1) {
          rowArray = row.split('\t').map(function(cell) { return cell.trim(); });
        } else if (row.indexOf(',') !== -1) {
          rowArray = row.split(',').map(function(cell) { return cell.trim(); });
        } else {
          rowArray = [row];
        }
      } else {
        rowArray = [String(row)];
      }
      
      // Clean up cells (remove leading/trailing pipes from markdown)
      rowArray = rowArray.filter(function(cell, cellIdx) {
        if (cellIdx === 0 && cell === '') return false;
        if (cellIdx === rowArray.length - 1 && cell === '') return false;
        return true;
      });
      
      // CRITICAL FIX: Convert numeric strings to actual numbers
      // This ensures formulas work correctly on the data
      rowArray = rowArray.map(function(cell) {
        if (typeof cell === 'string') {
          var trimmed = cell.trim();
          
          // Check for percentage (e.g., "13.6%", "-11.2%")
          if (trimmed.endsWith('%')) {
            var pctStr = trimmed.slice(0, -1).trim(); // Remove %
            var pctValue = parseFloat(pctStr);
            if (!isNaN(pctValue)) {
              return pctValue / 100; // Convert 13.6% to 0.136, -11.2% to -0.112
            }
          }
          
          // Check for number (possibly with commas or currency symbols)
          var cleaned = trimmed.replace(/[$,€£¥₹]/g, '').trim();
          
          // Check for date patterns (MM/DD/YYYY, YYYY-MM-DD, etc.)
          var datePattern = /^(\d{1,4}[-\/]\d{1,2}[-\/]\d{1,4})$/;
          if (datePattern.test(cleaned)) {
            return cell; // Keep dates as strings
          }
          
          // Check if it's a valid number (including negative)
          // Allow: 123, -123, 123.45, -123.45, .45, -.45
          if (cleaned !== '' && /^-?\d*\.?\d+$/.test(cleaned)) {
            return Number(cleaned);
          }
        }
        return cell;
      });
      
      if (rowArray.length > maxCols) {
        maxCols = rowArray.length;
      }
      
      normalizedData.push(rowArray);
    });
    
    // Second pass: pad all rows to same length
    normalizedData = normalizedData.map(function(row) {
      while (row.length < maxCols) {
        row.push('');
      }
      return row;
    });
    
    // Skip markdown separator rows (e.g., |---|---|---|)
    normalizedData = normalizedData.filter(function(row) {
      return !row.every(function(cell) {
        return /^[\-\:\s]*$/.test(cell);
      });
    });
    
    return normalizedData;
  }
  
  /**
   * Apply number format to columns that contain numbers
   */
  function _applyNumericFormats(sheet, normalizedData, startRow, startCol, numRows, numCols) {
    if (numRows <= 1) return;
    
    var dataRow = normalizedData.length > 1 ? normalizedData[1] : null;
    var headerRow = normalizedData[0] || [];
    
    if (dataRow) {
      for (var colIdx = 0; colIdx < numCols; colIdx++) {
        var sampleValue = dataRow[colIdx];
        var header = (headerRow[colIdx] || '').toString().toLowerCase();
        var formatType = null;
        
        // Detect column type based on value and header
        if (typeof sampleValue === 'number') {
          // Small decimals (0-1 range) with "%" headers are percentages
          if (sampleValue > -1 && sampleValue < 1 && 
              (header.includes('growth') || header.includes('rate') || header.includes('percent') || header.includes('%'))) {
            formatType = 'percent';
          } else if (sampleValue > 1000 || header.includes('sales') || header.includes('target') || header.includes('amount') || header.includes('price') || header.includes('cost')) {
            // Large numbers or sales/target columns get thousands separator
            formatType = 'number';
          } else {
            formatType = 'auto';  // Let Sheets decide
          }
        }
        
        if (formatType && numRows > 1) {
          var colRange = sheet.getRange(startRow + 1, startCol + colIdx, numRows - 1, 1);
          if (formatType === 'percent') {
            colRange.setNumberFormat('0.0%');
            Logger.log('[SheetActions_Data] Applied percent format to column ' + (colIdx + 1) + ' (' + header + ')');
          } else if (formatType === 'number') {
            colRange.setNumberFormat('#,##0');
            Logger.log('[SheetActions_Data] Applied number format to column ' + (colIdx + 1) + ' (' + header + ')');
          }
        }
      }
    }
  }
  
  // ============================================
  // CREATE TABLE
  // ============================================
  
  /**
   * Convert a range to a Google Sheets Table (native feature)
   * Uses the Sheets API batchUpdate with addTable request
   * 
   * Supports the full Sheets API Table spec:
   * - name: Table name
   * - tableId: Optional table ID
   * - range: Sheet location
   * - columnProperties: Column types (PERCENT, DROPDOWN, numeric, date, checkbox, rating, smart_chip)
   * - freezeHeader: Freeze header row
   * 
   * @param {Object} step - { config: { range, tableName, columnProperties, freezeHeader } }
   * @return {Object} Result with table details
   * 
   * @see https://developers.google.com/workspace/sheets/api/guides/tables
   */
  function createTable(step) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SpreadsheetApp.getActiveSheet();
    var config = step.config || {};
    
    // Get range - supports multiple input formats
    var rangeStr = config.range || step.range || step.inputRange;
    if (!rangeStr) {
      rangeStr = SheetActions_Utils.detectDataRange(sheet);
    }
    
    var range = sheet.getRange(rangeStr);
    var tableName = config.tableName || config.name || 'Table_' + new Date().getTime();
    
    // Get sheet ID and range details
    var sheetId = sheet.getSheetId();
    var startRow = range.getRow() - 1;  // 0-indexed
    var startCol = range.getColumn() - 1;  // 0-indexed
    var endRow = startRow + range.getNumRows();
    var endCol = startCol + range.getNumColumns();
    
    Logger.log('[SheetActions_Data] Creating table "' + tableName + '" at ' + rangeStr);
    
    // Build the table object for Sheets API
    var tableObj = {
      name: tableName,
      range: {
        sheetId: sheetId,
        startRowIndex: startRow,
        endRowIndex: endRow,
        startColumnIndex: startCol,
        endColumnIndex: endCol
      }
    };
    
    // Optional: Table ID
    if (config.tableId) {
      tableObj.tableId = config.tableId;
    }
    
    // Optional: Column properties (column types, names, validation)
    // Supported column types: PERCENT, DROPDOWN, numeric, date, checkbox, rating, smart_chip
    if (config.columnProperties && Array.isArray(config.columnProperties)) {
      tableObj.columnProperties = config.columnProperties.map(function(col) {
        var colProp = {
          columnIndex: col.columnIndex
        };
        
        if (col.columnName) colProp.columnName = col.columnName;
        if (col.columnType) colProp.columnType = col.columnType.toUpperCase();
        
        // Dropdown columns require dataValidationRule with ONE_OF_LIST
        if (col.columnType && col.columnType.toUpperCase() === 'DROPDOWN' && col.values) {
          colProp.dataValidationRule = {
            condition: {
              type: 'ONE_OF_LIST',
              values: col.values.map(function(v) { return { userEnteredValue: String(v) }; })
            }
          };
        }
        
        return colProp;
      });
      
      Logger.log('[SheetActions_Data] Column properties: ' + JSON.stringify(tableObj.columnProperties).substring(0, 300));
    }
    
    // Build the addTable request
    var addTableRequest = {
      addTable: {
        table: tableObj
      }
    };
    
    // Execute via Sheets API
    try {
      var response = Sheets.Spreadsheets.batchUpdate(
        { requests: [addTableRequest] },
        spreadsheet.getId()
      );
      
      Logger.log('[SheetActions_Data] Table created successfully');
      
      // Optionally freeze the header row
      if (config.freezeHeader !== false) {
        sheet.setFrozenRows(range.getRow());
      }
      
      return {
        tableName: tableName,
        range: rangeStr,
        tableCreated: true,
        headerFrozen: config.freezeHeader !== false,
        columnProperties: config.columnProperties ? config.columnProperties.length : 0
      };
    } catch (e) {
      Logger.log('[SheetActions_Data] Error: ' + e.message);
      
      // Fallback: Apply table-like formatting
      Logger.log('[SheetActions_Data] Applying table-like formatting as fallback...');
      
      try {
        _applyTableFormatting(range);
        return {
          tableName: tableName,
          range: rangeStr,
          tableCreated: false,
          fallbackFormatApplied: true,
          note: 'Native table API unavailable. Applied table-style formatting instead.'
        };
      } catch (formatError) {
        throw new Error('Failed to create table: ' + e.message + '. Fallback also failed: ' + formatError.message);
      }
    }
  }
  
  /**
   * Update an existing Google Sheets Table
   * Uses the Sheets API batchUpdate with UpdateTableRequest
   * 
   * Supports:
   * - Modify table range (add/remove rows/columns)
   * - Toggle footer
   * 
   * @param {Object} step - { config: { tableId, range, showFooter } }
   * @return {Object} Result with update details
   * 
   * @see https://developers.google.com/workspace/sheets/api/guides/tables
   */
  function updateTable(step) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SpreadsheetApp.getActiveSheet();
    var config = step.config || {};
    
    var tableId = config.tableId;
    if (!tableId) {
      throw new Error('tableId is required for updateTable');
    }
    
    var requests = [];
    
    // Update range if provided
    if (config.range) {
      var range = sheet.getRange(config.range);
      var sheetId = sheet.getSheetId();
      
      requests.push({
        updateTable: {
          table: {
            tableId: tableId,
            range: {
              sheetId: sheetId,
              startRowIndex: range.getRow() - 1,
              endRowIndex: range.getRow() - 1 + range.getNumRows(),
              startColumnIndex: range.getColumn() - 1,
              endColumnIndex: range.getColumn() - 1 + range.getNumColumns()
            }
          },
          fields: 'range'
        }
      });
    }
    
    if (requests.length === 0) {
      throw new Error('No updates specified for updateTable');
    }
    
    var response = Sheets.Spreadsheets.batchUpdate(
      { requests: requests },
      spreadsheet.getId()
    );
    
    Logger.log('[SheetActions_Data] Table updated: ' + tableId);
    
    return { tableId: tableId, updated: true };
  }
  
  /**
   * Delete a Google Sheets Table
   * Uses the Sheets API batchUpdate with DeleteTableRequest
   * 
   * Options:
   * - deleteContents: true = delete table + data (DeleteTableRequest)
   * - deleteContents: false = remove table formatting only (DeleteBandingRequest style)
   * 
   * @param {Object} step - { config: { tableId, deleteContents } }
   * @return {Object} Result with deletion details
   * 
   * @see https://developers.google.com/workspace/sheets/api/guides/tables
   */
  function deleteTable(step) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var config = step.config || {};
    
    var tableId = config.tableId;
    if (!tableId) {
      throw new Error('tableId is required for deleteTable');
    }
    
    var request = {
      deleteTable: {
        tableId: tableId
      }
    };
    
    var response = Sheets.Spreadsheets.batchUpdate(
      { requests: [request] },
      spreadsheet.getId()
    );
    
    Logger.log('[SheetActions_Data] Table deleted: ' + tableId);
    
    return { tableId: tableId, deleted: true };
  }
  
  /**
   * Append rows to an existing Google Sheets Table
   * Uses the Sheets API batchUpdate with AppendCellsRequest (tableId)
   * 
   * Smart appending: adds to first free row, aware of footers.
   * If no empty rows, inserts before footer.
   * 
   * @param {Object} step - { config: { tableId, data: [[row1], [row2]] } }
   * @return {Object} Result with append details
   * 
   * @see https://developers.google.com/workspace/sheets/api/guides/tables
   */
  function appendToTable(step) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SpreadsheetApp.getActiveSheet();
    var config = step.config || {};
    
    var tableId = config.tableId;
    var data = config.data || step.data;
    
    if (!tableId) {
      throw new Error('tableId is required for appendToTable');
    }
    if (!data || !Array.isArray(data) || data.length === 0) {
      throw new Error('data array is required for appendToTable');
    }
    
    // Build rows for AppendCellsRequest
    var rows = data.map(function(row) {
      var cells = (Array.isArray(row) ? row : [row]).map(function(cellValue) {
        var cell = { userEnteredValue: {} };
        if (typeof cellValue === 'number') {
          cell.userEnteredValue.numberValue = cellValue;
        } else if (typeof cellValue === 'boolean') {
          cell.userEnteredValue.boolValue = cellValue;
        } else if (cellValue === null || cellValue === undefined || cellValue === '') {
          cell.userEnteredValue.stringValue = '';
        } else {
          cell.userEnteredValue.stringValue = String(cellValue);
        }
        return cell;
      });
      return { values: cells };
    });
    
    var request = {
      appendCells: {
        sheetId: sheet.getSheetId(),
        tableId: tableId,
        rows: rows,
        fields: 'userEnteredValue'
      }
    };
    
    var response = Sheets.Spreadsheets.batchUpdate(
      { requests: [request] },
      spreadsheet.getId()
    );
    
    Logger.log('[SheetActions_Data] Appended ' + data.length + ' rows to table: ' + tableId);
    
    return { tableId: tableId, rowsAppended: data.length };
  }
  
  /**
   * Apply table-like formatting (fallback when Sheets API is unavailable)
   * Creates a professional-looking table with:
   * - Bold header with color
   * - Filter dropdowns on header row
   * - Alternating row colors (banding)
   * - Borders
   * - Auto-resized columns
   * - Frozen header row
   */
  function _applyTableFormatting(range) {
    var sheet = range.getSheet();
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
    var startRow = range.getRow();
    var startCol = range.getColumn();
    
    // Header row styling
    var headerRange = sheet.getRange(startRow, startCol, 1, numCols);
    headerRange.setBackground('#4285F4')
               .setFontColor('#FFFFFF')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
    
    // Add borders
    range.setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
    
    // Use native banding instead of manual alternating colors (more efficient)
    try {
      // Remove any existing banding first
      var existingBandings = range.getBandings();
      for (var b = 0; b < existingBandings.length; b++) {
        existingBandings[b].remove();
      }
      
      // Apply banding to data rows (excluding header)
      if (numRows > 1) {
        var dataRange = sheet.getRange(startRow + 1, startCol, numRows - 1, numCols);
        dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
      }
      Logger.log('[SheetActions_Data] Applied native banding');
    } catch (bandingError) {
      // Fallback to manual alternating colors
      Logger.log('[SheetActions_Data] Banding failed, using manual alternating: ' + bandingError.message);
      if (numRows > 1) {
        for (var i = 1; i < numRows; i++) {
          var rowRange = sheet.getRange(startRow + i, startCol, 1, numCols);
          if (i % 2 === 0) {
            rowRange.setBackground('#F8F9FA');
          } else {
            rowRange.setBackground('#FFFFFF');
          }
        }
      }
    }
    
    // Add filter dropdowns (CRITICAL for table-like experience)
    try {
      var existingFilter = sheet.getFilter();
      if (existingFilter) {
        existingFilter.remove();
      }
      range.createFilter();
      Logger.log('[SheetActions_Data] Added filter to table');
    } catch (filterError) {
      Logger.log('[SheetActions_Data] Could not add filter: ' + filterError.message);
    }
    
    // Freeze header row
    if (startRow === 1) {
      sheet.setFrozenRows(1);
      Logger.log('[SheetActions_Data] Froze header row');
    }
    
    // Auto-resize columns
    for (var col = startCol; col < startCol + numCols; col++) {
      sheet.autoResizeColumn(col);
    }
    
    Logger.log('[SheetActions_Data] Applied table-style formatting with filter and banding');
  }
  
  // ============================================
  // PUBLIC API
  // ============================================
  
  return {
    writeData: writeData,
    createTable: createTable,
    updateTable: updateTable,
    deleteTable: deleteTable,
    appendToTable: appendToTable
  };
  
})();
