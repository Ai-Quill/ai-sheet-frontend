/**
 * @file SheetActions_Format.gs
 * @version 2.0.0
 * @updated 2026-02-11
 * @author AISheeter Team
 * 
 * ============================================
 * SHEET ACTIONS - Cell Formatting
 * ============================================
 * 
 * Applies formatting to cells and ranges including:
 * - Number/date formats (currency, percent, date, etc.)
 * - Font styles (bold, italic, underline, strikethrough)
 * - Colors (background, text)
 * - Alignment (horizontal, vertical)
 * - Text wrapping and rotation
 * - Cell merging
 * - Borders (basic and advanced)
 * - Row banding (alternating colors)
 * - Notes/comments
 */

var SheetActions_Format = (function() {
  
  // ============================================
  // MAIN FORMAT FUNCTION
  // ============================================
  
  /**
   * Apply number/style formatting to a range (or multiple comma-separated ranges)
   * Supports:
   * - Single operation: { formatType, range, options }
   * - Multiple operations: { operations: [{ range, formatting }, ...] }
   * 
   * ROBUSTNESS: Checks multiple locations for parameters since the AI SDK agent
   * may place them at different nesting levels (config.*, step.*, etc.)
   * 
   * @param {Object} step - { formatType, range, config }
   * @return {Object} Result with formatting details
   */
  function applyFormat(step) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var config = step.config || {};
    var results = [];
    
    // DIAGNOSTIC: Log what we received to help debug parameter routing
    Logger.log('[SheetActions_Format] === RECEIVED STEP ===');
    Logger.log('[SheetActions_Format] step keys: ' + Object.keys(step).join(', '));
    Logger.log('[SheetActions_Format] config keys: ' + Object.keys(config).join(', '));
    Logger.log('[SheetActions_Format] step.operations: ' + (step.operations ? JSON.stringify(step.operations).substring(0, 300) : 'null'));
    Logger.log('[SheetActions_Format] config.operations: ' + (config.operations ? JSON.stringify(config.operations).substring(0, 300) : 'null'));
    Logger.log('[SheetActions_Format] config.options: ' + (config.options ? JSON.stringify(config.options).substring(0, 200) : 'null'));
    Logger.log('[SheetActions_Format] step.options: ' + (step.options ? JSON.stringify(step.options).substring(0, 200) : 'null'));
    Logger.log('[SheetActions_Format] config.range: ' + (config.range || 'null'));
    Logger.log('[SheetActions_Format] step.range: ' + (step.range || 'null'));
    
    // Check if we have multiple operations (AI's flexible format)
    // ROBUSTNESS: Check both config.operations AND step.operations (AI may place at either level)
    var operations = config.operations || step.operations || [];
    
    if (operations.length > 0) {
      // Handle multiple operations
      Logger.log('[SheetActions_Format] Processing ' + operations.length + ' operation(s)');
      var errorCount = 0;
      
      for (var idx = 0; idx < operations.length; idx++) {
        var op = operations[idx];
        var opRange = op.range;
        
        // ROBUSTNESS: Merge formatting from multiple locations
        // AI may put properties at op.formatting (correct) or at op level directly (wrong but common)
        var formatting = op.formatting || {};
        
        // CRITICAL FIX: Also check for format properties placed directly on the operation object
        // (AI sometimes puts textRotation, wrap, bold etc. at the operation level instead of inside formatting)
        var FORMAT_KEYS = ['bold', 'italic', 'underline', 'strikethrough', 'fontSize', 'fontFamily',
                           'textColor', 'fontColor', 'color', 'backgroundColor', 'background',
                           'horizontalAlignment', 'alignment', 'align', 'verticalAlignment', 'verticalAlign',
                           'wrap', 'wrapStrategy', 'textRotation', 'textDirection',
                           'merge', 'unmerge', 'borders', 'borderColor', 'borderStyle',
                           'borderTop', 'borderBottom', 'borderLeft', 'borderRight',
                           'borderVertical', 'borderHorizontal',
                           'banding', 'bandingTheme', 'showHeader', 'showFooter', 'removeBanding',
                           'note', 'clearNote', 'clearFormatting', 'clearFormat',
                           'autoFitColumns', 'autoResize', 'columnWidth', 'rowHeight',
                           'numberFormat', 'formatType', 'decimals', 'locale', 'pattern'];
        
        for (var ki = 0; ki < FORMAT_KEYS.length; ki++) {
          var key = FORMAT_KEYS[ki];
          if (op[key] !== undefined && formatting[key] === undefined) {
            formatting[key] = op[key];
            Logger.log('[SheetActions_Format] Hoisted op-level property to formatting: ' + key + '=' + op[key]);
          }
        }
        
        // Also merge from op.options if present (AI may nest formatting under options)
        if (op.options && typeof op.options === 'object') {
          for (var oki = 0; oki < FORMAT_KEYS.length; oki++) {
            var okey = FORMAT_KEYS[oki];
            if (op.options[okey] !== undefined && formatting[okey] === undefined) {
              formatting[okey] = op.options[okey];
              Logger.log('[SheetActions_Format] Hoisted op.options property to formatting: ' + okey + '=' + op.options[okey]);
            }
          }
        }
        
        if (!opRange) {
          Logger.log('[SheetActions_Format] Skipping operation ' + idx + ' - no range specified');
          continue;
        }
        
        Logger.log('[SheetActions_Format] Op ' + idx + ': range=' + opRange + ', formatting=' + JSON.stringify(formatting).substring(0, 200));
        
        try {
          var range = sheet.getRange(opRange);
          
          // Handle formatType at operation level (not just inside formatting)
          // AI might put formatType at op.formatType or op.formatting.formatType
          var opFormatType = op.formatType || formatting.formatType || formatting.numberFormat;
          var opNumberFormatApplied = false;
          if (opFormatType && opFormatType !== 'text') {
            var formatOptions = op.options || formatting;
            var opFormatStr = SheetActions_Utils.getNumberFormat(opFormatType, formatOptions);
            range.setNumberFormat(opFormatStr);
            opNumberFormatApplied = true;
            Logger.log('[SheetActions_Format] Applied number format "' + opFormatStr + '" (from: "' + opFormatType + '") to ' + opRange);
          }
          
          // Apply other formatting (font, colors, borders, etc.)
          // Pass flag to prevent double-applying number format
          _applyFormattingToRange(range, formatting, opNumberFormatApplied);
          Logger.log('[SheetActions_Format] ✅ Completed operation ' + idx + ' on range ' + opRange);
          results.push({ range: opRange, success: true });
        } catch (e) {
          Logger.log('[SheetActions_Format] ❌ Error in operation ' + idx + ' on range ' + opRange + ': ' + e.message);
          results.push({ range: opRange, success: false, error: e.message });
          errorCount++;
        }
      }
      
      // Flush changes to ensure they're committed
      SpreadsheetApp.flush();
      
      // Log summary
      Logger.log('[SheetActions_Format] Multi-op summary: ' + operations.length + ' operations, ' + errorCount + ' errors');
      
      return { operationCount: operations.length, results: results, errorCount: errorCount };
    }
    
    // Fallback to single operation format
    var rangeStr = config.range || step.range || step.inputRange;
    var formatType = config.formatType || step.formatType;
    // ROBUSTNESS: Check multiple locations for options
    var options = config.options || step.options || {};
    
    // CRITICAL FIX: If options is empty, check if format properties are at the config level directly
    // (AI sometimes puts textRotation, wrap etc. directly in config instead of config.options)
    if (Object.keys(options).length === 0) {
      var configAsOptions = {};
      var OPTION_KEYS = ['bold', 'italic', 'underline', 'strikethrough', 'fontSize', 'fontFamily',
                         'textColor', 'fontColor', 'color', 'backgroundColor', 'background',
                         'horizontalAlignment', 'alignment', 'align', 'verticalAlignment', 'verticalAlign',
                         'wrap', 'wrapStrategy', 'textRotation', 'textDirection',
                         'merge', 'unmerge', 'borders', 'borderColor', 'borderStyle',
                         'borderTop', 'borderBottom', 'borderLeft', 'borderRight',
                         'borderVertical', 'borderHorizontal',
                         'banding', 'bandingTheme', 'showHeader', 'showFooter', 'removeBanding',
                         'note', 'clearNote', 'clearFormatting', 'clearFormat',
                         'autoFitColumns', 'autoResize', 'columnWidth', 'rowHeight',
                         'numberFormat', 'decimals', 'locale', 'pattern'];
      var hoisted = false;
      for (var ci = 0; ci < OPTION_KEYS.length; ci++) {
        var ck = OPTION_KEYS[ci];
        if (config[ck] !== undefined) {
          configAsOptions[ck] = config[ck];
          hoisted = true;
        }
      }
      if (hoisted) {
        options = configAsOptions;
        Logger.log('[SheetActions_Format] Hoisted config-level properties to options: ' + JSON.stringify(options).substring(0, 200));
      }
    }
    
    Logger.log('[SheetActions_Format] Single format: range=' + rangeStr + ', formatType=' + formatType);
    Logger.log('[SheetActions_Format] Options: ' + JSON.stringify(options).substring(0, 200));
    
    // Handle multiple comma-separated ranges
    var rangeStrings = rangeStr.split(',').map(function(r) { return r.trim(); }).filter(Boolean);
    Logger.log('[SheetActions_Format] Processing ' + rangeStrings.length + ' range(s): ' + rangeStrings.join(', '));
    
    rangeStrings.forEach(function(singleRangeStr) {
      try {
        var range = sheet.getRange(singleRangeStr);
        
        // Number/date formatting (top-level formatType takes priority)
        var numberFormatApplied = false;
        if (formatType && formatType !== 'text') {
          var formatStr = SheetActions_Utils.getNumberFormat(formatType, options);
          range.setNumberFormat(formatStr);
          numberFormatApplied = true;
          Logger.log('[SheetActions_Format] Applied number format: ' + formatStr + ' (from formatType: ' + formatType + ')');
        }
        
        // Style formatting (pass flag to prevent double-applying number format)
        _applyFormattingToRange(range, options, numberFormatApplied);
        
        Logger.log('[SheetActions_Format] ✅ Successfully formatted range ' + singleRangeStr);
      } catch (e) {
        Logger.log('[SheetActions_Format] ❌ Error formatting range ' + singleRangeStr + ': ' + e.message);
        throw new Error('Failed to format range ' + singleRangeStr + ': ' + e.message);
      }
    });
    
    // Flush changes to ensure they're committed
    SpreadsheetApp.flush();
    
    return { formatType: formatType, range: rangeStr, rangeCount: rangeStrings.length };
  }
  
  // ============================================
  // FORMATTING HELPERS
  // ============================================
  
  /**
   * Apply formatting options to a range
   * @param {Range} range - Google Sheets Range object
   * @param {Object} formatting - Formatting options
   * @param {boolean} [numberFormatAlreadyApplied] - If true, skip number format to prevent double-apply
   */
  function _applyFormattingToRange(range, formatting, numberFormatAlreadyApplied) {
    var rangeA1 = range.getA1Notation();
    var appliedCount = 0;
    
    // ===== CLEAR FORMATTING (must be first — clears everything, then re-apply) =====
    if (formatting.clearFormatting || formatting.clearFormat) {
      range.clearFormat();
      Logger.log('[SheetActions_Format] Cleared all formatting on ' + rangeA1);
      appliedCount++;
    }
    
    // ===== FONT STYLES =====
    // Support both setting and unsetting (bold: true → bold, bold: false → normal)
    if (formatting.bold === true) { range.setFontWeight('bold'); appliedCount++; }
    else if (formatting.bold === false) { range.setFontWeight('normal'); appliedCount++; }
    
    if (formatting.italic === true) { range.setFontStyle('italic'); appliedCount++; }
    else if (formatting.italic === false) { range.setFontStyle('normal'); appliedCount++; }
    
    if (formatting.underline === true) { range.setFontLine('underline'); appliedCount++; }
    else if (formatting.underline === false) { range.setFontLine('none'); appliedCount++; }
    
    if (formatting.strikethrough === true) { range.setFontLine('line-through'); appliedCount++; }
    else if (formatting.strikethrough === false) { range.setFontLine('none'); appliedCount++; }
    
    // Font size
    if (formatting.fontSize) { range.setFontSize(formatting.fontSize); appliedCount++; }
    
    // Font family
    if (formatting.fontFamily) { range.setFontFamily(formatting.fontFamily); appliedCount++; }
    
    // ===== COLORS =====
    if (formatting.backgroundColor || formatting.background) {
      range.setBackground(formatting.backgroundColor || formatting.background);
      appliedCount++;
    }
    if (formatting.textColor || formatting.fontColor || formatting.color) {
      range.setFontColor(formatting.textColor || formatting.fontColor || formatting.color);
      appliedCount++;
    }
    
    // ===== ALIGNMENT =====
    if (formatting.horizontalAlignment || formatting.alignment || formatting.align) {
      range.setHorizontalAlignment(formatting.horizontalAlignment || formatting.alignment || formatting.align);
      appliedCount++;
    }
    if (formatting.verticalAlignment || formatting.verticalAlign) {
      range.setVerticalAlignment(formatting.verticalAlignment || formatting.verticalAlign);
      appliedCount++;
    }
    
    // ===== TEXT WRAP =====
    if (formatting.wrap !== undefined) {
      Logger.log('[SheetActions_Format] 🔄 Setting wrap=' + formatting.wrap + ' on ' + rangeA1);
      range.setWrap(formatting.wrap);
      appliedCount++;
    }
    
    // Wrap strategy - more control than simple wrap
    if (formatting.wrapStrategy) {
      Logger.log('[SheetActions_Format] 🔄 Setting wrapStrategy=' + formatting.wrapStrategy + ' on ' + rangeA1);
      range.setWrapStrategy(SheetActions_Utils.getWrapStrategy(formatting.wrapStrategy));
      appliedCount++;
    }
    
    // ===== TEXT ROTATION =====
    if (formatting.textRotation !== undefined) {
      var degrees = Number(formatting.textRotation);
      Logger.log('[SheetActions_Format] 🔄 Setting textRotation=' + degrees + '° on ' + rangeA1);
      range.setTextRotation(degrees);
      appliedCount++;
    }
    
    // ===== TEXT DIRECTION =====
    if (formatting.textDirection) {
      var direction = formatting.textDirection.toUpperCase();
      if (direction === 'RTL' || direction === 'RIGHT_TO_LEFT') {
        range.setTextDirection(SpreadsheetApp.TextDirection.RIGHT_TO_LEFT);
        appliedCount++;
      } else if (direction === 'LTR' || direction === 'LEFT_TO_RIGHT') {
        range.setTextDirection(SpreadsheetApp.TextDirection.LEFT_TO_RIGHT);
        appliedCount++;
      }
    }
    
    // Log if nothing was applied (indicates parameter routing issue)
    if (appliedCount === 0 && !numberFormatAlreadyApplied) {
      Logger.log('[SheetActions_Format] ⚠️ WARNING: No formatting properties matched on ' + rangeA1 + '. Formatting keys received: ' + Object.keys(formatting).join(', '));
    } else {
      Logger.log('[SheetActions_Format] Applied ' + appliedCount + ' formatting properties to ' + rangeA1);
    }
    
    // ===== CELL MERGING (NEW) =====
    if (formatting.merge === true) {
      range.merge();
    } else if (formatting.merge === 'across') {
      range.mergeAcross();
    } else if (formatting.merge === 'vertical') {
      range.mergeVertically();
    }
    if (formatting.unmerge) {
      range.breakApart();
    }
    
    // ===== BORDERS =====
    if (formatting.borders) {
      _applyBorders(range, formatting);
    }
    
    // ===== ROW BANDING / ALTERNATING COLORS (NEW) =====
    if (formatting.banding) {
      var bandingTheme = SheetActions_Utils.getBandingTheme(formatting.bandingTheme || 'LIGHT_GREY');
      range.applyRowBanding(bandingTheme, formatting.showHeader !== false, formatting.showFooter === true);
    }
    
    // ===== NOTES/COMMENTS (NEW) =====
    if (formatting.note) {
      range.setNote(formatting.note);
    }
    if (formatting.clearNote) {
      range.clearNote();
    }
    
    // ===== NUMBER FORMAT =====
    // Only apply if not already handled by the caller (prevents double-apply which can corrupt formats)
    if (!numberFormatAlreadyApplied && (formatting.numberFormat || formatting.formatType)) {
      var nfType = formatting.formatType || formatting.numberFormat;
      var nfStr = SheetActions_Utils.getNumberFormat(nfType, formatting);
      range.setNumberFormat(nfStr);
      Logger.log('[SheetActions_Format] Applied number format from options: ' + nfStr + ' (from: ' + nfType + ')');
    }
    
    // ===== AUTO-RESIZE COLUMNS (post-format convenience) =====
    if (formatting.autoFitColumns || formatting.autoResize) {
      try {
        var sheet = range.getSheet();
        var startCol = range.getColumn();
        var numCols = range.getNumColumns();
        for (var c = startCol; c < startCol + numCols; c++) {
          sheet.autoResizeColumn(c);
        }
        Logger.log('[SheetActions_Format] Auto-resized ' + numCols + ' columns');
      } catch (e) {
        Logger.log('[SheetActions_Format] Auto-resize failed: ' + e.message);
      }
    }
    
    // ===== SET COLUMN WIDTH =====
    if (formatting.columnWidth) {
      try {
        var sheet = range.getSheet();
        var startCol = range.getColumn();
        var numCols = range.getNumColumns();
        for (var c = startCol; c < startCol + numCols; c++) {
          sheet.setColumnWidth(c, formatting.columnWidth);
        }
        Logger.log('[SheetActions_Format] Set column width to ' + formatting.columnWidth + 'px');
      } catch (e) {
        Logger.log('[SheetActions_Format] Set column width failed: ' + e.message);
      }
    }
    
    // ===== SET ROW HEIGHT =====
    if (formatting.rowHeight) {
      try {
        var sheet = range.getSheet();
        var startRow = range.getRow();
        var numRows = range.getNumRows();
        for (var r = startRow; r < startRow + numRows; r++) {
          sheet.setRowHeight(r, formatting.rowHeight);
        }
        Logger.log('[SheetActions_Format] Set row height to ' + formatting.rowHeight + 'px');
      } catch (e) {
        Logger.log('[SheetActions_Format] Set row height failed: ' + e.message);
      }
    }
    
    // ===== REMOVE BANDING =====
    if (formatting.removeBanding || formatting.clearBanding) {
      try {
        var bandings = range.getSheet().getBandings();
        for (var b = 0; b < bandings.length; b++) {
          var bandRange = bandings[b].getRange();
          if (bandRange.getA1Notation() === range.getA1Notation()) {
            bandings[b].remove();
            Logger.log('[SheetActions_Format] Removed banding from ' + range.getA1Notation());
          }
        }
      } catch (e) {
        Logger.log('[SheetActions_Format] Remove banding failed: ' + e.message);
      }
    }
  }
  
  /**
   * Apply borders to a range
   * @param {Range} range - Google Sheets Range object
   * @param {Object} formatting - Formatting options with border settings
   */
  function _applyBorders(range, formatting) {
    // Simple borders (all sides, default style)
    if (formatting.borders === true) {
      range.setBorder(true, true, true, true, true, true);
      return;
    }
    
    // Advanced borders with color and style (NEW)
    var borderColor = formatting.borderColor || '#000000';
    var borderStyle = SheetActions_Utils.getBorderStyle(formatting.borderStyle || 'solid');
    
    // Individual side control
    var top = formatting.borderTop !== false;
    var left = formatting.borderLeft !== false;
    var bottom = formatting.borderBottom !== false;
    var right = formatting.borderRight !== false;
    var vertical = formatting.borderVertical !== false;
    var horizontal = formatting.borderHorizontal !== false;
    
    range.setBorder(top, left, bottom, right, vertical, horizontal, borderColor, borderStyle);
  }
  
  // ============================================
  // PUBLIC API
  // ============================================
  
  return {
    applyFormat: applyFormat
  };
  
})();
