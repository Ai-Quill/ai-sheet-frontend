/**
 * @file SheetActions_Utils.gs
 * @version 1.0.0
 * @created 2026-01-30
 * @author AISheeter Team
 * 
 * ============================================
 * SHEET ACTIONS - Shared Utilities
 * ============================================
 * 
 * Common utility functions used across all sheet action modules.
 * Includes column conversion, range detection, and enum mappings.
 */

var SheetActions_Utils = (function() {
  
  // ============================================
  // CONSTANTS
  // ============================================
  
  // Default color palette - professional, accessible colors
  var DEFAULT_COLORS = [
    '#4285F4', // Google Blue
    '#EA4335', // Google Red
    '#FBBC04', // Google Yellow
    '#34A853', // Google Green
    '#FF6D01', // Orange
    '#46BDC6', // Teal
    '#7BAAF7', // Light Blue
    '#F07B72', // Light Red
    '#FCD04F', // Light Yellow
    '#81C995'  // Light Green
  ];
  
  // ============================================
  // COLUMN CONVERSION UTILITIES
  // ============================================
  
  /**
   * Convert column letter to number (A=1, B=2, AA=27, etc.)
   * @param {string} letter - Column letter(s)
   * @return {number} Column index (1-based)
   */
  function letterToColumn(letter) {
    var column = 0;
    var length = letter.length;
    for (var i = 0; i < length; i++) {
      column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
  }
  
  /**
   * Convert column index to letter (1=A, 2=B, 27=AA, etc.)
   * @param {number} col - Column index (1-based)
   * @return {string} Column letter(s)
   */
  function columnToLetter(col) {
    var letter = '';
    while (col > 0) {
      var mod = (col - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      col = Math.floor((col - 1) / 26);
    }
    return letter;
  }
  
  /**
   * Convert column letter or name to index
   * @param {string|number} column - Column letter or index
   * @return {number} Column index (1-based)
   */
  function getColumnIndex(column) {
    if (typeof column === 'number') return column;
    // Handle single letter columns (A, B, C...)
    if (column.length === 1 && /[A-Z]/i.test(column)) {
      return column.toUpperCase().charCodeAt(0) - 64;
    }
    // For longer strings, assume it's a column letter like "AA"
    return letterToColumn(column.toUpperCase());
  }
  
  // ============================================
  // RANGE DETECTION UTILITIES
  // ============================================
  
  /**
   * Auto-detect data range from sheet
   * @param {Sheet} sheet - Google Sheets Sheet object
   * @return {string} Range in A1 notation
   */
  function detectDataRange(sheet) {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return 'A1:A1';
    return 'A1:' + columnToLetter(lastCol) + lastRow;
  }
  
  /**
   * Detect the start and end rows containing data for given columns
   * @param {Sheet} sheet - Google Sheets Sheet object
   * @param {number} minCol - Minimum column index
   * @param {number} maxCol - Maximum column index
   * @return {Object} { startRow, endRow }
   */
  function detectDataBoundaries(sheet, minCol, maxCol) {
    var lastRow = sheet.getLastRow();
    var numCols = maxCol - minCol + 1;
    
    // Get all data in the column range
    var data = sheet.getRange(1, minCol, lastRow, numCols).getValues();
    
    var startRow = 1;
    var endRow = lastRow;
    
    // Find first row with data (this is likely the header)
    for (var i = 0; i < data.length; i++) {
      var hasData = data[i].some(function(cell) {
        return cell !== '' && cell !== null;
      });
      if (hasData) {
        startRow = i + 1;
        break;
      }
    }
    
    // Find last row with data
    for (var j = data.length - 1; j >= 0; j--) {
      var hasData = data[j].some(function(cell) {
        return cell !== '' && cell !== null;
      });
      if (hasData) {
        endRow = j + 1;
        break;
      }
    }
    
    Logger.log('[SheetActions_Utils] Data boundaries: rows ' + startRow + ' to ' + endRow);
    
    return {
      startRow: startRow,
      endRow: endRow
    };
  }
  
  // ============================================
  // ENUM MAPPINGS
  // ============================================
  
  /**
   * Get border style enum from string
   * @param {string} style - Border style name
   * @return {BorderStyle} SpreadsheetApp.BorderStyle enum value
   */
  function getBorderStyle(style) {
    var styles = {
      'solid': SpreadsheetApp.BorderStyle.SOLID,
      'solid_medium': SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
      'solid_thick': SpreadsheetApp.BorderStyle.SOLID_THICK,
      'dashed': SpreadsheetApp.BorderStyle.DASHED,
      'dotted': SpreadsheetApp.BorderStyle.DOTTED,
      'double': SpreadsheetApp.BorderStyle.DOUBLE
    };
    return styles[style] || SpreadsheetApp.BorderStyle.SOLID;
  }
  
  /**
   * Get banding theme enum from string
   * @param {string} theme - Theme name
   * @return {BandingTheme} SpreadsheetApp.BandingTheme enum value
   */
  function getBandingTheme(theme) {
    var themes = {
      'LIGHT_GREY': SpreadsheetApp.BandingTheme.LIGHT_GREY,
      'CYAN': SpreadsheetApp.BandingTheme.CYAN,
      'GREEN': SpreadsheetApp.BandingTheme.GREEN,
      'YELLOW': SpreadsheetApp.BandingTheme.YELLOW,
      'ORANGE': SpreadsheetApp.BandingTheme.ORANGE,
      'BLUE': SpreadsheetApp.BandingTheme.BLUE,
      'TEAL': SpreadsheetApp.BandingTheme.TEAL,
      'GREY': SpreadsheetApp.BandingTheme.GREY,
      'BROWN': SpreadsheetApp.BandingTheme.BROWN,
      'LIGHT_GREEN': SpreadsheetApp.BandingTheme.LIGHT_GREEN,
      'INDIGO': SpreadsheetApp.BandingTheme.INDIGO,
      'PINK': SpreadsheetApp.BandingTheme.PINK
    };
    return themes[theme] || SpreadsheetApp.BandingTheme.LIGHT_GREY;
  }
  
  /**
   * Get wrap strategy enum from string
   * @param {string} strategy - Strategy name
   * @return {WrapStrategy} SpreadsheetApp.WrapStrategy enum value
   */
  function getWrapStrategy(strategy) {
    var strategies = {
      'overflow': SpreadsheetApp.WrapStrategy.OVERFLOW,
      'wrap': SpreadsheetApp.WrapStrategy.WRAP,
      'clip': SpreadsheetApp.WrapStrategy.CLIP
    };
    return strategies[strategy] || SpreadsheetApp.WrapStrategy.WRAP;
  }
  
  /**
   * Get chart type enum from string
   * @param {string} type - Chart type name
   * @return {ChartType} Charts.ChartType enum value
   */
  function getChartType(type) {
    var types = {
      'bar': Charts.ChartType.BAR,
      'column': Charts.ChartType.COLUMN,
      'line': Charts.ChartType.LINE,
      'pie': Charts.ChartType.PIE,
      'area': Charts.ChartType.AREA,
      'scatter': Charts.ChartType.SCATTER,
      'combo': Charts.ChartType.COMBO,
      'histogram': Charts.ChartType.HISTOGRAM,
      'stepped': Charts.ChartType.STEPPED_AREA,
      'steppedarea': Charts.ChartType.STEPPED_AREA
    };
    return types[type] || Charts.ChartType.COLUMN;
  }
  
  // ============================================
  // NUMBER FORMAT UTILITIES
  // ============================================
  
  /**
   * Build number format string from type and options
   * @param {string} type - Format type (currency, percent, number, date, etc.)
   * @param {Object} options - Format options
   * @return {string} Google Sheets number format pattern
   */
  function getNumberFormat(type, options) {
    options = options || {};
    var decimals = options.decimals !== undefined ? options.decimals : 2;
    var decimalStr = decimals > 0 ? '.' + '0'.repeat(decimals) : '';
    
    // For date formatting, check both 'pattern' (from AI) and 'dateFormat' (legacy)
    var dateFormat = options.pattern || options.dateFormat || 'yyyy-mm-dd';
    
    // Currency symbol handling - support multiple currencies
    var currencyFormat = '$#,##0' + decimalStr; // Default USD
    var locale = (options.locale || options.currency || '').toUpperCase();
    
    var currencySymbols = {
      'EUR': '€#,##0' + decimalStr,
      'GBP': '£#,##0' + decimalStr,
      'JPY': '¥#,##0',  // JPY typically no decimals
      'CNY': '¥#,##0' + decimalStr,
      'KRW': '₩#,##0',  // KRW typically no decimals
      'INR': '₹#,##0' + decimalStr,
      'RUB': '₽#,##0' + decimalStr,
      'BRL': 'R$#,##0' + decimalStr,
      'CAD': 'CA$#,##0' + decimalStr,
      'AUD': 'A$#,##0' + decimalStr,
      'CHF': 'CHF #,##0' + decimalStr,
      'USD': '$#,##0' + decimalStr,
      'VND': '#,##0 ₫'
    };
    
    if (locale && currencySymbols[locale]) {
      currencyFormat = currencySymbols[locale];
    } else if (options.symbol) {
      // AI might provide a custom symbol
      currencyFormat = options.symbol + '#,##0' + decimalStr;
    }
    
    Logger.log('[SheetActions_Utils] getNumberFormat: type=' + type + ', dateFormat=' + dateFormat + ', decimals=' + decimals + ', locale=' + locale);
    
    var formats = {
      // Currency
      'currency': currencyFormat,
      'accounting': '_($* #,##0' + decimalStr + '_);_($* (#,##0' + decimalStr + ');_($* "-"??_);_(@_)',
      'financial': currencyFormat,
      
      // Percentages
      'percent': '0' + decimalStr + '%',
      'percentage': '0' + decimalStr + '%',
      
      // Numbers
      'number': '#,##0' + decimalStr,
      'integer': '#,##0',
      'decimal': '#,##0' + decimalStr,
      'plain_number': '0',  // No commas, no decimals (e.g., years: 2014)
      'year': '0',  // Alias for plain_number — shows 2014 not 2,014
      
      // Scientific
      'scientific': '0' + decimalStr + 'E+00',
      'engineering': '##0' + decimalStr + 'E+00',
      
      // Fractions
      'fraction': '# ?/?',
      'fraction_precise': '# ??/??',
      
      // Date & Time
      'date': dateFormat,
      'datetime': options.pattern || 'yyyy-mm-dd hh:mm:ss',
      'time': options.pattern || 'hh:mm:ss',
      'duration': options.pattern || '[h]:mm:ss',
      'date_short': 'mm/dd/yy',
      'date_long': 'mmmm d, yyyy',
      'date_iso': 'yyyy-mm-dd',
      'time_12h': 'h:mm AM/PM',
      'time_24h': 'HH:mm',
      
      // Text
      'text': '@',
      'plain': '@',
      
      // Special patterns
      'phone': '(###) ###-####',
      'zip': '00000',
      'zip_plus4': '00000-0000',
      'ssn': '000-00-0000'
    };
    
    // Normalize type to lowercase
    var normalizedType = (type || '').toLowerCase();
    
    // Check named formats first
    if (formats[normalizedType]) {
      return formats[normalizedType];
    }
    
    // If type looks like a raw Google Sheets format string (contains # or 0 patterns),
    // pass it through directly instead of defaulting to '@' (text)
    if (type && (/[#0]/.test(type) || type.indexOf('%') !== -1)) {
      Logger.log('[SheetActions_Utils] Passing through raw format string: ' + type);
      return type;
    }
    
    // Unknown format type — default to auto (don't force text '@' which corrupts numbers)
    Logger.log('[SheetActions_Utils] Unknown format type "' + type + '", using auto');
    return '#,##0.##';
  }
  
  // ============================================
  // PUBLIC API
  // ============================================
  
  return {
    // Constants
    DEFAULT_COLORS: DEFAULT_COLORS,
    
    // Column utilities
    letterToColumn: letterToColumn,
    columnToLetter: columnToLetter,
    getColumnIndex: getColumnIndex,
    
    // Range utilities
    detectDataRange: detectDataRange,
    detectDataBoundaries: detectDataBoundaries,
    
    // Enum mappings
    getBorderStyle: getBorderStyle,
    getBandingTheme: getBandingTheme,
    getWrapStrategy: getWrapStrategy,
    getChartType: getChartType,
    
    // Format utilities
    getNumberFormat: getNumberFormat
  };
  
})();
