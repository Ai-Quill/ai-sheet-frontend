/**
 * @file FormulaCatalog_Lookup.gs
 * @description Lookup & Reference Functions for Google Sheets
 * 
 * COMPLETE COVERAGE of all lookup/reference functions in Google Sheets.
 * These functions help you find, retrieve, and reference data.
 */

var LOOKUP_FUNCTIONS = {
  
  // ========== VERTICAL/HORIZONTAL LOOKUPS ==========
  
  VLOOKUP: {
    syntax: 'VLOOKUP(search_key, range, index, [is_sorted])',
    description: 'Searches down the first column of a range for a key and returns a value from another column',
    params: {
      search_key: 'The value to search for',
      range: 'The range to search',
      index: 'The column index of the value to return (1-indexed)',
      is_sorted: 'Optional. FALSE for exact match (default), TRUE for approximate'
    },
    examples: [
      'VLOOKUP(A2, Products!A:C, 2, FALSE)',
      'VLOOKUP("Apple", A:B, 2, FALSE)'
    ],
    tips: [
      'Use FALSE for exact matches (most common)',
      'Search column must be the leftmost column',
      'Consider INDEX/MATCH for more flexibility',
      'Use IFERROR to handle missing values'
    ],
    relatedFunctions: ['HLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH']
  },
  
  HLOOKUP: {
    syntax: 'HLOOKUP(search_key, range, index, [is_sorted])',
    description: 'Searches across the first row of a range for a key and returns a value from another row',
    params: {
      search_key: 'The value to search for',
      range: 'The range to search',
      index: 'The row index of the value to return (1-indexed)',
      is_sorted: 'Optional. FALSE for exact match, TRUE for approximate'
    },
    examples: [
      'HLOOKUP("Q1", A1:D4, 2, FALSE)',
      'HLOOKUP(B1, Headers!1:10, 3, FALSE)'
    ],
    tips: [
      'Use for horizontal data layouts',
      'Search row must be the topmost row'
    ],
    relatedFunctions: ['VLOOKUP', 'INDEX', 'MATCH']
  },
  
  XLOOKUP: {
    syntax: 'XLOOKUP(search_key, lookup_range, result_range, [missing_value], [match_mode], [search_mode])',
    description: 'Modern lookup function - more flexible than VLOOKUP',
    params: {
      search_key: 'The value to search for',
      lookup_range: 'The range to search in',
      result_range: 'The range containing the result to return',
      missing_value: 'Optional. Value if not found (default: #N/A)',
      match_mode: 'Optional. 0=exact, -1=exact or smaller, 1=exact or larger, 2=wildcard',
      search_mode: 'Optional. 1=first to last, -1=last to first, 2=binary ascending, -2=binary descending'
    },
    examples: [
      'XLOOKUP(A2, Products!A:A, Products!B:B, "Not Found")',
      'XLOOKUP("*phone*", A:A, B:B, "", 2)'  // Wildcard search
    ],
    tips: [
      'Can look up to the left (unlike VLOOKUP)',
      'Built-in error handling with missing_value',
      'Supports wildcard matching',
      'More readable than INDEX/MATCH'
    ],
    relatedFunctions: ['VLOOKUP', 'INDEX', 'MATCH', 'FILTER']
  },
  
  LOOKUP: {
    syntax: 'LOOKUP(search_key, search_range, [result_range])',
    description: 'Looks for a key in a range and returns the corresponding value',
    params: {
      search_key: 'The value to search for',
      search_range: 'The range to search (must be sorted)',
      result_range: 'Optional. The range to return from'
    },
    examples: [
      'LOOKUP(100, A:A, B:B)',
      'LOOKUP(2, {1,2,3}, {"a","b","c"})'
    ],
    tips: [
      'Data must be sorted ascending',
      'Returns the last value less than or equal to search_key'
    ],
    relatedFunctions: ['VLOOKUP', 'HLOOKUP', 'XLOOKUP']
  },
  
  // ========== INDEX & MATCH ==========
  
  INDEX: {
    syntax: 'INDEX(reference, [row], [column])',
    description: 'Returns the content of a cell at the intersection of a row and column',
    params: {
      reference: 'The range or array',
      row: 'Optional. The row number (1-indexed)',
      column: 'Optional. The column number (1-indexed)'
    },
    examples: [
      'INDEX(A1:C10, 5, 2)',
      'INDEX(A:A, MATCH("Apple", B:B, 0))',  // INDEX/MATCH combo
      'INDEX(Data, 0, 3)'  // Entire column 3
    ],
    tips: [
      'Use 0 for row or column to return entire row/column',
      'Combine with MATCH for powerful lookups',
      'Can return arrays, not just single values'
    ],
    relatedFunctions: ['MATCH', 'VLOOKUP', 'OFFSET']
  },
  
  MATCH: {
    syntax: 'MATCH(search_key, range, [search_type])',
    description: 'Returns the relative position of a value in a range',
    params: {
      search_key: 'The value to search for',
      range: 'The range to search (single row or column)',
      search_type: 'Optional. 0=exact, 1=less than or equal, -1=greater than or equal'
    },
    examples: [
      'MATCH("Apple", A:A, 0)',
      'MATCH(100, B:B, 1)',  // Largest value <= 100
      'INDEX(C:C, MATCH(A2, B:B, 0))'  // With INDEX
    ],
    tips: [
      'Use 0 for exact match (most common)',
      'Returns position, not value - combine with INDEX',
      'For search_type 1 or -1, data must be sorted'
    ],
    relatedFunctions: ['INDEX', 'XMATCH', 'SEARCH', 'FIND']
  },
  
  XMATCH: {
    syntax: 'XMATCH(search_key, lookup_range, [match_mode], [search_mode])',
    description: 'Modern version of MATCH with more options',
    params: {
      search_key: 'The value to search for',
      lookup_range: 'The range to search',
      match_mode: 'Optional. 0=exact, -1=exact or smaller, 1=exact or larger, 2=wildcard',
      search_mode: 'Optional. 1=first to last, -1=last to first, 2=binary asc, -2=binary desc'
    },
    examples: [
      'XMATCH("Apple", A:A)',
      'XMATCH("*phone*", A:A, 2)',  // Wildcard
      'INDEX(B:B, XMATCH(A2, A:A))'
    ],
    tips: [
      'Supports wildcard matching',
      'Can search from bottom to top',
      'More flexible than MATCH'
    ],
    relatedFunctions: ['MATCH', 'INDEX', 'XLOOKUP']
  },
  
  // ========== INDIRECT & ADDRESS ==========
  
  INDIRECT: {
    syntax: 'INDIRECT(cell_reference_as_string, [is_A1_notation])',
    description: 'Returns a reference specified by a text string',
    params: {
      cell_reference_as_string: 'A text string containing a cell reference',
      is_A1_notation: 'Optional. TRUE for A1 style (default), FALSE for R1C1'
    },
    examples: [
      'INDIRECT("A" & B1)',  // Dynamic column
      'INDIRECT(Sheet2!A1)',  // Reference from another cell
      'SUM(INDIRECT("A1:A" & ROW()))'  // Dynamic range
    ],
    tips: [
      'Great for dynamic references',
      'Can reference other sheets dynamically',
      'Volatile - recalculates on every change'
    ],
    relatedFunctions: ['ADDRESS', 'OFFSET', 'ROW', 'COLUMN']
  },
  
  ADDRESS: {
    syntax: 'ADDRESS(row, column, [absolute_mode], [use_a1_notation], [sheet])',
    description: 'Creates a cell reference as a string',
    params: {
      row: 'Row number',
      column: 'Column number',
      absolute_mode: 'Optional. 1=$A$1, 2=A$1, 3=$A1, 4=A1',
      use_a1_notation: 'Optional. TRUE for A1 style',
      sheet: 'Optional. Sheet name to include'
    },
    examples: [
      'ADDRESS(1, 1)',  // Returns "$A$1"
      'ADDRESS(ROW(), COLUMN(), 4)',  // Returns "A1" (relative)
      'INDIRECT(ADDRESS(B1, C1))'  // Dynamic cell access
    ],
    tips: [
      'Combine with INDIRECT for dynamic references',
      'Use absolute_mode 4 for relative references'
    ],
    relatedFunctions: ['INDIRECT', 'ROW', 'COLUMN']
  },
  
  // ========== OFFSET ==========
  
  OFFSET: {
    syntax: 'OFFSET(cell_reference, offset_rows, offset_columns, [height], [width])',
    description: 'Returns a range offset from a starting cell',
    params: {
      cell_reference: 'Starting point',
      offset_rows: 'Number of rows to offset (negative = up)',
      offset_columns: 'Number of columns to offset (negative = left)',
      height: 'Optional. Height of returned range',
      width: 'Optional. Width of returned range'
    },
    examples: [
      'OFFSET(A1, 2, 1)',  // Cell 2 down, 1 right from A1
      'SUM(OFFSET(A1, 0, 0, 10, 1))',  // Sum of A1:A10
      'OFFSET(A1, 0, 0, COUNTA(A:A), 1)'  // Dynamic range
    ],
    tips: [
      'Great for creating dynamic ranges',
      'Volatile function - use sparingly',
      'Consider INDEX for non-volatile alternative'
    ],
    relatedFunctions: ['INDEX', 'INDIRECT', 'ADDRESS']
  },
  
  // ========== ROW & COLUMN INFO ==========
  
  ROW: {
    syntax: 'ROW([cell_reference])',
    description: 'Returns the row number of a cell',
    params: {
      cell_reference: 'Optional. Cell to get row number of (default: current cell)'
    },
    examples: [
      'ROW()',  // Current row
      'ROW(A5)',  // Returns 5
      'ROW(A1:A10)'  // Returns array {1;2;3;...;10}
    ],
    tips: [
      'Useful for creating sequential numbers',
      'Combine with MOD for alternating patterns'
    ],
    relatedFunctions: ['COLUMN', 'ROWS', 'ADDRESS']
  },
  
  COLUMN: {
    syntax: 'COLUMN([cell_reference])',
    description: 'Returns the column number of a cell',
    params: {
      cell_reference: 'Optional. Cell to get column number of'
    },
    examples: [
      'COLUMN()',  // Current column
      'COLUMN(C5)',  // Returns 3
      'CHAR(64+COLUMN())'  // Returns column letter
    ],
    tips: [
      'Column A = 1, B = 2, etc.',
      'Combine with CHAR to get column letter'
    ],
    relatedFunctions: ['ROW', 'COLUMNS', 'ADDRESS']
  },
  
  ROWS: {
    syntax: 'ROWS(range)',
    description: 'Returns the number of rows in a range',
    params: {
      range: 'The range to count rows of'
    },
    examples: [
      'ROWS(A1:A10)',  // Returns 10
      'ROWS(Selection)'  // Named range
    ],
    relatedFunctions: ['COLUMNS', 'ROW', 'COUNTA']
  },
  
  COLUMNS: {
    syntax: 'COLUMNS(range)',
    description: 'Returns the number of columns in a range',
    params: {
      range: 'The range to count columns of'
    },
    examples: [
      'COLUMNS(A1:D1)',  // Returns 4
      'COLUMNS(Data)'  // Named range
    ],
    relatedFunctions: ['ROWS', 'COLUMN', 'COUNTA']
  },
  
  // ========== SPECIALIZED LOOKUPS ==========
  
  CHOOSECOLS: {
    syntax: 'CHOOSECOLS(array, col_num1, [col_num2], ...)',
    description: 'Returns specified columns from an array',
    params: {
      array: 'The source array',
      col_num1: 'First column to return',
      col_num2: 'Additional columns (optional)'
    },
    examples: [
      'CHOOSECOLS(A1:D10, 1, 3)',  // Columns 1 and 3
      'CHOOSECOLS(Data, 2, 4, 1)'  // Reorder columns
    ],
    tips: [
      'Negative numbers count from end',
      'Great for rearranging data'
    ],
    relatedFunctions: ['CHOOSEROWS', 'INDEX', 'FILTER']
  },
  
  CHOOSEROWS: {
    syntax: 'CHOOSEROWS(array, row_num1, [row_num2], ...)',
    description: 'Returns specified rows from an array',
    params: {
      array: 'The source array',
      row_num1: 'First row to return',
      row_num2: 'Additional rows (optional)'
    },
    examples: [
      'CHOOSEROWS(A1:D10, 1, 5, 10)',
      'CHOOSEROWS(Data, -1)'  // Last row
    ],
    tips: [
      'Negative numbers count from end',
      '-1 = last row, -2 = second to last'
    ],
    relatedFunctions: ['CHOOSECOLS', 'INDEX', 'FILTER']
  },
  
  TAKE: {
    syntax: 'TAKE(array, rows, [columns])',
    description: 'Returns specified number of rows/columns from start or end',
    params: {
      array: 'The source array',
      rows: 'Number of rows (positive=from start, negative=from end)',
      columns: 'Optional. Number of columns'
    },
    examples: [
      'TAKE(A1:D10, 5)',  // First 5 rows
      'TAKE(A1:D10, -3)',  // Last 3 rows
      'TAKE(A1:D10, 5, 2)'  // First 5 rows, 2 columns
    ],
    relatedFunctions: ['DROP', 'CHOOSEROWS', 'CHOOSECOLS']
  },
  
  DROP: {
    syntax: 'DROP(array, rows, [columns])',
    description: 'Drops specified number of rows/columns from start or end',
    params: {
      array: 'The source array',
      rows: 'Number of rows to drop (positive=from start, negative=from end)',
      columns: 'Optional. Number of columns to drop'
    },
    examples: [
      'DROP(A1:D10, 1)',  // Drop header row
      'DROP(A1:D10, 0, 1)',  // Drop first column
      'DROP(A1:D10, -2)'  // Drop last 2 rows
    ],
    relatedFunctions: ['TAKE', 'CHOOSEROWS', 'CHOOSECOLS']
  },
  
  EXPAND: {
    syntax: 'EXPAND(array, rows, [columns], [pad_with])',
    description: 'Expands an array to specified dimensions',
    params: {
      array: 'The source array',
      rows: 'Target number of rows',
      columns: 'Optional. Target number of columns',
      pad_with: 'Optional. Value to fill empty cells (default: #N/A)'
    },
    examples: [
      'EXPAND(A1:B2, 5, 3)',  // Expand to 5x3
      'EXPAND({1,2,3}, 5, 3, 0)'  // Pad with zeros
    ],
    relatedFunctions: ['WRAPCOLS', 'WRAPROWS', 'TOCOL', 'TOROW']
  },
  
  // ========== ARRAY RESHAPING ==========
  
  TOCOL: {
    syntax: 'TOCOL(array, [ignore], [scan_by_column])',
    description: 'Converts an array to a single column',
    params: {
      array: 'The array to convert',
      ignore: 'Optional. 0=keep all, 1=ignore blanks, 2=ignore errors, 3=ignore both',
      scan_by_column: 'Optional. TRUE to scan by column, FALSE by row'
    },
    examples: [
      'TOCOL(A1:C3)',  // Flatten to column
      'TOCOL(A1:C3, 1)',  // Ignore blanks
      'UNIQUE(TOCOL(A1:C3, 1))'  // Unique values from range
    ],
    relatedFunctions: ['TOROW', 'FLATTEN', 'WRAPCOLS']
  },
  
  TOROW: {
    syntax: 'TOROW(array, [ignore], [scan_by_column])',
    description: 'Converts an array to a single row',
    params: {
      array: 'The array to convert',
      ignore: 'Optional. 0=keep all, 1=ignore blanks, 2=ignore errors, 3=ignore both',
      scan_by_column: 'Optional. TRUE to scan by column, FALSE by row'
    },
    examples: [
      'TOROW(A1:A10)',  // Convert column to row
      'TOROW(A1:C3, 1)'  // Flatten, ignore blanks
    ],
    relatedFunctions: ['TOCOL', 'FLATTEN', 'WRAPROWS']
  },
  
  WRAPCOLS: {
    syntax: 'WRAPCOLS(range, wrap_count, [pad_with])',
    description: 'Wraps a row/column into multiple columns',
    params: {
      range: 'The source range',
      wrap_count: 'Number of values per column',
      pad_with: 'Optional. Value to fill empty cells'
    },
    examples: [
      'WRAPCOLS(A1:A12, 4)',  // 12 values into 4 rows x 3 cols
      'WRAPCOLS(TOCOL(A1:C3), 3)'  // Reshape data
    ],
    relatedFunctions: ['WRAPROWS', 'TOCOL', 'TOROW']
  },
  
  WRAPROWS: {
    syntax: 'WRAPROWS(range, wrap_count, [pad_with])',
    description: 'Wraps a row/column into multiple rows',
    params: {
      range: 'The source range',
      wrap_count: 'Number of values per row',
      pad_with: 'Optional. Value to fill empty cells'
    },
    examples: [
      'WRAPROWS(A1:A12, 3)',  // 12 values into 4 rows x 3 cols
      'WRAPROWS(1:1, 5)'  // Split row into groups of 5
    ],
    relatedFunctions: ['WRAPCOLS', 'TOCOL', 'TOROW']
  },
  
  FLATTEN: {
    syntax: 'FLATTEN(range1, [range2], ...)',
    description: 'Flattens all values from ranges into a single column',
    params: {
      range1: 'First range to flatten',
      range2: 'Additional ranges (optional)'
    },
    examples: [
      'FLATTEN(A1:C3)',
      'FLATTEN(A:A, C:C, E:E)',  // Multiple columns
      'UNIQUE(FLATTEN(A1:C10))'  // All unique values
    ],
    tips: [
      'Google Sheets specific function',
      'Great for combining multiple ranges'
    ],
    relatedFunctions: ['TOCOL', 'TOROW', 'UNIQUE']
  },
  
  // ========== OTHER REFERENCE FUNCTIONS ==========
  
  FORMULATEXT: {
    syntax: 'FORMULATEXT(cell)',
    description: 'Returns the formula in a cell as a text string',
    params: {
      cell: 'The cell containing a formula'
    },
    examples: [
      'FORMULATEXT(A1)',  // Shows formula in A1
      'IF(ISERROR(FORMULATEXT(A1)), "Value", "Formula")'
    ],
    tips: [
      'Returns #N/A if cell has no formula',
      'Useful for documentation'
    ],
    relatedFunctions: ['ISFORMULA', 'CELL']
  },
  
  HYPERLINK: {
    syntax: 'HYPERLINK(url, [link_label])',
    description: 'Creates a clickable hyperlink',
    params: {
      url: 'The URL to link to',
      link_label: 'Optional. Display text for the link'
    },
    examples: [
      'HYPERLINK("https://google.com", "Google")',
      'HYPERLINK("#gid=0&range=A1", "Go to A1")',  // Internal link
      'HYPERLINK("mailto:" & A1, "Email")'  // Email link
    ],
    tips: [
      'Can link to other sheets in same doc',
      'Use # prefix for internal links',
      'Supports mailto: and tel: protocols'
    ],
    relatedFunctions: ['IMPORTHTML', 'IMPORTXML']
  },
  
  GETPIVOTDATA: {
    syntax: 'GETPIVOTDATA(value_name, any_pivot_table_cell, [field_name1, item1], ...)',
    description: 'Extracts data from a pivot table',
    params: {
      value_name: 'Name of the value field to extract',
      any_pivot_table_cell: 'Any cell in the pivot table',
      field_name1: 'Field name to filter by',
      item1: 'Item value in that field'
    },
    examples: [
      'GETPIVOTDATA("Sales", A3)',
      'GETPIVOTDATA("Total", A3, "Region", "West", "Year", 2024)'
    ],
    tips: [
      'Field names must match exactly',
      'Useful for extracting specific pivot values'
    ],
    relatedFunctions: ['QUERY']
  }
};

/**
 * Get lookup function info
 */
function getLookupFunction(name) {
  return LOOKUP_FUNCTIONS[name.toUpperCase()] || null;
}

/**
 * Get all lookup function names
 */
function getLookupFunctionNames() {
  return Object.keys(LOOKUP_FUNCTIONS);
}

Logger.log('📦 FormulaCatalog_Lookup loaded - ' + Object.keys(LOOKUP_FUNCTIONS).length + ' functions');
