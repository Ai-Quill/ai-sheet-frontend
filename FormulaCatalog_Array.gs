/**
 * @file FormulaCatalog_Array.gs
 * @description Array, Dynamic & LAMBDA Functions for Google Sheets
 * 
 * COMPLETE COVERAGE of modern array functions including LAMBDA family.
 * These are the most powerful functions in Google Sheets.
 */

var ARRAY_FUNCTIONS = {
  
  // ========== FILTERING ==========
  
  FILTER: {
    syntax: 'FILTER(range, condition1, [condition2], ...)',
    description: 'Returns rows that meet all conditions',
    params: {
      range: 'Range to filter',
      condition1: 'Boolean array or condition',
      condition2: 'Additional conditions (AND logic)'
    },
    examples: [
      'FILTER(A:C, B:B>100)',
      'FILTER(A:C, A:A="Apple", B:B>100)',  // Multiple conditions
      'FILTER(A:B, REGEXMATCH(A:A, "^[A-M]"))',  // Regex filter
      'FILTER(Data, (Region="West")*(Sales>1000))',  // Alternative AND
      'FILTER(A:B, (A:A="X")+(A:A="Y"))'  // OR using addition
    ],
    tips: [
      'Multiple conditions are ANDed together',
      'Use + for OR: FILTER(range, (cond1)+(cond2))',
      'Returns empty if no matches (use IFERROR)',
      'Results spill automatically'
    ],
    relatedFunctions: ['QUERY', 'UNIQUE', 'SORT']
  },
  
  // ========== SORTING ==========
  
  SORT: {
    syntax: 'SORT(range, sort_column, is_ascending, [sort_column2, is_ascending2], ...)',
    description: 'Sorts rows by specified columns',
    params: {
      range: 'Range to sort',
      sort_column: 'Column index (1-based) or range to sort by',
      is_ascending: 'TRUE for A-Z/1-9, FALSE for Z-A/9-1',
      sort_column2: 'Secondary sort column',
      is_ascending2: 'Secondary sort order'
    },
    examples: [
      'SORT(A:C, 1, TRUE)',  // Sort by column 1 ascending
      'SORT(A:C, 2, FALSE)',  // Sort by column 2 descending
      'SORT(A:C, 2, FALSE, 1, TRUE)',  // Multiple sort columns
      'SORT(FILTER(A:C, B:B>100), 2, FALSE)'  // Sort filtered results
    ],
    tips: [
      'Column index is relative to the range',
      'Can chain with FILTER',
      'Results spill automatically'
    ],
    relatedFunctions: ['SORTN', 'FILTER', 'UNIQUE']
  },
  
  SORTN: {
    syntax: 'SORTN(range, [n], [display_ties_mode], [sort_column, is_ascending], ...)',
    description: 'Returns top N rows after sorting',
    params: {
      range: 'Range to sort',
      n: 'Number of rows to return',
      display_ties_mode: '0=show n rows, 1=show all ties with n-th, 2=show at most n with all ties, 3=show n after removing duplicates',
      sort_column: 'Column to sort by',
      is_ascending: 'Sort order'
    },
    examples: [
      'SORTN(A:B, 10, 0, 2, FALSE)',  // Top 10 by column 2
      'SORTN(A:B, 5, 1, 2, FALSE)',   // Top 5 with ties
      'SORTN(A:B, 3, 3, 2, FALSE)'    // Top 3 unique values
    ],
    tips: [
      'Efficient for "Top N" calculations',
      'Mode 3 is great for unique rankings'
    ],
    relatedFunctions: ['SORT', 'LARGE', 'FILTER']
  },
  
  // ========== UNIQUE VALUES ==========
  
  UNIQUE: {
    syntax: 'UNIQUE(range, [by_column], [exactly_once])',
    description: 'Returns unique rows from a range',
    params: {
      range: 'Range to get unique values from',
      by_column: 'Optional. TRUE to find unique columns instead of rows',
      exactly_once: 'Optional. TRUE to only return values appearing exactly once'
    },
    examples: [
      'UNIQUE(A:A)',  // Unique values in column
      'UNIQUE(A:C)',  // Unique rows across columns
      'UNIQUE(A:A, FALSE, TRUE)',  // Values appearing only once
      'SORT(UNIQUE(A:A))'  // Sorted unique values
    ],
    tips: [
      'First occurrence is kept',
      'Case-sensitive',
      'Works with multiple columns'
    ],
    relatedFunctions: ['COUNTUNIQUE', 'FILTER', 'SORT']
  },
  
  // ========== ARRAY CREATION ==========
  
  ARRAYFORMULA: {
    syntax: 'ARRAYFORMULA(array_formula)',
    description: 'Enables array calculations and spilling results',
    params: {
      array_formula: 'Formula to apply across arrays'
    },
    examples: [
      'ARRAYFORMULA(A:A * B:B)',  // Multiply columns
      'ARRAYFORMULA(IF(A:A="", "", A:A & " - " & B:B))',  // Concatenate columns
      'ARRAYFORMULA(YEAR(A:A))',  // Apply function to column
      'ARRAYFORMULA(VLOOKUP(A2:A, Data, 2, FALSE))'  // Lookup entire column
    ],
    tips: [
      'Essential for applying formulas to entire columns',
      'Avoids copying formulas down',
      'Use with IF to handle empty cells'
    ],
    relatedFunctions: ['MAP', 'BYROW', 'BYCOL']
  },
  
  MAKEARRAY: {
    syntax: 'MAKEARRAY(rows, columns, LAMBDA(row, col, calculation))',
    description: 'Creates an array using a LAMBDA function',
    params: {
      rows: 'Number of rows',
      columns: 'Number of columns',
      LAMBDA: 'Function to generate each cell value'
    },
    examples: [
      'MAKEARRAY(5, 5, LAMBDA(r, c, r*c))',  // Multiplication table
      'MAKEARRAY(10, 1, LAMBDA(r, c, r^2))',  // Squares 1-10
      'MAKEARRAY(12, 7, LAMBDA(r, c, DATE(2024, r, c)))'  // Calendar grid
    ],
    tips: [
      'Row and column indices are 1-based',
      'Powerful for generating structured data'
    ],
    relatedFunctions: ['SEQUENCE', 'LAMBDA', 'MAP']
  },
  
  // ========== LAMBDA FAMILY ==========
  
  LAMBDA: {
    syntax: 'LAMBDA([name1], [name2], ..., calculation)',
    description: 'Creates a custom function',
    params: {
      name1: 'Parameter name',
      name2: 'Additional parameter names',
      calculation: 'Formula using the parameters'
    },
    examples: [
      'LAMBDA(x, x*2)(5)',  // Returns 10
      'LAMBDA(a, b, a+b)(3, 4)',  // Returns 7
      'MAP(A1:A10, LAMBDA(x, x*2))',  // Apply to each cell
      // Named function style (in named ranges):
      // LAMBDA(price, tax, price*(1+tax))
    ],
    tips: [
      'Define reusable formulas',
      'Use with MAP, REDUCE, SCAN for iteration',
      'Can create complex custom functions',
      'Last argument is always the calculation'
    ],
    relatedFunctions: ['MAP', 'REDUCE', 'SCAN', 'LET']
  },
  
  LET: {
    syntax: 'LET(name1, value1, [name2, value2], ..., calculation)',
    description: 'Assigns names to values for use in calculation',
    params: {
      name1: 'Name for first value',
      value1: 'First value or expression',
      calculation: 'Final calculation using the names'
    },
    examples: [
      'LET(x, A1, y, B1, x+y)',
      'LET(total, SUM(A:A), count, COUNTA(A:A), total/count)',
      'LET(data, FILTER(A:B, A:A>100), SORT(data, 2, FALSE))',
      'LET(rate, 0.08, price, A1, price*(1+rate))'
    ],
    tips: [
      'Improves readability',
      'Avoids recalculating same expression',
      'Can nest multiple assignments',
      'Great for complex formulas'
    ],
    relatedFunctions: ['LAMBDA', 'IF']
  },
  
  MAP: {
    syntax: 'MAP(array1, [array2], ..., LAMBDA)',
    description: 'Applies a LAMBDA function to each element',
    params: {
      array1: 'First array',
      array2: 'Additional arrays (same size)',
      LAMBDA: 'Function to apply to each element'
    },
    examples: [
      'MAP(A1:A10, LAMBDA(x, x*2))',  // Double each value
      'MAP(A1:A10, B1:B10, LAMBDA(a, b, a+b))',  // Add corresponding elements
      'MAP(A1:A10, LAMBDA(x, IF(x>100, "High", "Low")))',  // Classify
      'MAP(A1:A10, LAMBDA(cell, LEN(cell)))'  // Length of each
    ],
    tips: [
      'Like ARRAYFORMULA but more flexible',
      'Can use multiple arrays (must be same size)',
      'LAMBDA params match number of arrays'
    ],
    relatedFunctions: ['ARRAYFORMULA', 'BYROW', 'BYCOL', 'LAMBDA']
  },
  
  REDUCE: {
    syntax: 'REDUCE(initial_value, array, LAMBDA)',
    description: 'Reduces array to single value using accumulator',
    params: {
      initial_value: 'Starting value for accumulator',
      array: 'Array to reduce',
      LAMBDA: 'Function(accumulator, current_value)'
    },
    examples: [
      'REDUCE(0, A1:A10, LAMBDA(acc, val, acc+val))',  // Sum
      'REDUCE(1, A1:A10, LAMBDA(acc, val, acc*val))',  // Product
      'REDUCE("", A1:A10, LAMBDA(acc, val, acc&val&", "))',  // Concatenate
      'REDUCE(0, A1:A10, LAMBDA(acc, val, MAX(acc, val)))'  // Maximum
    ],
    tips: [
      'Like JavaScript Array.reduce()',
      'Accumulator carries through iterations',
      'Great for custom aggregations'
    ],
    relatedFunctions: ['SCAN', 'MAP', 'LAMBDA']
  },
  
  SCAN: {
    syntax: 'SCAN(initial_value, array, LAMBDA)',
    description: 'Like REDUCE but returns all intermediate values',
    params: {
      initial_value: 'Starting value',
      array: 'Array to scan',
      LAMBDA: 'Function(accumulator, current_value)'
    },
    examples: [
      'SCAN(0, A1:A10, LAMBDA(acc, val, acc+val))',  // Running total
      'SCAN(1, A1:A10, LAMBDA(acc, val, acc*val))',  // Running product
      'SCAN(100, B1:B10, LAMBDA(bal, change, bal+change))'  // Running balance
    ],
    tips: [
      'Returns array of same size as input',
      'Great for running totals/cumulative calculations'
    ],
    relatedFunctions: ['REDUCE', 'MAP', 'LAMBDA']
  },
  
  BYROW: {
    syntax: 'BYROW(array, LAMBDA)',
    description: 'Applies LAMBDA to each row, returns column of results',
    params: {
      array: 'Two-dimensional array',
      LAMBDA: 'Function receiving each row as array'
    },
    examples: [
      'BYROW(A1:C10, LAMBDA(row, SUM(row)))',  // Sum each row
      'BYROW(A1:C10, LAMBDA(row, MAX(row)))',  // Max of each row
      'BYROW(A1:C10, LAMBDA(row, TEXTJOIN(",", TRUE, row)))'  // Join row values
    ],
    tips: [
      'LAMBDA receives entire row as array',
      'Returns single column of results',
      'Row operations that return single value'
    ],
    relatedFunctions: ['BYCOL', 'MAP', 'LAMBDA']
  },
  
  BYCOL: {
    syntax: 'BYCOL(array, LAMBDA)',
    description: 'Applies LAMBDA to each column, returns row of results',
    params: {
      array: 'Two-dimensional array',
      LAMBDA: 'Function receiving each column as array'
    },
    examples: [
      'BYCOL(A1:C10, LAMBDA(col, SUM(col)))',  // Sum each column
      'BYCOL(A1:C10, LAMBDA(col, AVERAGE(col)))',  // Average of each column
      'BYCOL(A1:C10, LAMBDA(col, COUNTA(col)))'  // Count per column
    ],
    tips: [
      'LAMBDA receives entire column as array',
      'Returns single row of results',
      'Column operations that return single value'
    ],
    relatedFunctions: ['BYROW', 'MAP', 'LAMBDA']
  },
  
  // ========== ARRAY MANIPULATION ==========
  
  HSTACK: {
    syntax: 'HSTACK(range1, [range2], ...)',
    description: 'Horizontally stacks arrays (side by side)',
    params: {
      range1: 'First array',
      range2: 'Additional arrays'
    },
    examples: [
      'HSTACK(A1:A10, B1:B10)',  // Combine two columns
      'HSTACK(A:A, SEQUENCE(COUNTA(A:A)))',  // Add row numbers
      'HSTACK(UNIQUE(A:A), COUNTIF(A:A, UNIQUE(A:A)))'  // Unique with counts
    ],
    tips: [
      'Arrays must have same number of rows',
      'Great for building result tables'
    ],
    relatedFunctions: ['VSTACK', 'CHOOSECOLS']
  },
  
  VSTACK: {
    syntax: 'VSTACK(range1, [range2], ...)',
    description: 'Vertically stacks arrays (one below another)',
    params: {
      range1: 'First array',
      range2: 'Additional arrays'
    },
    examples: [
      'VSTACK(A1:C5, A10:C15)',  // Combine ranges vertically
      'VSTACK({"Header1","Header2"}, A1:B10)',  // Add header
      'VSTACK(Sheet1!A:B, Sheet2!A:B)'  // Combine from sheets
    ],
    tips: [
      'Arrays must have same number of columns',
      'Null/missing values shown as empty'
    ],
    relatedFunctions: ['HSTACK', 'CHOOSEROWS']
  },
  
  TRANSPOSE: {
    syntax: 'TRANSPOSE(array_or_range)',
    description: 'Swaps rows and columns',
    params: {
      array_or_range: 'Array to transpose'
    },
    examples: [
      'TRANSPOSE(A1:A10)',  // Column to row
      'TRANSPOSE(A1:C3)',   // 3x3 swap
      'TRANSPOSE(SPLIT(A1, ","))'  // Horizontal split to vertical
    ],
    tips: [
      'Rows become columns, columns become rows',
      'Useful for reformatting data'
    ],
    relatedFunctions: ['TOCOL', 'TOROW']
  },
  
  // ========== ARRAY INFO ==========
  
  ROWS: {
    syntax: 'ROWS(range)',
    description: 'Returns number of rows in range/array',
    params: {
      range: 'Range or array'
    },
    examples: [
      'ROWS(A1:A10)',  // Returns 10
      'ROWS(FILTER(A:B, A:A>100))'  // Count filtered rows
    ],
    relatedFunctions: ['COLUMNS', 'COUNTA']
  },
  
  COLUMNS: {
    syntax: 'COLUMNS(range)',
    description: 'Returns number of columns in range/array',
    params: {
      range: 'Range or array'
    },
    examples: [
      'COLUMNS(A1:D1)',  // Returns 4
      'COLUMNS(Data)'    // Named range width
    ],
    relatedFunctions: ['ROWS', 'COUNTA']
  },
  
  // ========== ERROR HANDLING FOR ARRAYS ==========
  
  IFNA: {
    // Also in Logic, included here for array context
    syntax: 'IFNA(value, value_if_na)',
    description: 'Returns alternative value when #N/A error occurs (crucial for empty FILTER results)',
    params: {
      value: 'Value to check',
      value_if_na: 'Alternative if #N/A'
    },
    examples: [
      'IFNA(FILTER(A:B, A:A="X"), "No matches")',
      'IFNA(XLOOKUP(A1, B:B, C:C), "Not found")'
    ],
    tips: [
      'FILTER returns #N/A when no matches',
      'XLOOKUP returns #N/A when not found'
    ],
    relatedFunctions: ['IFERROR', 'FILTER']
  },
  
  // ========== ARRAY FORMULAS WITH CONDITIONS ==========
  
  SUMPRODUCT: {
    // Also in Math, but crucial for array conditions
    syntax: 'SUMPRODUCT((condition1)*(condition2)*values)',
    description: 'Sum with multiple conditions using boolean multiplication',
    params: {
      condition1: 'First condition (boolean array)',
      condition2: 'Additional conditions',
      values: 'Values to sum'
    },
    examples: [
      'SUMPRODUCT((A:A="Apple")*(B:B>100)*(C:C))',  // Sum C where A="Apple" AND B>100
      'SUMPRODUCT((A:A="X")*(B:B))/SUMPRODUCT((A:A="X")*1)'  // Weighted average
    ],
    tips: [
      'Boolean TRUE becomes 1, FALSE becomes 0',
      'Multiplying booleans = AND logic',
      'Legacy alternative to SUMIFS for complex conditions'
    ],
    relatedFunctions: ['SUMIFS', 'FILTER']
  },
  
  // ========== CROSS JOIN / COMBINATIONS ==========
  
  CROSSJOIN_PATTERN: {
    description: 'Pattern for creating cross join (all combinations)',
    formula: 'TOCOL(MAKEARRAY(ROWS(list1), ROWS(list2), LAMBDA(r,c, INDEX(list1,r)&","&INDEX(list2,c))))',
    examples: [
      // All combinations of products and regions
      'LET(products, A:A, regions, B:B, TOCOL(MAKEARRAY(COUNTA(products), COUNTA(regions), LAMBDA(r,c, INDEX(products,r)&" - "&INDEX(regions,c)))))'
    ],
    tips: [
      'Useful for generating all combinations',
      'Can split result back with SPLIT'
    ]
  }
};

/**
 * Get array function info
 */
function getArrayFunction(name) {
  return ARRAY_FUNCTIONS[name.toUpperCase()] || null;
}

/**
 * Get all array function names
 */
function getArrayFunctionNames() {
  return Object.keys(ARRAY_FUNCTIONS);
}

Logger.log('📦 FormulaCatalog_Array loaded - ' + Object.keys(ARRAY_FUNCTIONS).length + ' functions');
