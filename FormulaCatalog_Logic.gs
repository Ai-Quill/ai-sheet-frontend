/**
 * @file FormulaCatalog_Logic.gs
 * @description Logical & Information Functions for Google Sheets
 * 
 * COMPLETE COVERAGE of logical functions and type checking functions.
 */

var LOGIC_FUNCTIONS = {
  
  // ========== CONDITIONAL ==========
  
  IF: {
    syntax: 'IF(logical_expression, value_if_true, value_if_false)',
    description: 'Returns one value if condition is true, another if false',
    params: {
      logical_expression: 'Condition to test',
      value_if_true: 'Value if condition is TRUE',
      value_if_false: 'Value if condition is FALSE'
    },
    examples: [
      'IF(A1>100, "High", "Low")',
      'IF(A1="", "Empty", A1)',
      'IF(AND(A1>0, A1<100), "In range", "Out of range")'
    ],
    tips: [
      'Can be nested for multiple conditions',
      'Consider IFS for cleaner multiple conditions'
    ],
    relatedFunctions: ['IFS', 'SWITCH', 'AND', 'OR']
  },
  
  IFS: {
    syntax: 'IFS(condition1, value1, [condition2, value2], ...)',
    description: 'Checks multiple conditions and returns first match',
    params: {
      condition1: 'First condition',
      value1: 'Value if condition1 is TRUE',
      condition2: 'Second condition',
      value2: 'Value if condition2 is TRUE'
    },
    examples: [
      'IFS(A1>=90, "A", A1>=80, "B", A1>=70, "C", TRUE, "F")',
      'IFS(Status="Done", "✓", Status="In Progress", "→", TRUE, "○")'
    ],
    tips: [
      'Use TRUE as last condition for "else"',
      'Evaluates in order - first TRUE wins',
      'Cleaner than nested IF statements'
    ],
    relatedFunctions: ['IF', 'SWITCH']
  },
  
  SWITCH: {
    syntax: 'SWITCH(expression, case1, value1, [case2, value2], ..., [default])',
    description: 'Tests expression against list of cases, returns matching value',
    params: {
      expression: 'Value to test',
      case1: 'First case to match',
      value1: 'Value if expression equals case1',
      default: 'Optional. Default value if no match'
    },
    examples: [
      'SWITCH(A1, 1, "One", 2, "Two", 3, "Three", "Other")',
      'SWITCH(WEEKDAY(A1), 1, "Sun", 2, "Mon", 3, "Tue", 4, "Wed", 5, "Thu", 6, "Fri", 7, "Sat")',
      'SWITCH(Status, "A", "Active", "I", "Inactive", "P", "Pending", "Unknown")'
    ],
    tips: [
      'Like a lookup table in formula form',
      'Last odd argument is default value',
      'More readable than nested IFs for exact matches'
    ],
    relatedFunctions: ['IFS', 'IF', 'CHOOSE']
  },
  
  CHOOSE: {
    syntax: 'CHOOSE(index, choice1, [choice2], ...)',
    description: 'Returns value from list based on index',
    params: {
      index: 'Index number (1-based)',
      choice1: 'First choice',
      choice2: 'Additional choices'
    },
    examples: [
      'CHOOSE(MONTH(A1), "Jan", "Feb", "Mar", ...)',
      'CHOOSE(WEEKDAY(A1), "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")',
      'CHOOSE(A1, B1, C1, D1)'  // Return value from column based on A1
    ],
    tips: [
      'Index must be 1 to number of choices',
      'Can return ranges for use in other functions'
    ],
    relatedFunctions: ['SWITCH', 'INDEX']
  },
  
  // ========== BOOLEAN LOGIC ==========
  
  AND: {
    syntax: 'AND(logical1, [logical2], ...)',
    description: 'Returns TRUE if ALL arguments are TRUE',
    params: {
      logical1: 'First condition',
      logical2: 'Additional conditions'
    },
    examples: [
      'AND(A1>0, A1<100)',
      'IF(AND(A1>0, B1>0, C1>0), "All positive", "Not all positive")',
      'AND(A1:A10>0)'  // All values in range > 0
    ],
    tips: [
      'Can test arrays/ranges',
      'Use inside IF for multiple conditions'
    ],
    relatedFunctions: ['OR', 'NOT', 'IF']
  },
  
  OR: {
    syntax: 'OR(logical1, [logical2], ...)',
    description: 'Returns TRUE if ANY argument is TRUE',
    params: {
      logical1: 'First condition',
      logical2: 'Additional conditions'
    },
    examples: [
      'OR(A1="Yes", A1="Y", A1="TRUE")',
      'IF(OR(A1<0, A1>100), "Out of range", "OK")',
      'OR(A1:A10="Error")'  // Any cell contains "Error"
    ],
    relatedFunctions: ['AND', 'NOT', 'XOR']
  },
  
  NOT: {
    syntax: 'NOT(logical)',
    description: 'Returns opposite of logical value',
    params: {
      logical: 'Value or expression to negate'
    },
    examples: [
      'NOT(A1>100)',  // Same as A1<=100
      'NOT(ISBLANK(A1))',  // Cell is NOT blank
      'IF(NOT(ISERROR(A1)), A1, "Error")'
    ],
    relatedFunctions: ['AND', 'OR']
  },
  
  XOR: {
    syntax: 'XOR(logical1, [logical2], ...)',
    description: 'Returns TRUE if odd number of arguments are TRUE',
    params: {
      logical1: 'First condition',
      logical2: 'Additional conditions'
    },
    examples: [
      'XOR(A1>0, B1>0)',  // Exactly one is positive
      'XOR(TRUE, FALSE)',  // Returns TRUE
      'XOR(TRUE, TRUE)'   // Returns FALSE
    ],
    tips: [
      'True "exclusive or" when 2 arguments',
      'With more arguments: TRUE if odd count of TRUEs'
    ],
    relatedFunctions: ['AND', 'OR']
  },
  
  // ========== ERROR HANDLING ==========
  
  IFERROR: {
    syntax: 'IFERROR(value, value_if_error)',
    description: 'Returns value_if_error if value is an error, otherwise returns value',
    params: {
      value: 'Value to check',
      value_if_error: 'Value to return if error'
    },
    examples: [
      'IFERROR(A1/B1, 0)',
      'IFERROR(VLOOKUP(A1, Data, 2, FALSE), "Not found")',
      'IFERROR(1/0, "Cannot divide")'
    ],
    tips: [
      'Catches ALL error types',
      'Use IFNA to only catch #N/A'
    ],
    relatedFunctions: ['IFNA', 'ISERROR', 'ERROR.TYPE']
  },
  
  IFNA: {
    syntax: 'IFNA(value, value_if_na)',
    description: 'Returns value_if_na only if value is #N/A error',
    params: {
      value: 'Value to check',
      value_if_na: 'Value to return if #N/A'
    },
    examples: [
      'IFNA(VLOOKUP(A1, Data, 2, FALSE), "Not found")',
      'IFNA(MATCH(A1, B:B, 0), 0)'
    ],
    tips: [
      'Only catches #N/A, not other errors',
      'More precise than IFERROR for lookups'
    ],
    relatedFunctions: ['IFERROR', 'ISNA']
  },
  
  // ========== TYPE CHECKING ==========
  
  ISBLANK: {
    syntax: 'ISBLANK(value)',
    description: 'Returns TRUE if cell is empty',
    params: {
      value: 'Cell or value to check'
    },
    examples: [
      'ISBLANK(A1)',
      'IF(ISBLANK(A1), "Fill this", A1)',
      'COUNTIF(A:A, "")'  // Alternative for counting blanks
    ],
    tips: [
      'Cell with formula returning "" is NOT blank',
      'Cell with space is NOT blank'
    ],
    relatedFunctions: ['COUNTBLANK', 'IF']
  },
  
  ISERROR: {
    syntax: 'ISERROR(value)',
    description: 'Returns TRUE if value is any error type',
    params: {
      value: 'Value to check'
    },
    examples: [
      'ISERROR(A1/B1)',
      'IF(ISERROR(VLOOKUP(...)), "Not found", VLOOKUP(...))'
    ],
    tips: [
      'Catches all errors: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?',
      'Use IFERROR instead of IF(ISERROR(...),...)'
    ],
    relatedFunctions: ['IFERROR', 'ISNA', 'ISERR', 'ERROR.TYPE']
  },
  
  ISERR: {
    syntax: 'ISERR(value)',
    description: 'Returns TRUE if value is error EXCEPT #N/A',
    params: {
      value: 'Value to check'
    },
    examples: [
      'ISERR(A1/0)',   // TRUE (#DIV/0!)
      'ISERR(NA())'    // FALSE (it's #N/A)
    ],
    tips: [
      'Use when #N/A should be handled differently',
      'ISERROR catches all including #N/A'
    ],
    relatedFunctions: ['ISERROR', 'ISNA']
  },
  
  ISNA: {
    syntax: 'ISNA(value)',
    description: 'Returns TRUE if value is #N/A error',
    params: {
      value: 'Value to check'
    },
    examples: [
      'ISNA(VLOOKUP(A1, Data, 2, FALSE))',
      'ISNA(MATCH(A1, B:B, 0))'
    ],
    tips: [
      '#N/A usually means "not found" in lookups',
      'Use IFNA for cleaner handling'
    ],
    relatedFunctions: ['IFNA', 'ISERROR', 'NA']
  },
  
  ISNUMBER: {
    syntax: 'ISNUMBER(value)',
    description: 'Returns TRUE if value is a number',
    params: {
      value: 'Value to check'
    },
    examples: [
      'ISNUMBER(A1)',
      'ISNUMBER(SEARCH("text", A1))',  // Check if text found
      'IF(ISNUMBER(A1), A1*2, 0)'
    ],
    tips: [
      'Dates are numbers (returns TRUE)',
      'Text that looks like number returns FALSE'
    ],
    relatedFunctions: ['ISTEXT', 'VALUE', 'N']
  },
  
  ISTEXT: {
    syntax: 'ISTEXT(value)',
    description: 'Returns TRUE if value is text',
    params: {
      value: 'Value to check'
    },
    examples: [
      'ISTEXT(A1)',
      'IF(ISTEXT(A1), "Text", "Not text")'
    ],
    tips: [
      'Empty cell returns FALSE',
      'Number stored as text returns TRUE'
    ],
    relatedFunctions: ['ISNUMBER', 'T']
  },
  
  ISLOGICAL: {
    syntax: 'ISLOGICAL(value)',
    description: 'Returns TRUE if value is TRUE or FALSE',
    params: {
      value: 'Value to check'
    },
    examples: [
      'ISLOGICAL(TRUE)',  // TRUE
      'ISLOGICAL(1)',     // FALSE (1 is not boolean)
      'ISLOGICAL(A1>0)'   // TRUE (comparison result)
    ],
    relatedFunctions: ['ISNUMBER', 'ISTEXT']
  },
  
  ISNONTEXT: {
    syntax: 'ISNONTEXT(value)',
    description: 'Returns TRUE if value is NOT text',
    params: {
      value: 'Value to check'
    },
    examples: [
      'ISNONTEXT(123)',     // TRUE
      'ISNONTEXT("hello")' // FALSE
    ],
    tips: [
      'Empty cells return TRUE',
      'Opposite of ISTEXT (mostly)'
    ],
    relatedFunctions: ['ISTEXT', 'ISNUMBER']
  },
  
  ISREF: {
    syntax: 'ISREF(value)',
    description: 'Returns TRUE if value is a reference',
    params: {
      value: 'Value to check'
    },
    examples: [
      'ISREF(A1)',          // TRUE
      'ISREF(INDIRECT("A1"))' // TRUE
    ],
    relatedFunctions: ['INDIRECT', 'ADDRESS']
  },
  
  ISFORMULA: {
    syntax: 'ISFORMULA(cell)',
    description: 'Returns TRUE if cell contains a formula',
    params: {
      cell: 'Cell to check'
    },
    examples: [
      'ISFORMULA(A1)',
      'IF(ISFORMULA(A1), "Calculated", "Static")'
    ],
    tips: [
      'Useful for auditing spreadsheets',
      'Cannot check ranges'
    ],
    relatedFunctions: ['FORMULATEXT']
  },
  
  ISEVEN: {
    syntax: 'ISEVEN(value)',
    description: 'Returns TRUE if value is even',
    params: {
      value: 'Number to check'
    },
    examples: [
      'ISEVEN(4)',   // TRUE
      'ISEVEN(5)',   // FALSE
      'IF(ISEVEN(ROW()), "Even row", "Odd row")'
    ],
    relatedFunctions: ['ISODD', 'MOD']
  },
  
  ISODD: {
    syntax: 'ISODD(value)',
    description: 'Returns TRUE if value is odd',
    params: {
      value: 'Number to check'
    },
    examples: [
      'ISODD(5)',   // TRUE
      'ISODD(4)',   // FALSE
      'IF(ISODD(ROW()), "Odd row", "Even row")'
    ],
    relatedFunctions: ['ISEVEN', 'MOD']
  },
  
  // ========== SPECIAL VALUES ==========
  
  TRUE: {
    syntax: 'TRUE()',
    description: 'Returns the logical value TRUE',
    params: {},
    examples: [
      'TRUE()',
      'IF(A1=TRUE(), "Yes", "No")'
    ],
    tips: [
      'Usually just type TRUE without ()',
      'Function form useful in some contexts'
    ],
    relatedFunctions: ['FALSE', 'AND', 'OR']
  },
  
  FALSE: {
    syntax: 'FALSE()',
    description: 'Returns the logical value FALSE',
    params: {},
    examples: [
      'FALSE()',
      'IF(A1=FALSE(), "No", "Yes")'
    ],
    relatedFunctions: ['TRUE', 'AND', 'OR']
  },
  
  NA: {
    syntax: 'NA()',
    description: 'Returns the error value #N/A',
    params: {},
    examples: [
      'IF(A1="", NA(), A1)',  // Mark blank as N/A
      'IFERROR(A1, NA())'
    ],
    tips: [
      '#N/A means "not available"',
      'Use to indicate missing data',
      'Charts skip #N/A values'
    ],
    relatedFunctions: ['ISNA', 'IFNA']
  },
  
  // ========== INFORMATION ==========
  
  TYPE: {
    syntax: 'TYPE(value)',
    description: 'Returns a number indicating the data type',
    params: {
      value: 'Value to check'
    },
    examples: [
      'TYPE(A1)',
      'SWITCH(TYPE(A1), 1, "Number", 2, "Text", 4, "Boolean", 16, "Error", 64, "Array")'
    ],
    tips: [
      '1=Number, 2=Text, 4=Boolean, 16=Error, 64=Array, 128=Compound'
    ],
    relatedFunctions: ['ISNUMBER', 'ISTEXT', 'ISERROR']
  },
  
  N: {
    syntax: 'N(value)',
    description: 'Converts value to a number',
    params: {
      value: 'Value to convert'
    },
    examples: [
      'N(TRUE)',   // Returns 1
      'N(FALSE)',  // Returns 0
      'N("text")' // Returns 0
    ],
    tips: [
      'TRUE = 1, FALSE = 0',
      'Text = 0, Dates = serial number',
      'Error = error'
    ],
    relatedFunctions: ['T', 'VALUE']
  },
  
  ERROR_TYPE: {
    syntax: 'ERROR.TYPE(reference)',
    description: 'Returns number indicating the type of error',
    params: {
      reference: 'Cell containing error'
    },
    examples: [
      'ERROR.TYPE(A1)',
      'SWITCH(ERROR.TYPE(A1), 1, "NULL", 2, "DIV/0", 3, "VALUE", 4, "REF", 5, "NAME", 6, "NUM", 7, "N/A", 8, "GETTING_DATA")'
    ],
    tips: [
      '1=#NULL!, 2=#DIV/0!, 3=#VALUE!, 4=#REF!, 5=#NAME?, 6=#NUM!, 7=#N/A, 8=#GETTING_DATA',
      'Returns #N/A if no error'
    ],
    relatedFunctions: ['ISERROR', 'IFERROR']
  },
  
  CELL: {
    syntax: 'CELL(info_type, [reference])',
    description: 'Returns information about a cell',
    params: {
      info_type: 'Type of information to return',
      reference: 'Optional. Cell to get info about'
    },
    examples: [
      'CELL("address", A1)',   // Returns "$A$1"
      'CELL("col", A1)',       // Returns 1
      'CELL("row", A1)',       // Returns 1
      'CELL("contents", A1)',  // Returns cell value
      'CELL("type", A1)',      // "l"=blank, "v"=number, "b"=label
      'CELL("width", A1)',     // Column width
      'CELL("format", A1)'     // Number format code
    ],
    tips: [
      'info_type options: address, col, row, contents, type, width, format, prefix, protect, filename, color, parentheses'
    ],
    relatedFunctions: ['ROW', 'COLUMN', 'ADDRESS']
  },
  
  SHEET: {
    syntax: 'SHEET([value])',
    description: 'Returns the sheet number',
    params: {
      value: 'Optional. Sheet name or cell reference'
    },
    examples: [
      'SHEET()',         // Current sheet number
      'SHEET("Sheet2")', // Sheet2's number
      'SHEET(A1)'        // Sheet containing A1
    ],
    relatedFunctions: ['SHEETS']
  },
  
  SHEETS: {
    syntax: 'SHEETS([reference])',
    description: 'Returns the number of sheets in reference',
    params: {
      reference: 'Optional. Reference (default: all sheets)'
    },
    examples: [
      'SHEETS()'  // Total sheets in workbook
    ],
    relatedFunctions: ['SHEET']
  },
  
  // ========== DELTA COMPARISON ==========
  
  DELTA: {
    syntax: 'DELTA(number1, [number2])',
    description: 'Returns 1 if numbers are equal, 0 otherwise',
    params: {
      number1: 'First number',
      number2: 'Optional. Second number (default: 0)'
    },
    examples: [
      'DELTA(5, 5)',   // Returns 1
      'DELTA(5, 6)',   // Returns 0
      'DELTA(0)'       // Returns 1 (comparing to 0)
    ],
    tips: [
      'Useful for engineering calculations',
      'Can use for exact numeric matching'
    ],
    relatedFunctions: ['EXACT', 'GESTEP']
  },
  
  GESTEP: {
    syntax: 'GESTEP(value, [step])',
    description: 'Returns 1 if value >= step, 0 otherwise',
    params: {
      value: 'Value to test',
      step: 'Optional. Threshold (default: 0)'
    },
    examples: [
      'GESTEP(5, 3)',   // Returns 1 (5 >= 3)
      'GESTEP(2, 3)',   // Returns 0 (2 < 3)
      'GESTEP(-1)'      // Returns 0 (-1 < 0)
    ],
    tips: [
      'Engineering function',
      'Like IF(value>=step, 1, 0)'
    ],
    relatedFunctions: ['DELTA', 'IF']
  }
};

/**
 * Get logic function info
 */
function getLogicFunction(name) {
  return LOGIC_FUNCTIONS[name.toUpperCase()] || null;
}

/**
 * Get all logic function names
 */
function getLogicFunctionNames() {
  return Object.keys(LOGIC_FUNCTIONS);
}

Logger.log('📦 FormulaCatalog_Logic loaded - ' + Object.keys(LOGIC_FUNCTIONS).length + ' functions');
