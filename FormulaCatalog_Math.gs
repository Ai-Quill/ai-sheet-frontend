/**
 * @file FormulaCatalog_Math.gs
 * @description Math, Statistics & Aggregation Functions for Google Sheets
 * 
 * COMPLETE COVERAGE of mathematical and statistical functions.
 */

var MATH_FUNCTIONS = {
  
  // ========== BASIC AGGREGATION ==========
  
  SUM: {
    syntax: 'SUM(value1, [value2], ...)',
    description: 'Adds all the numbers in a range',
    params: {
      value1: 'First value or range',
      value2: 'Additional values or ranges'
    },
    examples: [
      'SUM(A1:A10)',
      'SUM(A1:A10, C1:C10)',
      'SUM(1, 2, 3, 4, 5)'
    ],
    tips: [
      'Ignores text and empty cells',
      'Can sum multiple ranges'
    ],
    relatedFunctions: ['SUMIF', 'SUMIFS', 'SUMPRODUCT']
  },
  
  SUMIF: {
    syntax: 'SUMIF(range, criterion, [sum_range])',
    description: 'Sums cells that meet a single condition',
    params: {
      range: 'Range to check against criterion',
      criterion: 'Condition (number, text, or expression)',
      sum_range: 'Optional. Range to sum (defaults to range)'
    },
    examples: [
      'SUMIF(A:A, "Apple", B:B)',
      'SUMIF(B:B, ">100")',
      'SUMIF(A:A, "*sales*", C:C)'  // Wildcard
    ],
    tips: [
      'Criterion can use wildcards (* and ?)',
      'Use "" for empty cells criterion',
      'For multiple conditions, use SUMIFS'
    ],
    relatedFunctions: ['SUMIFS', 'COUNTIF', 'AVERAGEIF']
  },
  
  SUMIFS: {
    syntax: 'SUMIFS(sum_range, criteria_range1, criterion1, [criteria_range2, criterion2], ...)',
    description: 'Sums cells that meet multiple conditions',
    params: {
      sum_range: 'Range to sum',
      criteria_range1: 'First range to check',
      criterion1: 'First condition',
      criteria_range2: 'Second range to check',
      criterion2: 'Second condition'
    },
    examples: [
      'SUMIFS(C:C, A:A, "Apple", B:B, ">100")',
      'SUMIFS(Sales, Region, "West", Year, 2024)',
      'SUMIFS(Amount, Date, ">="&DATE(2024,1,1), Date, "<"&DATE(2025,1,1))'
    ],
    tips: [
      'All ranges must be same size',
      'Use & to build criteria with cell references'
    ],
    relatedFunctions: ['SUMIF', 'COUNTIFS', 'AVERAGEIFS']
  },
  
  SUMPRODUCT: {
    syntax: 'SUMPRODUCT(array1, [array2], ...)',
    description: 'Multiplies corresponding elements and sums the products',
    params: {
      array1: 'First array',
      array2: 'Additional arrays (same dimensions)'
    },
    examples: [
      'SUMPRODUCT(A1:A10, B1:B10)',  // Sum of A*B for each row
      'SUMPRODUCT((A:A="Apple")*(B:B))',  // Conditional sum (legacy)
      'SUMPRODUCT(Qty, Price)'
    ],
    tips: [
      'Powerful for weighted calculations',
      'Can emulate SUMIFS with boolean multiplication',
      'Arrays must have same dimensions'
    ],
    relatedFunctions: ['SUM', 'SUMIFS', 'MMULT']
  },
  
  // ========== AVERAGE ==========
  
  AVERAGE: {
    syntax: 'AVERAGE(value1, [value2], ...)',
    description: 'Returns the arithmetic mean of values',
    params: {
      value1: 'First value or range',
      value2: 'Additional values or ranges'
    },
    examples: [
      'AVERAGE(A1:A10)',
      'AVERAGE(A1:A10, C1:C10)'
    ],
    tips: [
      'Ignores text and empty cells',
      'Use AVERAGEA to include text as 0'
    ],
    relatedFunctions: ['AVERAGEIF', 'AVERAGEIFS', 'MEDIAN', 'MODE']
  },
  
  AVERAGEIF: {
    syntax: 'AVERAGEIF(criteria_range, criterion, [average_range])',
    description: 'Averages cells that meet a condition',
    params: {
      criteria_range: 'Range to check',
      criterion: 'Condition',
      average_range: 'Optional. Range to average'
    },
    examples: [
      'AVERAGEIF(A:A, "Apple", B:B)',
      'AVERAGEIF(B:B, ">0")',
      'AVERAGEIF(Status, "Complete", Score)'
    ],
    relatedFunctions: ['AVERAGE', 'AVERAGEIFS', 'SUMIF']
  },
  
  AVERAGEIFS: {
    syntax: 'AVERAGEIFS(average_range, criteria_range1, criterion1, ...)',
    description: 'Averages cells that meet multiple conditions',
    params: {
      average_range: 'Range to average',
      criteria_range1: 'First range to check',
      criterion1: 'First condition'
    },
    examples: [
      'AVERAGEIFS(Score, Region, "West", Year, 2024)',
      'AVERAGEIFS(C:C, A:A, "Apple", B:B, ">100")'
    ],
    relatedFunctions: ['AVERAGEIF', 'AVERAGE', 'SUMIFS']
  },
  
  AVERAGEA: {
    syntax: 'AVERAGEA(value1, [value2], ...)',
    description: 'Average including text (as 0) and FALSE',
    params: {
      value1: 'First value or range'
    },
    examples: [
      'AVERAGEA(A1:A10)'
    ],
    tips: [
      'Text = 0, TRUE = 1, FALSE = 0',
      'Empty cells are ignored'
    ],
    relatedFunctions: ['AVERAGE', 'COUNTA']
  },
  
  // ========== COUNT ==========
  
  COUNT: {
    syntax: 'COUNT(value1, [value2], ...)',
    description: 'Counts cells containing numbers',
    params: {
      value1: 'First value or range'
    },
    examples: [
      'COUNT(A1:A10)',
      'COUNT(A:A, C:C)'
    ],
    tips: [
      'Only counts numbers',
      'Use COUNTA for all non-empty cells'
    ],
    relatedFunctions: ['COUNTA', 'COUNTIF', 'COUNTBLANK']
  },
  
  COUNTA: {
    syntax: 'COUNTA(value1, [value2], ...)',
    description: 'Counts non-empty cells',
    params: {
      value1: 'First value or range'
    },
    examples: [
      'COUNTA(A:A)',
      'COUNTA(A1:A10) - 1'  // Subtract header
    ],
    tips: [
      'Counts text, numbers, errors, TRUE/FALSE',
      'Does not count empty cells'
    ],
    relatedFunctions: ['COUNT', 'COUNTBLANK', 'COUNTIF']
  },
  
  COUNTIF: {
    syntax: 'COUNTIF(range, criterion)',
    description: 'Counts cells that meet a condition',
    params: {
      range: 'Range to check',
      criterion: 'Condition'
    },
    examples: [
      'COUNTIF(A:A, "Apple")',
      'COUNTIF(B:B, ">100")',
      'COUNTIF(A:A, "*@gmail.com")',  // Wildcard
      'COUNTIF(A:A, "")'  // Count empty
    ],
    tips: [
      'Use wildcards * and ?',
      'Use <> for "not equal"',
      'Use "" for empty cells'
    ],
    relatedFunctions: ['COUNTIFS', 'SUMIF', 'COUNT']
  },
  
  COUNTIFS: {
    syntax: 'COUNTIFS(criteria_range1, criterion1, [criteria_range2, criterion2], ...)',
    description: 'Counts cells meeting multiple conditions',
    params: {
      criteria_range1: 'First range',
      criterion1: 'First condition'
    },
    examples: [
      'COUNTIFS(A:A, "Apple", B:B, ">100")',
      'COUNTIFS(Status, "Complete", Region, "West")'
    ],
    relatedFunctions: ['COUNTIF', 'SUMIFS', 'AVERAGEIFS']
  },
  
  COUNTBLANK: {
    syntax: 'COUNTBLANK(range)',
    description: 'Counts empty cells',
    params: {
      range: 'Range to check'
    },
    examples: [
      'COUNTBLANK(A1:A100)',
      'ROWS(A:A) - COUNTA(A:A)'  // Alternative
    ],
    tips: [
      'Cells with "" formula count as blank',
      'Cells with space do NOT count as blank'
    ],
    relatedFunctions: ['COUNTA', 'ISBLANK']
  },
  
  COUNTUNIQUE: {
    syntax: 'COUNTUNIQUE(value1, [value2], ...)',
    description: 'Counts unique values (Google Sheets specific)',
    params: {
      value1: 'Values or ranges to count unique values in'
    },
    examples: [
      'COUNTUNIQUE(A:A)',
      'COUNTUNIQUE(A1:A100, B1:B100)'
    ],
    tips: [
      'Case-sensitive',
      'Google Sheets specific function'
    ],
    relatedFunctions: ['UNIQUE', 'COUNTA', 'COUNTIF']
  },
  
  // ========== MIN/MAX ==========
  
  MIN: {
    syntax: 'MIN(value1, [value2], ...)',
    description: 'Returns the minimum value',
    params: {
      value1: 'First value or range'
    },
    examples: [
      'MIN(A1:A10)',
      'MIN(A:A, C:C)'
    ],
    relatedFunctions: ['MAX', 'MINIFS', 'SMALL']
  },
  
  MAX: {
    syntax: 'MAX(value1, [value2], ...)',
    description: 'Returns the maximum value',
    params: {
      value1: 'First value or range'
    },
    examples: [
      'MAX(A1:A10)',
      'MAX(A:A, C:C)'
    ],
    relatedFunctions: ['MIN', 'MAXIFS', 'LARGE']
  },
  
  MINIFS: {
    syntax: 'MINIFS(range, criteria_range1, criterion1, ...)',
    description: 'Returns minimum value meeting conditions',
    params: {
      range: 'Range to find minimum in',
      criteria_range1: 'Range to check',
      criterion1: 'Condition'
    },
    examples: [
      'MINIFS(B:B, A:A, "Apple")',
      'MINIFS(Price, Category, "Electronics", Stock, ">0")'
    ],
    relatedFunctions: ['MIN', 'MAXIFS', 'AVERAGEIFS']
  },
  
  MAXIFS: {
    syntax: 'MAXIFS(range, criteria_range1, criterion1, ...)',
    description: 'Returns maximum value meeting conditions',
    params: {
      range: 'Range to find maximum in',
      criteria_range1: 'Range to check',
      criterion1: 'Condition'
    },
    examples: [
      'MAXIFS(B:B, A:A, "Apple")',
      'MAXIFS(Score, Status, "Active")'
    ],
    relatedFunctions: ['MAX', 'MINIFS', 'AVERAGEIFS']
  },
  
  SMALL: {
    syntax: 'SMALL(data, n)',
    description: 'Returns the n-th smallest value',
    params: {
      data: 'Array or range',
      n: 'Position (1 = smallest)'
    },
    examples: [
      'SMALL(A1:A10, 1)',  // Same as MIN
      'SMALL(A1:A10, 3)'   // 3rd smallest
    ],
    relatedFunctions: ['LARGE', 'MIN', 'PERCENTILE']
  },
  
  LARGE: {
    syntax: 'LARGE(data, n)',
    description: 'Returns the n-th largest value',
    params: {
      data: 'Array or range',
      n: 'Position (1 = largest)'
    },
    examples: [
      'LARGE(A1:A10, 1)',  // Same as MAX
      'LARGE(A1:A10, 3)'   // 3rd largest
    ],
    relatedFunctions: ['SMALL', 'MAX', 'PERCENTILE']
  },
  
  // ========== ROUNDING ==========
  
  ROUND: {
    syntax: 'ROUND(value, [places])',
    description: 'Rounds to specified decimal places',
    params: {
      value: 'Number to round',
      places: 'Optional. Decimal places (default: 0)'
    },
    examples: [
      'ROUND(3.14159, 2)',  // Returns 3.14
      'ROUND(1234, -2)'     // Returns 1200
    ],
    tips: [
      'Negative places round to left of decimal',
      'Rounds 5 up (banker\'s rounding may differ)'
    ],
    relatedFunctions: ['ROUNDUP', 'ROUNDDOWN', 'MROUND', 'TRUNC']
  },
  
  ROUNDUP: {
    syntax: 'ROUNDUP(value, [places])',
    description: 'Rounds up (away from zero)',
    params: {
      value: 'Number to round',
      places: 'Optional. Decimal places'
    },
    examples: [
      'ROUNDUP(3.14, 1)',   // Returns 3.2
      'ROUNDUP(-3.14, 1)'   // Returns -3.2 (away from zero)
    ],
    relatedFunctions: ['ROUNDDOWN', 'CEILING', 'ROUND']
  },
  
  ROUNDDOWN: {
    syntax: 'ROUNDDOWN(value, [places])',
    description: 'Rounds down (toward zero)',
    params: {
      value: 'Number to round',
      places: 'Optional. Decimal places'
    },
    examples: [
      'ROUNDDOWN(3.99, 0)',  // Returns 3
      'ROUNDDOWN(-3.99, 0)'  // Returns -3 (toward zero)
    ],
    relatedFunctions: ['ROUNDUP', 'FLOOR', 'TRUNC']
  },
  
  CEILING: {
    syntax: 'CEILING(value, [factor])',
    description: 'Rounds up to nearest multiple of factor',
    params: {
      value: 'Number to round',
      factor: 'Optional. Multiple to round to (default: 1)'
    },
    examples: [
      'CEILING(2.3)',       // Returns 3
      'CEILING(2.3, 0.5)',  // Returns 2.5
      'CEILING(23, 10)'     // Returns 30
    ],
    relatedFunctions: ['FLOOR', 'MROUND', 'CEILING.MATH']
  },
  
  FLOOR: {
    syntax: 'FLOOR(value, [factor])',
    description: 'Rounds down to nearest multiple of factor',
    params: {
      value: 'Number to round',
      factor: 'Optional. Multiple to round to'
    },
    examples: [
      'FLOOR(2.7)',        // Returns 2
      'FLOOR(2.7, 0.5)',   // Returns 2.5
      'FLOOR(27, 10)'      // Returns 20
    ],
    relatedFunctions: ['CEILING', 'MROUND', 'FLOOR.MATH']
  },
  
  MROUND: {
    syntax: 'MROUND(value, factor)',
    description: 'Rounds to nearest multiple of factor',
    params: {
      value: 'Number to round',
      factor: 'Multiple to round to'
    },
    examples: [
      'MROUND(10, 3)',    // Returns 9
      'MROUND(1.25, 0.5)' // Returns 1.5
    ],
    relatedFunctions: ['ROUND', 'CEILING', 'FLOOR']
  },
  
  TRUNC: {
    syntax: 'TRUNC(value, [places])',
    description: 'Truncates to specified decimal places (no rounding)',
    params: {
      value: 'Number to truncate',
      places: 'Optional. Decimal places to keep'
    },
    examples: [
      'TRUNC(3.99)',     // Returns 3
      'TRUNC(3.99, 1)'   // Returns 3.9
    ],
    tips: [
      'Does NOT round - just cuts off',
      'Same as INT for positive numbers'
    ],
    relatedFunctions: ['INT', 'ROUND', 'ROUNDDOWN']
  },
  
  INT: {
    syntax: 'INT(value)',
    description: 'Rounds down to nearest integer',
    params: {
      value: 'Number to convert'
    },
    examples: [
      'INT(3.9)',   // Returns 3
      'INT(-3.9)'   // Returns -4 (rounds toward negative infinity)
    ],
    tips: [
      'Always rounds toward negative infinity',
      'Different from TRUNC for negative numbers'
    ],
    relatedFunctions: ['TRUNC', 'ROUND', 'FLOOR']
  },
  
  // ========== STATISTICAL ==========
  
  MEDIAN: {
    syntax: 'MEDIAN(value1, [value2], ...)',
    description: 'Returns the median (middle value)',
    params: {
      value1: 'First value or range'
    },
    examples: [
      'MEDIAN(A1:A10)',
      'MEDIAN(1, 3, 5, 7, 9)'  // Returns 5
    ],
    tips: [
      'For even count, averages middle two values',
      'More resistant to outliers than AVERAGE'
    ],
    relatedFunctions: ['AVERAGE', 'MODE', 'PERCENTILE']
  },
  
  MODE: {
    syntax: 'MODE(value1, [value2], ...)',
    description: 'Returns the most frequent value',
    params: {
      value1: 'First value or range'
    },
    examples: [
      'MODE(A1:A10)',
      'MODE(1, 2, 2, 3, 3, 3)'  // Returns 3
    ],
    tips: [
      'Returns #N/A if no duplicates',
      'Use MODE.MULT for multiple modes'
    ],
    relatedFunctions: ['MEDIAN', 'AVERAGE', 'MODE.SNGL', 'MODE.MULT']
  },
  
  STDEV: {
    syntax: 'STDEV(value1, [value2], ...)',
    description: 'Standard deviation of a sample',
    params: {
      value1: 'Values or range'
    },
    examples: [
      'STDEV(A1:A100)'
    ],
    tips: [
      'For sample data (n-1 denominator)',
      'Use STDEVP for entire population'
    ],
    relatedFunctions: ['STDEVP', 'VAR', 'AVERAGE']
  },
  
  STDEVP: {
    syntax: 'STDEVP(value1, [value2], ...)',
    description: 'Standard deviation of entire population',
    params: {
      value1: 'Values or range'
    },
    examples: [
      'STDEVP(A1:A100)'
    ],
    relatedFunctions: ['STDEV', 'VARP']
  },
  
  VAR: {
    syntax: 'VAR(value1, [value2], ...)',
    description: 'Variance of a sample',
    params: {
      value1: 'Values or range'
    },
    examples: [
      'VAR(A1:A100)'
    ],
    relatedFunctions: ['VARP', 'STDEV']
  },
  
  VARP: {
    syntax: 'VARP(value1, [value2], ...)',
    description: 'Variance of entire population',
    params: {
      value1: 'Values or range'
    },
    examples: [
      'VARP(A1:A100)'
    ],
    relatedFunctions: ['VAR', 'STDEVP']
  },
  
  PERCENTILE: {
    syntax: 'PERCENTILE(data, percentile)',
    description: 'Returns value at given percentile',
    params: {
      data: 'Array or range',
      percentile: 'Percentile (0 to 1)'
    },
    examples: [
      'PERCENTILE(A1:A100, 0.5)',   // Median
      'PERCENTILE(A1:A100, 0.9)',   // 90th percentile
      'PERCENTILE(Scores, 0.25)'    // First quartile
    ],
    relatedFunctions: ['PERCENTRANK', 'QUARTILE', 'MEDIAN']
  },
  
  PERCENTRANK: {
    syntax: 'PERCENTRANK(data, value, [significant_digits])',
    description: 'Returns percentile rank of a value',
    params: {
      data: 'Array or range',
      value: 'Value to find rank of',
      significant_digits: 'Optional. Precision'
    },
    examples: [
      'PERCENTRANK(A1:A100, 75)'
    ],
    relatedFunctions: ['PERCENTILE', 'RANK']
  },
  
  QUARTILE: {
    syntax: 'QUARTILE(data, quart)',
    description: 'Returns quartile of data set',
    params: {
      data: 'Array or range',
      quart: '0=min, 1=Q1, 2=median, 3=Q3, 4=max'
    },
    examples: [
      'QUARTILE(A:A, 1)',  // First quartile (25%)
      'QUARTILE(A:A, 3)'   // Third quartile (75%)
    ],
    relatedFunctions: ['PERCENTILE', 'MEDIAN']
  },
  
  RANK: {
    syntax: 'RANK(value, data, [is_ascending])',
    description: 'Returns rank of a value in a data set',
    params: {
      value: 'Value to rank',
      data: 'Array or range',
      is_ascending: 'Optional. TRUE for ascending rank'
    },
    examples: [
      'RANK(A1, A:A)',       // Descending rank (highest = 1)
      'RANK(A1, A:A, TRUE)'  // Ascending rank (lowest = 1)
    ],
    relatedFunctions: ['PERCENTRANK', 'LARGE', 'SMALL']
  },
  
  // ========== BASIC MATH ==========
  
  ABS: {
    syntax: 'ABS(value)',
    description: 'Returns absolute value',
    params: {
      value: 'Number'
    },
    examples: [
      'ABS(-5)',  // Returns 5
      'ABS(A1-B1)'
    ],
    relatedFunctions: ['SIGN']
  },
  
  SIGN: {
    syntax: 'SIGN(value)',
    description: 'Returns sign of number (-1, 0, or 1)',
    params: {
      value: 'Number'
    },
    examples: [
      'SIGN(-5)',  // Returns -1
      'SIGN(0)',   // Returns 0
      'SIGN(5)'    // Returns 1
    ],
    relatedFunctions: ['ABS']
  },
  
  MOD: {
    syntax: 'MOD(dividend, divisor)',
    description: 'Returns remainder after division',
    params: {
      dividend: 'Number to divide',
      divisor: 'Number to divide by'
    },
    examples: [
      'MOD(10, 3)',     // Returns 1
      'MOD(ROW(), 2)',  // Alternating 0 and 1
      'MOD(A1, 1)=0'    // Check if whole number
    ],
    tips: [
      'Result has same sign as divisor',
      'Useful for alternating patterns'
    ],
    relatedFunctions: ['QUOTIENT', 'INT']
  },
  
  QUOTIENT: {
    syntax: 'QUOTIENT(dividend, divisor)',
    description: 'Returns integer portion of division',
    params: {
      dividend: 'Number to divide',
      divisor: 'Number to divide by'
    },
    examples: [
      'QUOTIENT(10, 3)',  // Returns 3
      'QUOTIENT(A1, B1)'
    ],
    relatedFunctions: ['MOD', 'INT']
  },
  
  POWER: {
    syntax: 'POWER(base, exponent)',
    description: 'Returns number raised to a power',
    params: {
      base: 'Base number',
      exponent: 'Power to raise to'
    },
    examples: [
      'POWER(2, 8)',    // Returns 256
      'POWER(A1, 0.5)'  // Square root
    ],
    tips: [
      'Can also use ^ operator: 2^8',
      'POWER(x, 0.5) = SQRT(x)'
    ],
    relatedFunctions: ['SQRT', 'EXP', 'LOG']
  },
  
  SQRT: {
    syntax: 'SQRT(value)',
    description: 'Returns square root',
    params: {
      value: 'Number (must be >= 0)'
    },
    examples: [
      'SQRT(16)',  // Returns 4
      'SQRT(2)'    // Returns ~1.414
    ],
    relatedFunctions: ['POWER', 'SQRTPI']
  },
  
  SQRTPI: {
    syntax: 'SQRTPI(value)',
    description: 'Returns square root of (value * PI)',
    params: {
      value: 'Number'
    },
    examples: [
      'SQRTPI(1)'  // Returns ~1.772 (sqrt of pi)
    ],
    relatedFunctions: ['SQRT', 'PI']
  },
  
  EXP: {
    syntax: 'EXP(exponent)',
    description: 'Returns e raised to a power',
    params: {
      exponent: 'Power'
    },
    examples: [
      'EXP(1)',   // Returns e (~2.718)
      'EXP(2)'    // Returns e^2 (~7.389)
    ],
    relatedFunctions: ['LN', 'LOG', 'POWER']
  },
  
  LN: {
    syntax: 'LN(value)',
    description: 'Returns natural logarithm (base e)',
    params: {
      value: 'Number (must be > 0)'
    },
    examples: [
      'LN(EXP(1))',  // Returns 1
      'LN(100)'
    ],
    relatedFunctions: ['LOG', 'LOG10', 'EXP']
  },
  
  LOG: {
    syntax: 'LOG(value, [base])',
    description: 'Returns logarithm with specified base',
    params: {
      value: 'Number (must be > 0)',
      base: 'Optional. Base (default: 10)'
    },
    examples: [
      'LOG(100)',      // Returns 2 (log base 10)
      'LOG(8, 2)',     // Returns 3 (log base 2)
      'LOG(A1, 10)'
    ],
    relatedFunctions: ['LN', 'LOG10']
  },
  
  LOG10: {
    syntax: 'LOG10(value)',
    description: 'Returns base-10 logarithm',
    params: {
      value: 'Number'
    },
    examples: [
      'LOG10(100)',  // Returns 2
      'LOG10(1000)'  // Returns 3
    ],
    relatedFunctions: ['LOG', 'LN']
  },
  
  // ========== PRODUCT & FACTORIALS ==========
  
  PRODUCT: {
    syntax: 'PRODUCT(factor1, [factor2], ...)',
    description: 'Multiplies all values',
    params: {
      factor1: 'First value or range'
    },
    examples: [
      'PRODUCT(A1:A10)',
      'PRODUCT(1, 2, 3, 4, 5)'  // Returns 120
    ],
    relatedFunctions: ['SUM', 'SUMPRODUCT']
  },
  
  FACT: {
    syntax: 'FACT(value)',
    description: 'Returns factorial (n!)',
    params: {
      value: 'Non-negative integer'
    },
    examples: [
      'FACT(5)',  // Returns 120 (5! = 5*4*3*2*1)
      'FACT(0)'   // Returns 1
    ],
    relatedFunctions: ['FACTDOUBLE', 'COMBIN', 'PERMUT']
  },
  
  FACTDOUBLE: {
    syntax: 'FACTDOUBLE(value)',
    description: 'Returns double factorial',
    params: {
      value: 'Non-negative integer'
    },
    examples: [
      'FACTDOUBLE(6)',  // Returns 48 (6*4*2)
      'FACTDOUBLE(7)'   // Returns 105 (7*5*3*1)
    ],
    relatedFunctions: ['FACT']
  },
  
  // ========== COMBINATIONS & PERMUTATIONS ==========
  
  COMBIN: {
    syntax: 'COMBIN(n, k)',
    description: 'Returns number of combinations (n choose k)',
    params: {
      n: 'Total items',
      k: 'Items to choose'
    },
    examples: [
      'COMBIN(10, 3)',  // Ways to choose 3 from 10
      'COMBIN(52, 5)'   // Poker hands
    ],
    tips: [
      'Order does not matter',
      'Formula: n! / (k! * (n-k)!)'
    ],
    relatedFunctions: ['COMBINA', 'PERMUT', 'FACT']
  },
  
  COMBINA: {
    syntax: 'COMBINA(n, k)',
    description: 'Combinations with repetition allowed',
    params: {
      n: 'Total items',
      k: 'Items to choose'
    },
    examples: [
      'COMBINA(4, 3)'
    ],
    relatedFunctions: ['COMBIN', 'PERMUT']
  },
  
  PERMUT: {
    syntax: 'PERMUT(n, k)',
    description: 'Returns number of permutations',
    params: {
      n: 'Total items',
      k: 'Items to arrange'
    },
    examples: [
      'PERMUT(10, 3)'  // Ways to arrange 3 from 10
    ],
    tips: [
      'Order matters',
      'Formula: n! / (n-k)!'
    ],
    relatedFunctions: ['COMBIN', 'FACT']
  },
  
  // ========== RANDOM ==========
  
  RAND: {
    syntax: 'RAND()',
    description: 'Returns random number between 0 and 1',
    params: {},
    examples: [
      'RAND()',
      'RAND() * 100',  // Random 0-100
      'INT(RAND() * 10) + 1'  // Random integer 1-10
    ],
    tips: [
      'Recalculates on every change',
      'Use RANDBETWEEN for integers'
    ],
    relatedFunctions: ['RANDBETWEEN', 'RANDARRAY']
  },
  
  RANDBETWEEN: {
    syntax: 'RANDBETWEEN(low, high)',
    description: 'Returns random integer between low and high',
    params: {
      low: 'Minimum value (inclusive)',
      high: 'Maximum value (inclusive)'
    },
    examples: [
      'RANDBETWEEN(1, 100)',
      'RANDBETWEEN(1, 6)'  // Dice roll
    ],
    relatedFunctions: ['RAND', 'RANDARRAY']
  },
  
  RANDARRAY: {
    syntax: 'RANDARRAY([rows], [columns], [min], [max], [whole_number])',
    description: 'Returns array of random numbers',
    params: {
      rows: 'Optional. Number of rows (default: 1)',
      columns: 'Optional. Number of columns (default: 1)',
      min: 'Optional. Minimum value (default: 0)',
      max: 'Optional. Maximum value (default: 1)',
      whole_number: 'Optional. TRUE for integers'
    },
    examples: [
      'RANDARRAY(5, 3)',  // 5x3 array of random 0-1
      'RANDARRAY(10, 1, 1, 100, TRUE)'  // 10 random integers 1-100
    ],
    relatedFunctions: ['RAND', 'RANDBETWEEN', 'SEQUENCE']
  },
  
  // ========== SEQUENCES ==========
  
  SEQUENCE: {
    syntax: 'SEQUENCE(rows, [columns], [start], [step])',
    description: 'Generates a sequence of numbers',
    params: {
      rows: 'Number of rows',
      columns: 'Optional. Number of columns (default: 1)',
      start: 'Optional. Starting number (default: 1)',
      step: 'Optional. Increment (default: 1)'
    },
    examples: [
      'SEQUENCE(10)',        // 1 to 10
      'SEQUENCE(5, 1, 0, 2)',  // 0, 2, 4, 6, 8
      'SEQUENCE(3, 3)',      // 3x3 grid
      'SEQUENCE(12, 1, DATE(2024,1,1), 30)'  // Monthly dates
    ],
    tips: [
      'Great for creating row numbers',
      'Can generate dates with step'
    ],
    relatedFunctions: ['ROW', 'RANDARRAY']
  }
};

/**
 * Get math function info
 */
function getMathFunction(name) {
  return MATH_FUNCTIONS[name.toUpperCase()] || null;
}

/**
 * Get all math function names
 */
function getMathFunctionNames() {
  return Object.keys(MATH_FUNCTIONS);
}

Logger.log('📦 FormulaCatalog_Math loaded - ' + Object.keys(MATH_FUNCTIONS).length + ' functions');
