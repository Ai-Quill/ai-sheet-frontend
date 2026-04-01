/**
 * @file FormulaCatalog_Text.gs
 * @description Text & String Functions for Google Sheets
 * 
 * COMPLETE COVERAGE of all text manipulation functions.
 * These functions help you manipulate, search, and transform text.
 */

var TEXT_FUNCTIONS = {
  
  // ========== BASIC TEXT MANIPULATION ==========
  
  CONCATENATE: {
    syntax: 'CONCATENATE(string1, [string2], ...)',
    description: 'Joins multiple text strings into one',
    params: {
      string1: 'First string',
      string2: 'Additional strings (optional)'
    },
    examples: [
      'CONCATENATE(A1, " ", B1)',
      'CONCATENATE("Hello", " ", "World")'
    ],
    tips: [
      'Can also use & operator: A1 & " " & B1',
      'Consider TEXTJOIN for more control'
    ],
    relatedFunctions: ['TEXTJOIN', 'CONCAT', 'JOIN']
  },
  
  CONCAT: {
    syntax: 'CONCAT(value1, value2)',
    description: 'Joins exactly two values',
    params: {
      value1: 'First value',
      value2: 'Second value'
    },
    examples: [
      'CONCAT(A1, B1)',
      'CONCAT("ID-", ROW())'
    ],
    tips: [
      'Only takes 2 arguments',
      'Use CONCATENATE or TEXTJOIN for multiple values'
    ],
    relatedFunctions: ['CONCATENATE', 'TEXTJOIN']
  },
  
  TEXTJOIN: {
    syntax: 'TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)',
    description: 'Joins text with a delimiter, optionally ignoring empty cells',
    params: {
      delimiter: 'Text to insert between values',
      ignore_empty: 'TRUE to skip empty cells',
      text1: 'First text or range',
      text2: 'Additional text or ranges'
    },
    examples: [
      'TEXTJOIN(", ", TRUE, A1:A10)',
      'TEXTJOIN("-", FALSE, A1, B1, C1)',
      'TEXTJOIN(CHAR(10), TRUE, A:A)'  // Join with newlines
    ],
    tips: [
      'Can join entire ranges',
      'CHAR(10) for newline delimiter',
      'Much more powerful than CONCATENATE'
    ],
    relatedFunctions: ['CONCATENATE', 'JOIN', 'SPLIT']
  },
  
  JOIN: {
    syntax: 'JOIN(delimiter, value_or_array1, [value_or_array2], ...)',
    description: 'Joins values with a delimiter (Google Sheets specific)',
    params: {
      delimiter: 'Text to insert between values',
      value_or_array1: 'Values or arrays to join'
    },
    examples: [
      'JOIN(", ", A1:A5)',
      'JOIN(" | ", "A", "B", "C")'
    ],
    tips: [
      'Similar to TEXTJOIN but simpler syntax',
      'Does not have ignore_empty option'
    ],
    relatedFunctions: ['TEXTJOIN', 'CONCATENATE', 'SPLIT']
  },
  
  SPLIT: {
    syntax: 'SPLIT(text, delimiter, [split_by_each], [remove_empty_text])',
    description: 'Divides text by a delimiter into separate cells',
    params: {
      text: 'Text to split',
      delimiter: 'Character(s) to split by',
      split_by_each: 'Optional. TRUE to split by each character in delimiter',
      remove_empty_text: 'Optional. TRUE to remove empty results'
    },
    examples: [
      'SPLIT(A1, ",")',
      'SPLIT("a-b_c", "-_", TRUE)',  // Split by - or _
      'SPLIT(A1, " ", FALSE, TRUE)'  // Remove empty strings
    ],
    tips: [
      'Results spill horizontally',
      'Use TRANSPOSE to make vertical',
      'Use INDEX to get specific part: INDEX(SPLIT(A1,","),1,2)'
    ],
    relatedFunctions: ['TEXTJOIN', 'JOIN', 'LEFT', 'RIGHT', 'MID']
  },
  
  // ========== EXTRACTING TEXT ==========
  
  LEFT: {
    syntax: 'LEFT(string, [number_of_characters])',
    description: 'Returns the leftmost characters from a text string',
    params: {
      string: 'The text string',
      number_of_characters: 'Optional. Number of characters (default: 1)'
    },
    examples: [
      'LEFT(A1, 3)',  // First 3 characters
      'LEFT("Hello", 2)'  // Returns "He"
    ],
    tips: [
      'Returns empty if string is empty',
      'Returns entire string if number > length'
    ],
    relatedFunctions: ['RIGHT', 'MID', 'LEFTB']
  },
  
  RIGHT: {
    syntax: 'RIGHT(string, [number_of_characters])',
    description: 'Returns the rightmost characters from a text string',
    params: {
      string: 'The text string',
      number_of_characters: 'Optional. Number of characters (default: 1)'
    },
    examples: [
      'RIGHT(A1, 4)',  // Last 4 characters
      'RIGHT("Hello", 2)'  // Returns "lo"
    ],
    tips: [
      'Useful for extracting file extensions',
      'Combine with LEN for dynamic extraction'
    ],
    relatedFunctions: ['LEFT', 'MID', 'RIGHTB']
  },
  
  MID: {
    syntax: 'MID(string, starting_at, extract_length)',
    description: 'Returns a segment of a string from the middle',
    params: {
      string: 'The text string',
      starting_at: 'Starting position (1-indexed)',
      extract_length: 'Number of characters to extract'
    },
    examples: [
      'MID(A1, 3, 5)',  // 5 chars starting from position 3
      'MID("Hello World", 7, 5)'  // Returns "World"
    ],
    tips: [
      'Position is 1-indexed (first char is 1)',
      'Use FIND/SEARCH to locate start position'
    ],
    relatedFunctions: ['LEFT', 'RIGHT', 'MIDB']
  },
  
  // ========== FINDING & SEARCHING ==========
  
  FIND: {
    syntax: 'FIND(search_for, text_to_search, [starting_at])',
    description: 'Returns the position of a string within another (case-sensitive)',
    params: {
      search_for: 'Text to find',
      text_to_search: 'Text to search in',
      starting_at: 'Optional. Position to start searching (default: 1)'
    },
    examples: [
      'FIND("@", A1)',  // Position of @ in email
      'FIND("a", "Banana", 3)'  // Find "a" starting from position 3
    ],
    tips: [
      'Case-sensitive (use SEARCH for case-insensitive)',
      'Returns #VALUE! if not found',
      'Use IFERROR to handle not found'
    ],
    relatedFunctions: ['SEARCH', 'MATCH', 'REGEXMATCH']
  },
  
  SEARCH: {
    syntax: 'SEARCH(search_for, text_to_search, [starting_at])',
    description: 'Returns the position of a string within another (case-insensitive)',
    params: {
      search_for: 'Text to find (can include wildcards)',
      text_to_search: 'Text to search in',
      starting_at: 'Optional. Position to start searching (default: 1)'
    },
    examples: [
      'SEARCH("world", "Hello World")',  // Returns 7
      'SEARCH("*@*", A1)',  // Wildcard search
      'ISNUMBER(SEARCH("error", A1))'  // Check if contains "error"
    ],
    tips: [
      'Case-insensitive',
      'Supports wildcards: * (any chars) and ? (single char)',
      'Use ISNUMBER(SEARCH(...)) to check if text contains string'
    ],
    relatedFunctions: ['FIND', 'REGEXMATCH', 'MATCH']
  },
  
  // ========== REGEX FUNCTIONS ==========
  
  REGEXMATCH: {
    syntax: 'REGEXMATCH(text, regular_expression)',
    description: 'Tests whether text matches a regular expression',
    params: {
      text: 'Text to test',
      regular_expression: 'RE2 regular expression pattern'
    },
    examples: [
      'REGEXMATCH(A1, "^[0-9]+$")',  // Is all digits?
      'REGEXMATCH(A1, "[a-zA-Z]")',  // Contains letters?
      'REGEXMATCH(A1, "\\d{3}-\\d{4}")'  // Phone pattern
    ],
    tips: [
      'Uses RE2 regex syntax',
      'Returns TRUE or FALSE',
      'Great for validation'
    ],
    relatedFunctions: ['REGEXEXTRACT', 'REGEXREPLACE', 'SEARCH']
  },
  
  REGEXEXTRACT: {
    syntax: 'REGEXEXTRACT(text, regular_expression)',
    description: 'Extracts the first matching substring using a regex',
    params: {
      text: 'Text to extract from',
      regular_expression: 'RE2 pattern (use groups to extract specific parts)'
    },
    examples: [
      'REGEXEXTRACT(A1, "[0-9]+")',  // First number sequence
      'REGEXEXTRACT(A1, "[\\w.-]+@[\\w.-]+")',  // Extract email
      'REGEXEXTRACT(A1, "ID:\\s*(\\d+)")'  // Extract ID number (using group)
    ],
    tips: [
      'Use parentheses () to capture specific groups',
      'Returns #N/A if no match',
      'Can extract multiple groups into columns'
    ],
    relatedFunctions: ['REGEXMATCH', 'REGEXREPLACE', 'MID']
  },
  
  REGEXREPLACE: {
    syntax: 'REGEXREPLACE(text, regular_expression, replacement)',
    description: 'Replaces text matching a regex pattern',
    params: {
      text: 'Text to modify',
      regular_expression: 'RE2 pattern to find',
      replacement: 'Replacement text (can use $1, $2 for groups)'
    },
    examples: [
      'REGEXREPLACE(A1, "[0-9]", "")',  // Remove all digits
      'REGEXREPLACE(A1, "\\s+", " ")',  // Collapse whitespace
      'REGEXREPLACE(A1, "(\\d{3})(\\d{4})", "$1-$2")'  // Format phone
    ],
    tips: [
      'Use $1, $2 to reference captured groups',
      'Replaces ALL matches (not just first)',
      'Use ^ and $ for whole string matching'
    ],
    relatedFunctions: ['REGEXMATCH', 'REGEXEXTRACT', 'SUBSTITUTE']
  },
  
  // ========== REPLACING & SUBSTITUTING ==========
  
  SUBSTITUTE: {
    syntax: 'SUBSTITUTE(text_to_search, search_for, replace_with, [occurrence_number])',
    description: 'Replaces existing text with new text in a string',
    params: {
      text_to_search: 'Text to modify',
      search_for: 'Text to find',
      replace_with: 'Replacement text',
      occurrence_number: 'Optional. Which occurrence to replace (default: all)'
    },
    examples: [
      'SUBSTITUTE(A1, " ", "")',  // Remove all spaces
      'SUBSTITUTE(A1, ",", ";")',  // Replace commas with semicolons
      'SUBSTITUTE(A1, "old", "new", 1)'  // Replace first occurrence only
    ],
    tips: [
      'Case-sensitive',
      'Use REGEXREPLACE for pattern-based replacement',
      'Chain multiple SUBSTITUTEs for multiple replacements'
    ],
    relatedFunctions: ['REPLACE', 'REGEXREPLACE']
  },
  
  REPLACE: {
    syntax: 'REPLACE(text, position, length, new_text)',
    description: 'Replaces a part of text based on position',
    params: {
      text: 'Original text',
      position: 'Starting position (1-indexed)',
      length: 'Number of characters to replace',
      new_text: 'Replacement text'
    },
    examples: [
      'REPLACE(A1, 1, 3, "XXX")',  // Replace first 3 chars
      'REPLACE(A1, FIND("@", A1), 100, "@newdomain.com")'
    ],
    tips: [
      'Position-based (use SUBSTITUTE for text-based)',
      'new_text can be different length than replaced text'
    ],
    relatedFunctions: ['SUBSTITUTE', 'REGEXREPLACE', 'MID']
  },
  
  // ========== CASE CONVERSION ==========
  
  UPPER: {
    syntax: 'UPPER(text)',
    description: 'Converts text to uppercase',
    params: {
      text: 'Text to convert'
    },
    examples: [
      'UPPER(A1)',
      'UPPER("hello")'  // Returns "HELLO"
    ],
    relatedFunctions: ['LOWER', 'PROPER']
  },
  
  LOWER: {
    syntax: 'LOWER(text)',
    description: 'Converts text to lowercase',
    params: {
      text: 'Text to convert'
    },
    examples: [
      'LOWER(A1)',
      'LOWER("HELLO")'  // Returns "hello"
    ],
    relatedFunctions: ['UPPER', 'PROPER']
  },
  
  PROPER: {
    syntax: 'PROPER(text)',
    description: 'Capitalizes first letter of each word',
    params: {
      text: 'Text to convert'
    },
    examples: [
      'PROPER(A1)',
      'PROPER("hello world")'  // Returns "Hello World"
    ],
    tips: [
      'Also lowercases other letters',
      'Works with any word separator'
    ],
    relatedFunctions: ['UPPER', 'LOWER']
  },
  
  // ========== CLEANING & FORMATTING ==========
  
  TRIM: {
    syntax: 'TRIM(text)',
    description: 'Removes leading, trailing, and extra spaces',
    params: {
      text: 'Text to trim'
    },
    examples: [
      'TRIM(A1)',
      'TRIM("  Hello   World  ")'  // Returns "Hello World"
    ],
    tips: [
      'Keeps single spaces between words',
      'Removes all types of space characters'
    ],
    relatedFunctions: ['CLEAN', 'SUBSTITUTE']
  },
  
  CLEAN: {
    syntax: 'CLEAN(text)',
    description: 'Removes non-printable characters (ASCII 0-31)',
    params: {
      text: 'Text to clean'
    },
    examples: [
      'CLEAN(A1)',
      'TRIM(CLEAN(A1))'  // Clean and trim together
    ],
    tips: [
      'Removes tabs, line breaks, etc.',
      'Often combined with TRIM'
    ],
    relatedFunctions: ['TRIM', 'SUBSTITUTE']
  },
  
  // ========== LENGTH & PADDING ==========
  
  LEN: {
    syntax: 'LEN(text)',
    description: 'Returns the length of a text string',
    params: {
      text: 'Text to measure'
    },
    examples: [
      'LEN(A1)',
      'IF(LEN(A1)>100, "Too long", "OK")'
    ],
    tips: [
      'Counts all characters including spaces',
      'Returns 0 for empty cells'
    ],
    relatedFunctions: ['LENB', 'COUNTA']
  },
  
  REPT: {
    syntax: 'REPT(text_to_repeat, number_of_repetitions)',
    description: 'Repeats text a specified number of times',
    params: {
      text_to_repeat: 'Text to repeat',
      number_of_repetitions: 'Number of times to repeat'
    },
    examples: [
      'REPT("*", 10)',  // Returns "**********"
      'REPT("-", LEN(A1))',  // Underline same length as A1
      'REPT("█", A1/10)'  // Simple bar chart
    ],
    tips: [
      'Great for creating visual indicators',
      'Max result is 32,767 characters'
    ],
    relatedFunctions: ['CONCATENATE', 'CHAR']
  },
  
  // ========== COMPARISON ==========
  
  EXACT: {
    syntax: 'EXACT(string1, string2)',
    description: 'Tests whether two strings are identical (case-sensitive)',
    params: {
      string1: 'First string',
      string2: 'Second string'
    },
    examples: [
      'EXACT(A1, B1)',
      'EXACT("Hello", "hello")'  // Returns FALSE
    ],
    tips: [
      'Case-sensitive comparison',
      'Use = for case-insensitive comparison'
    ],
    relatedFunctions: ['DELTA']
  },
  
  // ========== NUMBER/TEXT CONVERSION ==========
  
  TEXT: {
    syntax: 'TEXT(number, format)',
    description: 'Converts a number to text with specified format',
    params: {
      number: 'Number or date to format',
      format: 'Format string'
    },
    examples: [
      'TEXT(A1, "$#,##0.00")',  // Currency
      'TEXT(A1, "000")',  // Leading zeros
      'TEXT(A1, "YYYY-MM-DD")',  // Date format
      'TEXT(A1, "0.00%")',  // Percentage
      'TEXT(A1, "dddd")'  // Day name
    ],
    tips: [
      'Same format codes as cell formatting',
      'Result is TEXT, not number'
    ],
    relatedFunctions: ['VALUE', 'FIXED', 'DOLLAR']
  },
  
  VALUE: {
    syntax: 'VALUE(text)',
    description: 'Converts a text string to a number',
    params: {
      text: 'Text to convert'
    },
    examples: [
      'VALUE("123")',
      'VALUE("$1,234.56")',  // Handles currency
      'VALUE(LEFT(A1, 5))'
    ],
    tips: [
      'Understands currency symbols and thousands separators',
      'Returns error if cannot convert'
    ],
    relatedFunctions: ['TEXT', 'NUMBERVALUE']
  },
  
  NUMBERVALUE: {
    syntax: 'NUMBERVALUE(text, [decimal_separator], [group_separator])',
    description: 'Converts text to number with custom separators',
    params: {
      text: 'Text to convert',
      decimal_separator: 'Optional. Character used for decimals',
      group_separator: 'Optional. Character used for thousands'
    },
    examples: [
      'NUMBERVALUE("1.234,56", ",", ".")',  // European format
      'NUMBERVALUE("1 234", "", " ")'  // Space separator
    ],
    tips: [
      'Useful for international number formats',
      'More flexible than VALUE'
    ],
    relatedFunctions: ['VALUE', 'TEXT']
  },
  
  FIXED: {
    syntax: 'FIXED(number, [decimals], [no_commas])',
    description: 'Formats a number with fixed decimal places',
    params: {
      number: 'Number to format',
      decimals: 'Optional. Number of decimal places (default: 2)',
      no_commas: 'Optional. TRUE to omit commas'
    },
    examples: [
      'FIXED(1234.567, 2)',  // Returns "1,234.57"
      'FIXED(1234.567, 2, TRUE)'  // Returns "1234.57"
    ],
    tips: [
      'Returns text, not number',
      'Rounds the number'
    ],
    relatedFunctions: ['TEXT', 'ROUND', 'DOLLAR']
  },
  
  DOLLAR: {
    syntax: 'DOLLAR(number, [decimals])',
    description: 'Formats number as currency text with dollar sign',
    params: {
      number: 'Number to format',
      decimals: 'Optional. Decimal places (default: 2)'
    },
    examples: [
      'DOLLAR(1234.56)',  // Returns "$1,234.56"
      'DOLLAR(1234.56, 0)'  // Returns "$1,235"
    ],
    tips: [
      'Uses locale settings for currency symbol',
      'Returns text, not number'
    ],
    relatedFunctions: ['TEXT', 'FIXED']
  },
  
  // ========== CHARACTER FUNCTIONS ==========
  
  CHAR: {
    syntax: 'CHAR(number)',
    description: 'Returns the character for a number (Unicode code point)',
    params: {
      number: 'Unicode code point'
    },
    examples: [
      'CHAR(65)',  // Returns "A"
      'CHAR(10)',  // Newline character
      'CHAR(9)',   // Tab character
      'CHAR(8226)'  // Bullet •
    ],
    tips: [
      'CHAR(10) = newline, CHAR(9) = tab',
      'Useful for inserting special characters in formulas'
    ],
    relatedFunctions: ['CODE', 'UNICHAR', 'UNICODE']
  },
  
  CODE: {
    syntax: 'CODE(string)',
    description: 'Returns the numeric code for the first character',
    params: {
      string: 'Text string'
    },
    examples: [
      'CODE("A")',  // Returns 65
      'CODE(A1)'
    ],
    tips: [
      'Returns code of FIRST character only',
      'Useful for sorting or categorizing'
    ],
    relatedFunctions: ['CHAR', 'UNICODE', 'UNICHAR']
  },
  
  UNICHAR: {
    syntax: 'UNICHAR(number)',
    description: 'Returns the Unicode character for a number',
    params: {
      number: 'Unicode code point'
    },
    examples: [
      'UNICHAR(128512)',  // 😀 emoji
      'UNICHAR(9733)',    // ★ star
      'UNICHAR(10004)'    // ✔ checkmark
    ],
    tips: [
      'Can display emojis and special symbols',
      'Full Unicode support'
    ],
    relatedFunctions: ['UNICODE', 'CHAR', 'CODE']
  },
  
  UNICODE: {
    syntax: 'UNICODE(text)',
    description: 'Returns the Unicode code point of the first character',
    params: {
      text: 'Text string'
    },
    examples: [
      'UNICODE("😀")',  // Returns 128512
      'UNICODE("★")'   // Returns 9733
    ],
    tips: [
      'Returns code of FIRST character only',
      'Works with emojis and symbols'
    ],
    relatedFunctions: ['UNICHAR', 'CODE', 'CHAR']
  },
  
  // ========== BYTE-BASED FUNCTIONS ==========
  
  LENB: {
    syntax: 'LENB(text)',
    description: 'Returns length in bytes (for double-byte languages)',
    params: {
      text: 'Text to measure'
    },
    examples: [
      'LENB("Hello")',  // Same as LEN for ASCII
      'LENB("日本語")'   // Different for CJK characters
    ],
    tips: [
      'Useful for languages like Chinese, Japanese, Korean',
      'ASCII chars = 1 byte, CJK chars = 2+ bytes'
    ],
    relatedFunctions: ['LEN', 'LEFTB', 'RIGHTB', 'MIDB']
  },
  
  LEFTB: {
    syntax: 'LEFTB(text, [num_bytes])',
    description: 'Returns leftmost bytes from text',
    params: {
      text: 'Text string',
      num_bytes: 'Number of bytes'
    },
    examples: [
      'LEFTB("日本語", 4)'
    ],
    relatedFunctions: ['LEFT', 'RIGHTB', 'MIDB']
  },
  
  RIGHTB: {
    syntax: 'RIGHTB(text, [num_bytes])',
    description: 'Returns rightmost bytes from text',
    params: {
      text: 'Text string',
      num_bytes: 'Number of bytes'
    },
    examples: [
      'RIGHTB("日本語", 4)'
    ],
    relatedFunctions: ['RIGHT', 'LEFTB', 'MIDB']
  },
  
  MIDB: {
    syntax: 'MIDB(text, start_byte, num_bytes)',
    description: 'Returns bytes from middle of text',
    params: {
      text: 'Text string',
      start_byte: 'Starting byte position',
      num_bytes: 'Number of bytes'
    },
    examples: [
      'MIDB("日本語", 3, 4)'
    ],
    relatedFunctions: ['MID', 'LEFTB', 'RIGHTB']
  },
  
  // ========== SPECIAL TEXT FUNCTIONS ==========
  
  T: {
    syntax: 'T(value)',
    description: 'Returns text if value is text, empty string otherwise',
    params: {
      value: 'Value to test'
    },
    examples: [
      'T(A1)',  // Returns A1 if text, "" if number
      'T(123)'  // Returns ""
    ],
    tips: [
      'Useful for ensuring text type',
      'Returns empty string for numbers, errors, booleans'
    ],
    relatedFunctions: ['N', 'ISTEXT', 'ISNUMBER']
  },
  
  ARABIC: {
    syntax: 'ARABIC(roman_numeral)',
    description: 'Converts Roman numerals to Arabic numbers',
    params: {
      roman_numeral: 'Roman numeral text (e.g., "XIV")'
    },
    examples: [
      'ARABIC("XIV")',  // Returns 14
      'ARABIC("MCMXCIX")'  // Returns 1999
    ],
    relatedFunctions: ['ROMAN']
  },
  
  ROMAN: {
    syntax: 'ROMAN(number, [form])',
    description: 'Converts Arabic numbers to Roman numerals',
    params: {
      number: 'Number to convert (1-3999)',
      form: 'Optional. 0-4 for different formats (0 is classic)'
    },
    examples: [
      'ROMAN(14)',  // Returns "XIV"
      'ROMAN(1999)'  // Returns "MCMXCIX"
    ],
    relatedFunctions: ['ARABIC']
  },
  
  ASC: {
    syntax: 'ASC(text)',
    description: 'Converts full-width to half-width ASCII characters',
    params: {
      text: 'Text to convert'
    },
    examples: [
      'ASC("ＡＢＣＤ")'  // Returns "ABCD"
    ],
    tips: [
      'Useful for normalizing Japanese/Chinese text',
      'Full-width characters common in CJK text'
    ],
    relatedFunctions: ['JIS', 'CLEAN']
  }
};

/**
 * Get text function info
 */
function getTextFunction(name) {
  return TEXT_FUNCTIONS[name.toUpperCase()] || null;
}

/**
 * Get all text function names
 */
function getTextFunctionNames() {
  return Object.keys(TEXT_FUNCTIONS);
}

Logger.log('📦 FormulaCatalog_Text loaded - ' + Object.keys(TEXT_FUNCTIONS).length + ' functions');
