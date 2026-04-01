/**
 * @file FormulaCatalog_Google.gs
 * @description Google Sheets Specific & Web Functions
 * 
 * COMPLETE COVERAGE of Google-exclusive functions including:
 * - QUERY (the most powerful function)
 * - Import functions (IMPORTRANGE, IMPORTHTML, IMPORTXML, etc.)
 * - GOOGLEFINANCE
 * - SPARKLINE
 * - GOOGLETRANSLATE
 * - IMAGE
 */

var GOOGLE_FUNCTIONS = {
  
  // ========== QUERY - THE POWERHOUSE ==========
  
  QUERY: {
    syntax: 'QUERY(data, query, [headers])',
    description: 'Runs a Google Visualization API Query Language query on data',
    params: {
      data: 'Range or array to query',
      query: 'Query string in Google Query Language (SQL-like)',
      headers: 'Optional. Number of header rows (default: auto-detect)'
    },
    examples: [
      // Basic SELECT
      'QUERY(A:E, "SELECT A, B, C")',
      'QUERY(A:E, "SELECT * WHERE B > 100")',
      
      // Filtering
      'QUERY(A:E, "SELECT A, B WHERE C = \'Apple\'")',
      'QUERY(A:E, "SELECT * WHERE B > 100 AND C = \'Active\'")',
      'QUERY(A:E, "SELECT * WHERE A CONTAINS \'test\'")',
      'QUERY(A:E, "SELECT * WHERE A MATCHES \'.*@gmail\\.com\'")',
      
      // Aggregation
      'QUERY(A:E, "SELECT A, SUM(B) GROUP BY A")',
      'QUERY(A:E, "SELECT A, AVG(B), COUNT(C) GROUP BY A")',
      'QUERY(A:E, "SELECT A, SUM(B) GROUP BY A PIVOT C")',
      
      // Sorting and Limiting
      'QUERY(A:E, "SELECT * ORDER BY B DESC")',
      'QUERY(A:E, "SELECT * ORDER BY A ASC, B DESC LIMIT 10")',
      'QUERY(A:E, "SELECT * LIMIT 10 OFFSET 5")',
      
      // Formatting
      'QUERY(A:E, "SELECT A, B FORMAT B \'$#,##0\'")',
      'QUERY(A:E, "SELECT A, C FORMAT C \'YYYY-MM-DD\'")',
      
      // Labels
      'QUERY(A:E, "SELECT A, SUM(B) GROUP BY A LABEL SUM(B) \'Total Sales\'")',
      
      // With cell references
      'QUERY(A:E, "SELECT * WHERE B > "&D1)',
      'QUERY(A:E, "SELECT * WHERE A = \'"&F1&"\'")'
    ],
    tips: [
      'Most powerful function in Google Sheets',
      'Use single quotes for strings IN the query',
      'Column references: Col1, Col2 or A, B, C...',
      'Functions: SUM, AVG, COUNT, MAX, MIN, YEAR, MONTH, etc.',
      'Operators: =, !=, <, >, <=, >=, CONTAINS, STARTS WITH, ENDS WITH, MATCHES (regex), LIKE',
      'Combine with IMPORTRANGE for cross-sheet queries'
    ],
    queryLanguage: {
      SELECT: 'Choose columns: SELECT A, B, SUM(C)',
      WHERE: 'Filter rows: WHERE A > 100',
      GROUP_BY: 'Aggregate: GROUP BY A',
      PIVOT: 'Create pivot: PIVOT B',
      ORDER_BY: 'Sort: ORDER BY A DESC',
      LIMIT: 'Row limit: LIMIT 10',
      OFFSET: 'Skip rows: OFFSET 5',
      LABEL: 'Rename columns: LABEL A "Name"',
      FORMAT: 'Format output: FORMAT A "#,##0"',
      OPTIONS: 'no_values to hide data rows'
    },
    whereClauses: {
      '=': 'Equals',
      '!=': 'Not equals',
      '<': 'Less than',
      '>': 'Greater than',
      '<=': 'Less than or equal',
      '>=': 'Greater than or equal',
      'CONTAINS': 'Contains substring',
      'STARTS WITH': 'Begins with',
      'ENDS WITH': 'Ends with',
      'MATCHES': 'Regex match',
      'LIKE': 'Wildcard match (% = any, _ = single char)',
      'IS NULL': 'Is empty',
      'IS NOT NULL': 'Is not empty'
    },
    aggregateFunctions: ['SUM', 'AVG', 'COUNT', 'MAX', 'MIN'],
    scalarFunctions: ['YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND', 'QUARTER', 'DAYOFWEEK', 'NOW', 'DATEDIFF', 'TODATE', 'UPPER', 'LOWER'],
    relatedFunctions: ['FILTER', 'SORT', 'UNIQUE']
  },
  
  // ========== IMPORT FUNCTIONS ==========
  
  IMPORTRANGE: {
    syntax: 'IMPORTRANGE(spreadsheet_url, range_string)',
    description: 'Imports data from another Google Sheets spreadsheet',
    params: {
      spreadsheet_url: 'URL of the source spreadsheet',
      range_string: 'Range to import (e.g., "Sheet1!A1:D10")'
    },
    examples: [
      'IMPORTRANGE("https://docs.google.com/spreadsheets/d/xxx", "Sheet1!A:D")',
      'IMPORTRANGE(A1, "Data!A1:Z100")',  // URL in cell A1
      'QUERY(IMPORTRANGE(url, "Sheet1!A:E"), "SELECT * WHERE Col2 > 100")'
    ],
    tips: [
      'First use requires clicking "Allow access"',
      'Once allowed, all ranges in that sheet are accessible',
      'Can combine with QUERY for cross-sheet analysis',
      'Updates automatically (may have delay)'
    ],
    relatedFunctions: ['QUERY', 'IMPORTHTML', 'IMPORTDATA']
  },
  
  IMPORTHTML: {
    syntax: 'IMPORTHTML(url, query, index)',
    description: 'Imports data from a table or list in an HTML page',
    params: {
      url: 'URL of the web page',
      query: '"table" or "list"',
      index: 'Index of the table/list (1-based)'
    },
    examples: [
      'IMPORTHTML("https://en.wikipedia.org/wiki/List_of_countries", "table", 1)',
      'IMPORTHTML(A1, "table", 2)',
      'IMPORTHTML("https://example.com", "list", 1)'
    ],
    tips: [
      'Great for importing HTML tables',
      'Index 1 = first table on page',
      'Some sites may block access'
    ],
    relatedFunctions: ['IMPORTXML', 'IMPORTDATA', 'IMPORTFEED']
  },
  
  IMPORTXML: {
    syntax: 'IMPORTXML(url, xpath_query)',
    description: 'Imports data from XML using XPath query',
    params: {
      url: 'URL of the XML/HTML page',
      xpath_query: 'XPath query to select data'
    },
    examples: [
      'IMPORTXML("https://example.com", "//h1")',  // All h1 tags
      'IMPORTXML(A1, "//table//tr")',  // All table rows
      'IMPORTXML("https://example.com", "//@href")',  // All href attributes
      'IMPORTXML("https://example.com", "//div[@class=\'price\']")',  // Div with class
      'IMPORTXML("https://example.com", "//meta[@name=\'description\']/@content")'
    ],
    tips: [
      'Powerful for web scraping',
      'Requires knowledge of XPath',
      'Works with both XML and HTML',
      'Rate limited by Google'
    ],
    commonXPath: {
      '//tag': 'All tags of type',
      '//tag[@attr="value"]': 'Tags with attribute',
      '//tag/text()': 'Text content',
      '//@attr': 'All attributes',
      '//tag[1]': 'First tag'
    },
    relatedFunctions: ['IMPORTHTML', 'IMPORTDATA']
  },
  
  IMPORTDATA: {
    syntax: 'IMPORTDATA(url)',
    description: 'Imports data from a URL (CSV or TSV format)',
    params: {
      url: 'URL of the CSV/TSV file'
    },
    examples: [
      'IMPORTDATA("https://example.com/data.csv")',
      'IMPORTDATA(A1)'
    ],
    tips: [
      'Expects CSV or TSV format',
      'Auto-detects delimiter',
      'Great for importing external data feeds'
    ],
    relatedFunctions: ['IMPORTHTML', 'IMPORTXML']
  },
  
  IMPORTFEED: {
    syntax: 'IMPORTFEED(url, [query], [headers], [num_items])',
    description: 'Imports RSS or Atom feed',
    params: {
      url: 'URL of the RSS/Atom feed',
      query: 'Optional. "feed" for feed info, "items" for items (default)',
      headers: 'Optional. Include headers',
      num_items: 'Optional. Number of items to import'
    },
    examples: [
      'IMPORTFEED("https://news.google.com/rss")',
      'IMPORTFEED(A1, "items", TRUE, 10)',  // Latest 10 with headers
      'IMPORTFEED(A1, "feed")'  // Feed metadata
    ],
    tips: [
      'Great for news/blog monitoring',
      'Returns title, URL, date, summary'
    ],
    relatedFunctions: ['IMPORTXML', 'IMPORTHTML']
  },
  
  // ========== GOOGLE FINANCE ==========
  
  GOOGLEFINANCE: {
    syntax: 'GOOGLEFINANCE(ticker, [attribute], [start_date], [end_date|num_days], [interval])',
    description: 'Retrieves current or historical stock market data',
    params: {
      ticker: 'Stock symbol (e.g., "GOOG", "NASDAQ:AAPL")',
      attribute: 'Optional. Data type to retrieve',
      start_date: 'Optional. Start date for historical data',
      end_date: 'Optional. End date or number of days',
      interval: 'Optional. "DAILY" or "WEEKLY"'
    },
    examples: [
      // Current data
      'GOOGLEFINANCE("GOOG")',  // Current price
      'GOOGLEFINANCE("GOOG", "price")',
      'GOOGLEFINANCE("NASDAQ:AAPL", "marketcap")',
      'GOOGLEFINANCE("GOOG", "pe")',  // P/E ratio
      
      // Historical data
      'GOOGLEFINANCE("GOOG", "close", DATE(2024,1,1), DATE(2024,12,31))',
      'GOOGLEFINANCE("GOOG", "close", TODAY()-365, 365)',  // Last year
      'GOOGLEFINANCE("GOOG", "close", DATE(2024,1,1), DATE(2024,12,31), "WEEKLY")',
      
      // Currency
      'GOOGLEFINANCE("CURRENCY:USDEUR")',
      'GOOGLEFINANCE("CURRENCY:GBPUSD")',
      
      // Mutual funds
      'GOOGLEFINANCE("MUTF:VFINX", "price")'
    ],
    attributes: {
      'price': 'Current price',
      'priceopen': 'Opening price',
      'high': 'Day high',
      'low': 'Day low',
      'volume': 'Trading volume',
      'marketcap': 'Market capitalization',
      'tradetime': 'Last trade time',
      'datadelay': 'Data delay in minutes',
      'volumeavg': 'Average volume',
      'pe': 'Price/Earnings ratio',
      'eps': 'Earnings per share',
      'high52': '52-week high',
      'low52': '52-week low',
      'change': 'Price change',
      'changepct': 'Price change %',
      'closeyest': 'Previous close',
      'shares': 'Shares outstanding',
      'currency': 'Currency code'
    },
    historicalAttributes: {
      'open': 'Opening price',
      'close': 'Closing price',
      'high': 'Daily high',
      'low': 'Daily low',
      'volume': 'Daily volume'
    },
    tips: [
      'Data has 15-20 minute delay',
      'Use "CURRENCY:XXXYYY" for exchange rates',
      'Historical data returns array with headers',
      'Some attributes only available for stocks'
    ],
    relatedFunctions: ['TODAY', 'IMPORTXML']
  },
  
  // ========== SPARKLINE ==========
  
  SPARKLINE: {
    syntax: 'SPARKLINE(data, [options])',
    description: 'Creates a miniature chart within a cell',
    params: {
      data: 'Range or array of numeric data',
      options: 'Optional. Object with chart options'
    },
    examples: [
      // Line chart (default)
      'SPARKLINE(B2:B10)',
      'SPARKLINE(B2:B10, {"color", "blue"; "linewidth", 2})',
      
      // Bar chart
      'SPARKLINE(A1, {"charttype", "bar"; "max", 100})',  // Progress bar
      'SPARKLINE({A1, 100-A1}, {"charttype", "bar"; "color1", "green"; "color2", "lightgray"})',
      
      // Column chart
      'SPARKLINE(B2:B10, {"charttype", "column"; "color", "green"; "negcolor", "red"})',
      
      // Win/Loss chart
      'SPARKLINE(B2:B10, {"charttype", "winloss"; "color", "green"; "negcolor", "red"})'
    ],
    chartTypes: {
      'line': 'Line chart (default)',
      'bar': 'Horizontal bar (good for progress)',
      'column': 'Vertical columns',
      'winloss': 'Win/loss indicators'
    },
    options: {
      'charttype': 'Chart type',
      'color': 'Main color',
      'negcolor': 'Negative value color',
      'linewidth': 'Line thickness',
      'max': 'Maximum value',
      'min': 'Minimum value',
      'rtl': 'Right to left',
      'empty': 'How to handle empty cells',
      'nan': 'How to handle non-numeric values',
      'color1': 'First bar color',
      'color2': 'Second bar color',
      'firstcolor': 'First point color',
      'lastcolor': 'Last point color',
      'highcolor': 'Highest point color',
      'lowcolor': 'Lowest point color'
    },
    tips: [
      'Great for dashboards',
      'Bar type perfect for progress indicators',
      'Options use array notation {"key", "value"}',
      'Can highlight first/last/high/low points'
    ],
    relatedFunctions: ['IMAGE']
  },
  
  // ========== TRANSLATION ==========
  
  GOOGLETRANSLATE: {
    syntax: 'GOOGLETRANSLATE(text, [source_language], [target_language])',
    description: 'Translates text between languages',
    params: {
      text: 'Text to translate',
      source_language: 'Optional. Source language code (default: "auto")',
      target_language: 'Optional. Target language code (default: "en")'
    },
    examples: [
      'GOOGLETRANSLATE(A1, "es", "en")',  // Spanish to English
      'GOOGLETRANSLATE(A1, "auto", "fr")',  // Auto-detect to French
      'GOOGLETRANSLATE("Hello", "en", "ja")',  // English to Japanese
      'GOOGLETRANSLATE(A1, , "de")'  // Auto to German
    ],
    languageCodes: {
      'en': 'English',
      'es': 'Spanish',
      'fr': 'French',
      'de': 'German',
      'it': 'Italian',
      'pt': 'Portuguese',
      'ja': 'Japanese',
      'ko': 'Korean',
      'zh': 'Chinese',
      'ar': 'Arabic',
      'ru': 'Russian',
      'auto': 'Auto-detect'
    },
    tips: [
      'Use "auto" for source to auto-detect',
      'Quality varies by language pair',
      'May have rate limits'
    ],
    relatedFunctions: ['DETECTLANGUAGE']
  },
  
  DETECTLANGUAGE: {
    syntax: 'DETECTLANGUAGE(text_or_range)',
    description: 'Detects the language of text',
    params: {
      text_or_range: 'Text or cell reference'
    },
    examples: [
      'DETECTLANGUAGE(A1)',
      'DETECTLANGUAGE("Bonjour")',  // Returns "fr"
      'DETECTLANGUAGE("こんにちは")'  // Returns "ja"
    ],
    tips: [
      'Returns ISO language code',
      'May need sufficient text for accuracy'
    ],
    relatedFunctions: ['GOOGLETRANSLATE']
  },
  
  // ========== IMAGE ==========
  
  IMAGE: {
    syntax: 'IMAGE(url, [mode], [height], [width])',
    description: 'Inserts an image from a URL into a cell',
    params: {
      url: 'URL of the image',
      mode: 'Optional. Sizing mode (1-4)',
      height: 'Optional. Custom height (mode 4)',
      width: 'Optional. Custom width (mode 4)'
    },
    examples: [
      'IMAGE("https://example.com/logo.png")',
      'IMAGE(A1, 1)',  // Fit to cell
      'IMAGE(A1, 2)',  // Stretch to fill
      'IMAGE(A1, 3)',  // Original size
      'IMAGE(A1, 4, 100, 100)'  // Custom 100x100
    ],
    modes: {
      1: 'Fit to cell, maintain aspect ratio (default)',
      2: 'Stretch to fill cell',
      3: 'Original size (may overflow)',
      4: 'Custom size'
    },
    tips: [
      'Works with PNG, JPG, GIF',
      'Image must be publicly accessible',
      'Great for product catalogs, avatars'
    ],
    relatedFunctions: ['SPARKLINE', 'HYPERLINK']
  },
  
  // ========== DATABASE FUNCTIONS ==========
  
  DSUM: {
    syntax: 'DSUM(database, field, criteria)',
    description: 'Sums values in a database column matching criteria',
    params: {
      database: 'Range including headers',
      field: 'Column name (text) or index (number)',
      criteria: 'Range with criteria headers and values'
    },
    examples: [
      'DSUM(A1:D100, "Sales", F1:G2)',
      'DSUM(A1:D100, 3, F1:G2)'  // 3rd column
    ],
    tips: [
      'Criteria range needs headers matching database',
      'Multiple rows in criteria = OR',
      'Multiple columns in criteria = AND'
    ],
    relatedFunctions: ['DAVERAGE', 'DCOUNT', 'SUMIFS']
  },
  
  DAVERAGE: {
    syntax: 'DAVERAGE(database, field, criteria)',
    description: 'Averages values in database column matching criteria',
    params: {
      database: 'Range including headers',
      field: 'Column name or index',
      criteria: 'Criteria range'
    },
    examples: [
      'DAVERAGE(A1:D100, "Price", F1:G2)'
    ],
    relatedFunctions: ['DSUM', 'DCOUNT', 'AVERAGEIFS']
  },
  
  DCOUNT: {
    syntax: 'DCOUNT(database, field, criteria)',
    description: 'Counts numeric values in database column matching criteria',
    params: {
      database: 'Range including headers',
      field: 'Column name or index',
      criteria: 'Criteria range'
    },
    examples: [
      'DCOUNT(A1:D100, "Quantity", F1:G2)'
    ],
    relatedFunctions: ['DCOUNTA', 'DSUM', 'COUNTIFS']
  },
  
  DCOUNTA: {
    syntax: 'DCOUNTA(database, field, criteria)',
    description: 'Counts non-empty values in database column matching criteria',
    params: {
      database: 'Range including headers',
      field: 'Column name or index',
      criteria: 'Criteria range'
    },
    examples: [
      'DCOUNTA(A1:D100, "Notes", F1:G2)'
    ],
    relatedFunctions: ['DCOUNT', 'COUNTA']
  },
  
  DMAX: {
    syntax: 'DMAX(database, field, criteria)',
    description: 'Returns maximum value matching criteria',
    params: {
      database: 'Range including headers',
      field: 'Column name or index',
      criteria: 'Criteria range'
    },
    examples: [
      'DMAX(A1:D100, "Price", F1:G2)'
    ],
    relatedFunctions: ['DMIN', 'MAXIFS']
  },
  
  DMIN: {
    syntax: 'DMIN(database, field, criteria)',
    description: 'Returns minimum value matching criteria',
    params: {
      database: 'Range including headers',
      field: 'Column name or index',
      criteria: 'Criteria range'
    },
    examples: [
      'DMIN(A1:D100, "Price", F1:G2)'
    ],
    relatedFunctions: ['DMAX', 'MINIFS']
  },
  
  DGET: {
    syntax: 'DGET(database, field, criteria)',
    description: 'Returns single value matching criteria (error if multiple)',
    params: {
      database: 'Range including headers',
      field: 'Column name or index',
      criteria: 'Criteria range'
    },
    examples: [
      'DGET(A1:D100, "Price", F1:G2)'
    ],
    tips: [
      'Returns #NUM! if multiple matches',
      'Returns #VALUE! if no matches'
    ],
    relatedFunctions: ['VLOOKUP', 'XLOOKUP', 'FILTER']
  },
  
  DPRODUCT: {
    syntax: 'DPRODUCT(database, field, criteria)',
    description: 'Multiplies values matching criteria',
    params: {
      database: 'Range including headers',
      field: 'Column name or index',
      criteria: 'Criteria range'
    },
    relatedFunctions: ['DSUM', 'PRODUCT']
  },
  
  DSTDEV: {
    syntax: 'DSTDEV(database, field, criteria)',
    description: 'Sample standard deviation of matching values',
    params: {
      database: 'Range including headers',
      field: 'Column name or index',
      criteria: 'Criteria range'
    },
    relatedFunctions: ['DSTDEVP', 'STDEV']
  },
  
  DSTDEVP: {
    syntax: 'DSTDEVP(database, field, criteria)',
    description: 'Population standard deviation of matching values',
    relatedFunctions: ['DSTDEV', 'STDEVP']
  },
  
  DVAR: {
    syntax: 'DVAR(database, field, criteria)',
    description: 'Sample variance of matching values',
    relatedFunctions: ['DVARP', 'VAR']
  },
  
  DVARP: {
    syntax: 'DVARP(database, field, criteria)',
    description: 'Population variance of matching values',
    relatedFunctions: ['DVAR', 'VARP']
  }
};

/**
 * Get Google-specific function info
 */
function getGoogleFunction(name) {
  return GOOGLE_FUNCTIONS[name.toUpperCase()] || null;
}

/**
 * Get all Google function names
 */
function getGoogleFunctionNames() {
  return Object.keys(GOOGLE_FUNCTIONS);
}

Logger.log('📦 FormulaCatalog_Google loaded - ' + Object.keys(GOOGLE_FUNCTIONS).length + ' functions');
