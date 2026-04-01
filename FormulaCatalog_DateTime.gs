/**
 * @file FormulaCatalog_DateTime.gs
 * @description Date & Time Functions for Google Sheets
 * 
 * COMPLETE COVERAGE of all date and time manipulation functions.
 */

var DATETIME_FUNCTIONS = {
  
  // ========== CURRENT DATE/TIME ==========
  
  TODAY: {
    syntax: 'TODAY()',
    description: 'Returns the current date',
    params: {},
    examples: [
      'TODAY()',
      'TODAY() - A1',  // Days since date in A1
      'YEAR(TODAY())'  // Current year
    ],
    tips: [
      'Updates automatically each day',
      'Does NOT include time',
      'Volatile - recalculates on changes'
    ],
    relatedFunctions: ['NOW', 'DATE']
  },
  
  NOW: {
    syntax: 'NOW()',
    description: 'Returns the current date and time',
    params: {},
    examples: [
      'NOW()',
      'INT(NOW())',  // Today's date only
      'NOW() - INT(NOW())'  // Current time only
    ],
    tips: [
      'Updates on every recalculation',
      'Use TODAY() if you only need date',
      'Volatile function'
    ],
    relatedFunctions: ['TODAY', 'TIME']
  },
  
  // ========== DATE CREATION ==========
  
  DATE: {
    syntax: 'DATE(year, month, day)',
    description: 'Creates a date from year, month, day components',
    params: {
      year: 'Year (1900-9999)',
      month: 'Month (1-12)',
      day: 'Day (1-31)'
    },
    examples: [
      'DATE(2024, 12, 25)',
      'DATE(YEAR(A1), MONTH(A1), 1)',  // First of month
      'DATE(2024, 1, 0)',  // Dec 31, 2023 (day 0 = last day of prev month)
      'DATE(2024, 13, 1)'  // Jan 1, 2025 (handles overflow)
    ],
    tips: [
      'Month/day overflow is handled (month 13 = Jan next year)',
      'Day 0 = last day of previous month',
      'Useful for date arithmetic'
    ],
    relatedFunctions: ['DATEVALUE', 'TIME', 'YEAR', 'MONTH', 'DAY']
  },
  
  DATEVALUE: {
    syntax: 'DATEVALUE(date_string)',
    description: 'Converts a date string to a date serial number',
    params: {
      date_string: 'Text representing a date'
    },
    examples: [
      'DATEVALUE("2024-12-25")',
      'DATEVALUE("December 25, 2024")',
      'DATEVALUE("25/12/2024")'  // Depends on locale
    ],
    tips: [
      'Understands many date formats',
      'Result is a number (date serial)',
      'Use for converting imported text dates'
    ],
    relatedFunctions: ['DATE', 'VALUE', 'TIMEVALUE']
  },
  
  TIME: {
    syntax: 'TIME(hour, minute, second)',
    description: 'Creates a time from hour, minute, second components',
    params: {
      hour: 'Hour (0-23)',
      minute: 'Minute (0-59)',
      second: 'Second (0-59)'
    },
    examples: [
      'TIME(14, 30, 0)',   // 2:30 PM
      'TIME(9, 0, 0)',     // 9:00 AM
      'TIME(25, 0, 0)'     // 1:00 AM next day
    ],
    tips: [
      'Result is decimal (fraction of day)',
      'Hour overflow wraps to next day'
    ],
    relatedFunctions: ['TIMEVALUE', 'HOUR', 'MINUTE', 'SECOND']
  },
  
  TIMEVALUE: {
    syntax: 'TIMEVALUE(time_string)',
    description: 'Converts a time string to a decimal number',
    params: {
      time_string: 'Text representing a time'
    },
    examples: [
      'TIMEVALUE("2:30 PM")',
      'TIMEVALUE("14:30:00")',
      'TIMEVALUE("9:00")'
    ],
    tips: [
      'Result is fraction of day',
      '0.5 = 12:00 PM (noon)'
    ],
    relatedFunctions: ['TIME', 'DATEVALUE']
  },
  
  // ========== DATE COMPONENTS ==========
  
  YEAR: {
    syntax: 'YEAR(date)',
    description: 'Extracts the year from a date',
    params: {
      date: 'Date value'
    },
    examples: [
      'YEAR(TODAY())',
      'YEAR(A1)',
      'YEAR("2024-12-25")'
    ],
    relatedFunctions: ['MONTH', 'DAY', 'DATE']
  },
  
  MONTH: {
    syntax: 'MONTH(date)',
    description: 'Extracts the month from a date (1-12)',
    params: {
      date: 'Date value'
    },
    examples: [
      'MONTH(TODAY())',
      'TEXT(DATE(2024, A1, 1), "MMMM")'  // Month name
    ],
    relatedFunctions: ['YEAR', 'DAY', 'DATE']
  },
  
  DAY: {
    syntax: 'DAY(date)',
    description: 'Extracts the day of month from a date (1-31)',
    params: {
      date: 'Date value'
    },
    examples: [
      'DAY(TODAY())',
      'DAY(A1)'
    ],
    relatedFunctions: ['YEAR', 'MONTH', 'WEEKDAY']
  },
  
  HOUR: {
    syntax: 'HOUR(time)',
    description: 'Extracts the hour from a time (0-23)',
    params: {
      time: 'Time value'
    },
    examples: [
      'HOUR(NOW())',
      'HOUR(A1)'
    ],
    relatedFunctions: ['MINUTE', 'SECOND', 'TIME']
  },
  
  MINUTE: {
    syntax: 'MINUTE(time)',
    description: 'Extracts the minute from a time (0-59)',
    params: {
      time: 'Time value'
    },
    examples: [
      'MINUTE(NOW())',
      'MINUTE(A1)'
    ],
    relatedFunctions: ['HOUR', 'SECOND', 'TIME']
  },
  
  SECOND: {
    syntax: 'SECOND(time)',
    description: 'Extracts the second from a time (0-59)',
    params: {
      time: 'Time value'
    },
    examples: [
      'SECOND(NOW())',
      'SECOND(A1)'
    ],
    relatedFunctions: ['HOUR', 'MINUTE', 'TIME']
  },
  
  // ========== WEEKDAY & WEEK ==========
  
  WEEKDAY: {
    syntax: 'WEEKDAY(date, [type])',
    description: 'Returns the day of week as a number',
    params: {
      date: 'Date value',
      type: 'Optional. Numbering system (1=Sun-Sat 1-7, 2=Mon-Sun 1-7, 3=Mon-Sun 0-6)'
    },
    examples: [
      'WEEKDAY(TODAY())',      // Sunday=1, Saturday=7
      'WEEKDAY(TODAY(), 2)',   // Monday=1, Sunday=7
      'TEXT(A1, "dddd")'       // Day name instead
    ],
    tips: [
      'Type 1 (default): Sunday=1 to Saturday=7',
      'Type 2: Monday=1 to Sunday=7',
      'Type 3: Monday=0 to Sunday=6'
    ],
    relatedFunctions: ['WEEKNUM', 'ISOWEEKNUM', 'TEXT']
  },
  
  WEEKNUM: {
    syntax: 'WEEKNUM(date, [type])',
    description: 'Returns the week number of the year',
    params: {
      date: 'Date value',
      type: 'Optional. Week start day (1=Sunday, 2=Monday)'
    },
    examples: [
      'WEEKNUM(TODAY())',
      'WEEKNUM(A1, 2)'   // Week starts Monday
    ],
    tips: [
      'Week 1 contains January 1',
      'Use ISOWEEKNUM for ISO standard'
    ],
    relatedFunctions: ['ISOWEEKNUM', 'WEEKDAY']
  },
  
  ISOWEEKNUM: {
    syntax: 'ISOWEEKNUM(date)',
    description: 'Returns the ISO week number',
    params: {
      date: 'Date value'
    },
    examples: [
      'ISOWEEKNUM(TODAY())',
      'ISOWEEKNUM(A1)'
    ],
    tips: [
      'ISO weeks start Monday',
      'Week 1 contains first Thursday',
      'May differ from WEEKNUM at year boundaries'
    ],
    relatedFunctions: ['WEEKNUM', 'WEEKDAY']
  },
  
  // ========== DATE CALCULATIONS ==========
  
  DATEDIF: {
    syntax: 'DATEDIF(start_date, end_date, unit)',
    description: 'Calculates the difference between two dates',
    params: {
      start_date: 'Start date (must be <= end_date)',
      end_date: 'End date',
      unit: 'Unit of measurement'
    },
    examples: [
      'DATEDIF(A1, TODAY(), "Y")',   // Complete years
      'DATEDIF(A1, TODAY(), "M")',   // Complete months
      'DATEDIF(A1, TODAY(), "D")',   // Days
      'DATEDIF(A1, B1, "YM")',       // Months excluding years
      'DATEDIF(A1, B1, "YD")',       // Days excluding years
      'DATEDIF(A1, B1, "MD")'        // Days excluding years and months
    ],
    tips: [
      'Units: "Y"=years, "M"=months, "D"=days',
      '"YM"=months ignoring years, "YD"=days ignoring years, "MD"=days ignoring months/years',
      'Start must be <= end (no negative results)'
    ],
    relatedFunctions: ['DAYS', 'MONTHS', 'YEARFRAC']
  },
  
  DAYS: {
    syntax: 'DAYS(end_date, start_date)',
    description: 'Returns the number of days between two dates',
    params: {
      end_date: 'End date',
      start_date: 'Start date'
    },
    examples: [
      'DAYS(B1, A1)',
      'DAYS(TODAY(), A1)',  // Days since A1
      'DAYS("2025-01-01", TODAY())'  // Days until date
    ],
    tips: [
      'Can return negative if end < start',
      'Simpler than DATEDIF for just days'
    ],
    relatedFunctions: ['DATEDIF', 'NETWORKDAYS', 'WORKDAY']
  },
  
  DAYS360: {
    syntax: 'DAYS360(start_date, end_date, [method])',
    description: 'Days between dates based on 360-day year',
    params: {
      start_date: 'Start date',
      end_date: 'End date',
      method: 'Optional. FALSE=US/NASD (default), TRUE=European'
    },
    examples: [
      'DAYS360(A1, B1)',
      'DAYS360(A1, B1, TRUE)'  // European method
    ],
    tips: [
      'Used in some financial calculations',
      'Assumes 30-day months, 360-day year'
    ],
    relatedFunctions: ['DAYS', 'YEARFRAC']
  },
  
  YEARFRAC: {
    syntax: 'YEARFRAC(start_date, end_date, [day_count_convention])',
    description: 'Returns fraction of year between dates',
    params: {
      start_date: 'Start date',
      end_date: 'End date',
      day_count_convention: 'Optional. 0-4 for different methods'
    },
    examples: [
      'YEARFRAC(A1, B1)',
      'YEARFRAC(A1, B1, 1)'  // Actual/actual
    ],
    tips: [
      '0 = 30/360 US (default)',
      '1 = Actual/actual',
      '2 = Actual/360',
      '3 = Actual/365',
      '4 = 30/360 European'
    ],
    relatedFunctions: ['DATEDIF', 'DAYS360']
  },
  
  // ========== DATE SHIFTING ==========
  
  EDATE: {
    syntax: 'EDATE(start_date, months)',
    description: 'Returns date that is specified months away',
    params: {
      start_date: 'Starting date',
      months: 'Number of months (positive=future, negative=past)'
    },
    examples: [
      'EDATE(TODAY(), 1)',    // Next month
      'EDATE(A1, -3)',        // 3 months ago
      'EDATE(A1, 12)'         // One year later
    ],
    tips: [
      'Same day of month (adjusted if needed)',
      'Handles month-end properly'
    ],
    relatedFunctions: ['EOMONTH', 'DATE']
  },
  
  EOMONTH: {
    syntax: 'EOMONTH(start_date, months)',
    description: 'Returns last day of month that is months away',
    params: {
      start_date: 'Starting date',
      months: 'Months to offset'
    },
    examples: [
      'EOMONTH(TODAY(), 0)',   // Last day of this month
      'EOMONTH(A1, 1)',        // Last day of next month
      'EOMONTH(A1, -1)',       // Last day of previous month
      'EOMONTH(A1, 0) + 1'     // First day of next month
    ],
    tips: [
      'Great for month-end reporting',
      'EOMONTH(date, 0)+1 = first of next month'
    ],
    relatedFunctions: ['EDATE', 'DATE', 'DAY']
  },
  
  // ========== WORK DAYS ==========
  
  WORKDAY: {
    syntax: 'WORKDAY(start_date, num_days, [holidays])',
    description: 'Returns date after specified working days',
    params: {
      start_date: 'Starting date',
      num_days: 'Number of working days',
      holidays: 'Optional. Range of dates to exclude'
    },
    examples: [
      'WORKDAY(TODAY(), 5)',  // 5 business days from now
      'WORKDAY(A1, 10, Holidays)',  // Excluding holidays
      'WORKDAY(A1, -3)'  // 3 business days ago
    ],
    tips: [
      'Excludes Saturdays and Sundays',
      'Can add custom holidays',
      'Negative days go backward'
    ],
    relatedFunctions: ['WORKDAY.INTL', 'NETWORKDAYS']
  },
  
  WORKDAY_INTL: {
    syntax: 'WORKDAY.INTL(start_date, num_days, [weekend], [holidays])',
    description: 'Like WORKDAY but with custom weekend days',
    params: {
      start_date: 'Starting date',
      num_days: 'Number of working days',
      weekend: 'Optional. Weekend definition (1=Sat-Sun, 2=Sun-Mon, etc., or "0000011")',
      holidays: 'Optional. Holiday dates'
    },
    examples: [
      'WORKDAY.INTL(A1, 5, "0000011")',  // Sat-Sun weekend
      'WORKDAY.INTL(A1, 5, 7)',  // Friday only weekend
      'WORKDAY.INTL(A1, 5, "1000001")'  // Sun and Mon off
    ],
    tips: [
      'weekend string: 1=non-working, 0=working (Mon-Sun)',
      'Or use number codes 1-17'
    ],
    relatedFunctions: ['WORKDAY', 'NETWORKDAYS.INTL']
  },
  
  NETWORKDAYS: {
    syntax: 'NETWORKDAYS(start_date, end_date, [holidays])',
    description: 'Returns number of working days between dates',
    params: {
      start_date: 'Start date',
      end_date: 'End date',
      holidays: 'Optional. Holiday dates to exclude'
    },
    examples: [
      'NETWORKDAYS(A1, B1)',
      'NETWORKDAYS(A1, A1+30, Holidays)',
      'NETWORKDAYS(TODAY(), EOMONTH(TODAY(), 0))'
    ],
    tips: [
      'Includes both start and end dates',
      'Excludes Saturdays and Sundays'
    ],
    relatedFunctions: ['NETWORKDAYS.INTL', 'WORKDAY']
  },
  
  NETWORKDAYS_INTL: {
    syntax: 'NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])',
    description: 'Like NETWORKDAYS with custom weekends',
    params: {
      start_date: 'Start date',
      end_date: 'End date',
      weekend: 'Optional. Weekend definition',
      holidays: 'Optional. Holiday dates'
    },
    examples: [
      'NETWORKDAYS.INTL(A1, B1, "0000011")',
      'NETWORKDAYS.INTL(A1, B1, 11)'  // Sunday only weekend
    ],
    relatedFunctions: ['NETWORKDAYS', 'WORKDAY.INTL']
  },
  
  // ========== DATE INFO ==========
  
  ISDATE: {
    syntax: 'ISDATE(value)',
    description: 'Returns TRUE if value is a valid date',
    params: {
      value: 'Value to check'
    },
    examples: [
      'ISDATE(A1)',
      'ISDATE("2024-12-25")',
      'ISDATE("not a date")'  // FALSE
    ],
    tips: [
      'Google Sheets specific',
      'Checks if value can be interpreted as date'
    ],
    relatedFunctions: ['ISNUMBER', 'DATEVALUE']
  },
  
  // ========== EPOCH/UNIX TIME ==========
  
  EPOCHTODATE: {
    syntax: 'EPOCHTODATE(timestamp, [unit])',
    description: 'Converts Unix timestamp to date',
    params: {
      timestamp: 'Unix timestamp',
      unit: 'Optional. 0=seconds (default), 1=milliseconds'
    },
    examples: [
      'EPOCHTODATE(1703462400)',        // Unix seconds
      'EPOCHTODATE(1703462400000, 1)'   // Unix milliseconds
    ],
    tips: [
      'Google Sheets specific function',
      'Unix timestamp = seconds since Jan 1, 1970'
    ],
    relatedFunctions: ['DATETOEPOCH']
  },
  
  // ========== FORMATTING (using TEXT) ==========
  
  // Note: TEXT function is in Text catalog but common date formats documented here
  DATE_FORMATS: {
    description: 'Common date format codes for TEXT function',
    formats: {
      'YYYY-MM-DD': 'ISO date: 2024-12-25',
      'MM/DD/YYYY': 'US date: 12/25/2024',
      'DD/MM/YYYY': 'European date: 25/12/2024',
      'MMMM D, YYYY': 'Long date: December 25, 2024',
      'MMM D': 'Short: Dec 25',
      'dddd': 'Day name: Wednesday',
      'ddd': 'Short day: Wed',
      'MMMM': 'Month name: December',
      'MMM': 'Short month: Dec',
      'HH:MM:SS': 'Time 24hr: 14:30:00',
      'H:MM AM/PM': 'Time 12hr: 2:30 PM',
      'YYYY-MM-DD HH:MM': 'Date and time'
    },
    examples: [
      'TEXT(A1, "YYYY-MM-DD")',
      'TEXT(A1, "dddd, MMMM D, YYYY")',
      'TEXT(A1, "HH:MM:SS")'
    ]
  }
};

/**
 * Get datetime function info
 */
function getDateTimeFunction(name) {
  return DATETIME_FUNCTIONS[name.toUpperCase()] || null;
}

/**
 * Get all datetime function names
 */
function getDateTimeFunctionNames() {
  return Object.keys(DATETIME_FUNCTIONS);
}

Logger.log('📦 FormulaCatalog_DateTime loaded - ' + Object.keys(DATETIME_FUNCTIONS).length + ' functions');
