/**
 * @file FormulaCatalog_Engineering.gs
 * @description Engineering & Conversion Functions for Google Sheets
 * 
 * COMPLETE COVERAGE of engineering functions including:
 * - Number base conversions
 * - Unit conversions
 * - Complex numbers
 * - Bessel functions
 */

var ENGINEERING_FUNCTIONS = {
  
  // ========== NUMBER BASE CONVERSIONS ==========
  
  BIN2DEC: {
    syntax: 'BIN2DEC(binary_number)',
    description: 'Converts binary to decimal',
    params: {
      binary_number: 'Binary number (up to 10 digits)'
    },
    examples: [
      'BIN2DEC("1010")',    // Returns 10
      'BIN2DEC("11111111")',  // Returns 255
      'BIN2DEC(A1)'
    ],
    tips: [
      'Max 10 binary digits',
      '10th bit is sign bit (two\'s complement)'
    ],
    relatedFunctions: ['DEC2BIN', 'BIN2HEX', 'BIN2OCT']
  },
  
  BIN2HEX: {
    syntax: 'BIN2HEX(binary_number, [significant_digits])',
    description: 'Converts binary to hexadecimal',
    params: {
      binary_number: 'Binary number',
      significant_digits: 'Optional. Minimum digits in result'
    },
    examples: [
      'BIN2HEX("11111111")',  // Returns "FF"
      'BIN2HEX("1010", 4)'    // Returns "000A"
    ],
    relatedFunctions: ['HEX2BIN', 'BIN2DEC']
  },
  
  BIN2OCT: {
    syntax: 'BIN2OCT(binary_number, [significant_digits])',
    description: 'Converts binary to octal',
    params: {
      binary_number: 'Binary number',
      significant_digits: 'Optional. Minimum digits'
    },
    examples: [
      'BIN2OCT("1010")',  // Returns "12"
      'BIN2OCT("111", 3)' // Returns "007"
    ],
    relatedFunctions: ['OCT2BIN', 'BIN2DEC']
  },
  
  DEC2BIN: {
    syntax: 'DEC2BIN(decimal_number, [significant_digits])',
    description: 'Converts decimal to binary',
    params: {
      decimal_number: 'Decimal number (-512 to 511)',
      significant_digits: 'Optional. Minimum digits'
    },
    examples: [
      'DEC2BIN(10)',      // Returns "1010"
      'DEC2BIN(255, 8)',  // Returns "11111111"
      'DEC2BIN(-1)'       // Returns "1111111111" (two's complement)
    ],
    relatedFunctions: ['BIN2DEC', 'DEC2HEX', 'DEC2OCT']
  },
  
  DEC2HEX: {
    syntax: 'DEC2HEX(decimal_number, [significant_digits])',
    description: 'Converts decimal to hexadecimal',
    params: {
      decimal_number: 'Decimal number',
      significant_digits: 'Optional. Minimum digits'
    },
    examples: [
      'DEC2HEX(255)',      // Returns "FF"
      'DEC2HEX(16, 4)',    // Returns "0010"
      'DEC2HEX(-1)'        // Returns "FFFFFFFFFF"
    ],
    relatedFunctions: ['HEX2DEC', 'DEC2BIN']
  },
  
  DEC2OCT: {
    syntax: 'DEC2OCT(decimal_number, [significant_digits])',
    description: 'Converts decimal to octal',
    params: {
      decimal_number: 'Decimal number',
      significant_digits: 'Optional. Minimum digits'
    },
    examples: [
      'DEC2OCT(64)',   // Returns "100"
      'DEC2OCT(8, 3)'  // Returns "010"
    ],
    relatedFunctions: ['OCT2DEC', 'DEC2BIN']
  },
  
  HEX2BIN: {
    syntax: 'HEX2BIN(hex_number, [significant_digits])',
    description: 'Converts hexadecimal to binary',
    params: {
      hex_number: 'Hexadecimal number',
      significant_digits: 'Optional. Minimum digits'
    },
    examples: [
      'HEX2BIN("F")',     // Returns "1111"
      'HEX2BIN("A", 8)'   // Returns "00001010"
    ],
    relatedFunctions: ['BIN2HEX', 'HEX2DEC']
  },
  
  HEX2DEC: {
    syntax: 'HEX2DEC(hex_number)',
    description: 'Converts hexadecimal to decimal',
    params: {
      hex_number: 'Hexadecimal number (up to 10 chars)'
    },
    examples: [
      'HEX2DEC("FF")',     // Returns 255
      'HEX2DEC("FFFF")',   // Returns 65535
      'HEX2DEC("A")'       // Returns 10
    ],
    relatedFunctions: ['DEC2HEX', 'HEX2BIN']
  },
  
  HEX2OCT: {
    syntax: 'HEX2OCT(hex_number, [significant_digits])',
    description: 'Converts hexadecimal to octal',
    params: {
      hex_number: 'Hexadecimal number',
      significant_digits: 'Optional. Minimum digits'
    },
    examples: [
      'HEX2OCT("F")',    // Returns "17"
      'HEX2OCT("3F")'    // Returns "77"
    ],
    relatedFunctions: ['OCT2HEX', 'HEX2DEC']
  },
  
  OCT2BIN: {
    syntax: 'OCT2BIN(octal_number, [significant_digits])',
    description: 'Converts octal to binary',
    params: {
      octal_number: 'Octal number',
      significant_digits: 'Optional. Minimum digits'
    },
    examples: [
      'OCT2BIN("7")',      // Returns "111"
      'OCT2BIN("12", 8)'   // Returns "00001010"
    ],
    relatedFunctions: ['BIN2OCT', 'OCT2DEC']
  },
  
  OCT2DEC: {
    syntax: 'OCT2DEC(octal_number)',
    description: 'Converts octal to decimal',
    params: {
      octal_number: 'Octal number (up to 10 digits)'
    },
    examples: [
      'OCT2DEC("77")',   // Returns 63
      'OCT2DEC("100")'   // Returns 64
    ],
    relatedFunctions: ['DEC2OCT', 'OCT2BIN']
  },
  
  OCT2HEX: {
    syntax: 'OCT2HEX(octal_number, [significant_digits])',
    description: 'Converts octal to hexadecimal',
    params: {
      octal_number: 'Octal number',
      significant_digits: 'Optional. Minimum digits'
    },
    examples: [
      'OCT2HEX("100")',  // Returns "40"
      'OCT2HEX("77")'    // Returns "3F"
    ],
    relatedFunctions: ['HEX2OCT', 'OCT2DEC']
  },
  
  // ========== UNIT CONVERSION ==========
  
  CONVERT: {
    syntax: 'CONVERT(value, from_unit, to_unit)',
    description: 'Converts between measurement units',
    params: {
      value: 'Number to convert',
      from_unit: 'Source unit',
      to_unit: 'Target unit'
    },
    examples: [
      // Distance
      'CONVERT(1, "mi", "km")',      // Miles to kilometers
      'CONVERT(100, "m", "ft")',     // Meters to feet
      'CONVERT(1, "in", "cm")',      // Inches to centimeters
      
      // Weight
      'CONVERT(1, "kg", "lbm")',     // Kilograms to pounds
      'CONVERT(1, "oz", "g")',       // Ounces to grams
      
      // Temperature
      'CONVERT(100, "C", "F")',      // Celsius to Fahrenheit
      'CONVERT(0, "C", "K")',        // Celsius to Kelvin
      
      // Volume
      'CONVERT(1, "gal", "l")',      // Gallons to liters
      'CONVERT(1, "cup", "ml")',     // Cups to milliliters
      
      // Time
      'CONVERT(1, "hr", "min")',     // Hours to minutes
      'CONVERT(1, "day", "sec")',    // Days to seconds
      
      // Area
      'CONVERT(1, "ha", "m^2")',     // Hectares to sq meters
      'CONVERT(1, "acre", "ft^2")',  // Acres to sq feet
      
      // Energy
      'CONVERT(1, "cal", "J")',      // Calories to joules
      'CONVERT(1, "BTU", "kWh")',    // BTU to kilowatt-hours
      
      // Speed
      'CONVERT(100, "km/h", "mph")', // KPH to MPH
      'CONVERT(1, "m/s", "mph")'     // M/s to MPH
    ],
    units: {
      distance: ['m', 'mi', 'Nmi', 'in', 'ft', 'yd', 'ang', 'Pica', 'km', 'cm', 'mm'],
      weight: ['g', 'kg', 'lbm', 'oz', 'ton', 'u', 'grain', 'cwt', 'stone'],
      temperature: ['C', 'F', 'K', 'Rank', 'Reau'],
      volume: ['l', 'gal', 'qt', 'pt', 'cup', 'tsp', 'tbs', 'ml', 'oz'],
      area: ['m^2', 'ha', 'acre', 'ft^2', 'in^2', 'km^2', 'mi^2'],
      time: ['yr', 'day', 'hr', 'min', 'sec'],
      speed: ['m/s', 'km/h', 'mph', 'kn'],
      energy: ['J', 'e', 'c', 'cal', 'eV', 'HPh', 'Wh', 'BTU'],
      power: ['W', 'HP', 'PS'],
      pressure: ['Pa', 'atm', 'mmHg', 'psi', 'Torr'],
      force: ['N', 'dyn', 'lbf'],
      magnetism: ['T', 'ga'],
      information: ['bit', 'byte']
    },
    tips: [
      'Units are case-sensitive',
      'Use prefixes: k (kilo), M (mega), G (giga), m (milli), u (micro)',
      'For squared/cubed: "m^2", "m^3"'
    ],
    relatedFunctions: []
  },
  
  // ========== COMPLEX NUMBERS ==========
  
  COMPLEX: {
    syntax: 'COMPLEX(real, imaginary, [suffix])',
    description: 'Creates complex number from real and imaginary parts',
    params: {
      real: 'Real part',
      imaginary: 'Imaginary part',
      suffix: 'Optional. "i" or "j" (default: "i")'
    },
    examples: [
      'COMPLEX(3, 4)',       // Returns "3+4i"
      'COMPLEX(3, -4)',      // Returns "3-4i"
      'COMPLEX(0, 1, "j")'   // Returns "j"
    ],
    relatedFunctions: ['IMAGINARY', 'IMREAL', 'IMABS']
  },
  
  IMAGINARY: {
    syntax: 'IMAGINARY(complex_number)',
    description: 'Returns imaginary coefficient of complex number',
    params: {
      complex_number: 'Complex number string'
    },
    examples: [
      'IMAGINARY("3+4i")',  // Returns 4
      'IMAGINARY("5i")'    // Returns 5
    ],
    relatedFunctions: ['IMREAL', 'COMPLEX']
  },
  
  IMREAL: {
    syntax: 'IMREAL(complex_number)',
    description: 'Returns real coefficient of complex number',
    params: {
      complex_number: 'Complex number string'
    },
    examples: [
      'IMREAL("3+4i")',  // Returns 3
      'IMREAL("5i")'    // Returns 0
    ],
    relatedFunctions: ['IMAGINARY', 'COMPLEX']
  },
  
  IMABS: {
    syntax: 'IMABS(complex_number)',
    description: 'Returns absolute value (modulus) of complex number',
    params: {
      complex_number: 'Complex number string'
    },
    examples: [
      'IMABS("3+4i")'  // Returns 5 (sqrt(3²+4²))
    ],
    relatedFunctions: ['IMARGUMENT', 'COMPLEX']
  },
  
  IMARGUMENT: {
    syntax: 'IMARGUMENT(complex_number)',
    description: 'Returns argument (angle) of complex number in radians',
    params: {
      complex_number: 'Complex number string'
    },
    examples: [
      'IMARGUMENT("1+i")',  // Returns PI/4 (~0.785)
      'DEGREES(IMARGUMENT("1+i"))'  // Returns 45
    ],
    relatedFunctions: ['IMABS', 'COMPLEX']
  },
  
  IMCONJUGATE: {
    syntax: 'IMCONJUGATE(complex_number)',
    description: 'Returns complex conjugate',
    params: {
      complex_number: 'Complex number string'
    },
    examples: [
      'IMCONJUGATE("3+4i")'  // Returns "3-4i"
    ],
    relatedFunctions: ['COMPLEX', 'IMAGINARY']
  },
  
  IMSUM: {
    syntax: 'IMSUM(complex1, [complex2], ...)',
    description: 'Sums complex numbers',
    params: {
      complex1: 'First complex number',
      complex2: 'Additional complex numbers'
    },
    examples: [
      'IMSUM("3+4i", "1+2i")'  // Returns "4+6i"
    ],
    relatedFunctions: ['IMSUB', 'IMPRODUCT']
  },
  
  IMSUB: {
    syntax: 'IMSUB(complex1, complex2)',
    description: 'Subtracts complex numbers',
    params: {
      complex1: 'First complex number',
      complex2: 'Number to subtract'
    },
    examples: [
      'IMSUB("5+3i", "2+1i")'  // Returns "3+2i"
    ],
    relatedFunctions: ['IMSUM', 'IMPRODUCT']
  },
  
  IMPRODUCT: {
    syntax: 'IMPRODUCT(complex1, [complex2], ...)',
    description: 'Multiplies complex numbers',
    params: {
      complex1: 'First complex number',
      complex2: 'Additional complex numbers'
    },
    examples: [
      'IMPRODUCT("1+2i", "3+4i")'  // Returns "-5+10i"
    ],
    relatedFunctions: ['IMDIV', 'IMSUM']
  },
  
  IMDIV: {
    syntax: 'IMDIV(complex1, complex2)',
    description: 'Divides complex numbers',
    params: {
      complex1: 'Dividend',
      complex2: 'Divisor'
    },
    examples: [
      'IMDIV("4+2i", "1+i")'  // Returns "3-i"
    ],
    relatedFunctions: ['IMPRODUCT']
  },
  
  IMPOWER: {
    syntax: 'IMPOWER(complex_number, power)',
    description: 'Raises complex number to a power',
    params: {
      complex_number: 'Complex number string',
      power: 'Power to raise to'
    },
    examples: [
      'IMPOWER("2+3i", 2)'
    ],
    relatedFunctions: ['IMSQRT', 'IMEXP']
  },
  
  IMSQRT: {
    syntax: 'IMSQRT(complex_number)',
    description: 'Square root of complex number',
    params: {
      complex_number: 'Complex number string'
    },
    examples: [
      'IMSQRT("-1")'  // Returns "i"
    ],
    relatedFunctions: ['IMPOWER']
  },
  
  IMEXP: {
    syntax: 'IMEXP(complex_number)',
    description: 'Exponential of complex number (e^z)',
    relatedFunctions: ['IMLN', 'IMPOWER']
  },
  
  IMLN: {
    syntax: 'IMLN(complex_number)',
    description: 'Natural logarithm of complex number',
    relatedFunctions: ['IMEXP', 'IMLOG10', 'IMLOG2']
  },
  
  IMLOG10: {
    syntax: 'IMLOG10(complex_number)',
    description: 'Base-10 logarithm of complex number',
    relatedFunctions: ['IMLN', 'IMLOG2']
  },
  
  IMLOG2: {
    syntax: 'IMLOG2(complex_number)',
    description: 'Base-2 logarithm of complex number',
    relatedFunctions: ['IMLN', 'IMLOG10']
  },
  
  // Complex trig functions
  IMSIN: { syntax: 'IMSIN(complex_number)', description: 'Sine of complex number' },
  IMCOS: { syntax: 'IMCOS(complex_number)', description: 'Cosine of complex number' },
  IMTAN: { syntax: 'IMTAN(complex_number)', description: 'Tangent of complex number' },
  IMSINH: { syntax: 'IMSINH(complex_number)', description: 'Hyperbolic sine of complex number' },
  IMCOSH: { syntax: 'IMCOSH(complex_number)', description: 'Hyperbolic cosine of complex number' },
  IMSEC: { syntax: 'IMSEC(complex_number)', description: 'Secant of complex number' },
  IMCSC: { syntax: 'IMCSC(complex_number)', description: 'Cosecant of complex number' },
  IMCOT: { syntax: 'IMCOT(complex_number)', description: 'Cotangent of complex number' },
  IMSECH: { syntax: 'IMSECH(complex_number)', description: 'Hyperbolic secant' },
  IMCSCH: { syntax: 'IMCSCH(complex_number)', description: 'Hyperbolic cosecant' },
  
  // ========== BESSEL FUNCTIONS ==========
  
  BESSELI: {
    syntax: 'BESSELI(x, n)',
    description: 'Modified Bessel function In(x)',
    params: {
      x: 'Value to evaluate',
      n: 'Order of Bessel function'
    },
    examples: [
      'BESSELI(1.5, 1)'
    ],
    relatedFunctions: ['BESSELJ', 'BESSELK', 'BESSELY']
  },
  
  BESSELJ: {
    syntax: 'BESSELJ(x, n)',
    description: 'Bessel function Jn(x)',
    params: {
      x: 'Value to evaluate',
      n: 'Order of Bessel function'
    },
    relatedFunctions: ['BESSELI', 'BESSELK', 'BESSELY']
  },
  
  BESSELK: {
    syntax: 'BESSELK(x, n)',
    description: 'Modified Bessel function Kn(x)',
    relatedFunctions: ['BESSELI', 'BESSELJ', 'BESSELY']
  },
  
  BESSELY: {
    syntax: 'BESSELY(x, n)',
    description: 'Bessel function Yn(x)',
    relatedFunctions: ['BESSELI', 'BESSELJ', 'BESSELK']
  },
  
  // ========== ERROR FUNCTION ==========
  
  ERF: {
    syntax: 'ERF(lower_bound, [upper_bound])',
    description: 'Error function integrated between bounds',
    params: {
      lower_bound: 'Lower bound (or only bound)',
      upper_bound: 'Optional. Upper bound'
    },
    examples: [
      'ERF(1)',      // erf(1) from 0 to 1
      'ERF(0, 1.5)'  // erf from 0 to 1.5
    ],
    relatedFunctions: ['ERFC']
  },
  
  ERFC: {
    syntax: 'ERFC(z)',
    description: 'Complementary error function',
    params: {
      z: 'Value'
    },
    examples: [
      'ERFC(1)'  // Returns 1 - ERF(1)
    ],
    relatedFunctions: ['ERF']
  },
  
  // ========== BIT OPERATIONS ==========
  
  BITAND: {
    syntax: 'BITAND(value1, value2)',
    description: 'Bitwise AND of two numbers',
    params: {
      value1: 'First number',
      value2: 'Second number'
    },
    examples: [
      'BITAND(12, 10)',  // 1100 AND 1010 = 1000 = 8
      'BITAND(255, 15)'  // Returns 15
    ],
    relatedFunctions: ['BITOR', 'BITXOR', 'BITNOT']
  },
  
  BITOR: {
    syntax: 'BITOR(value1, value2)',
    description: 'Bitwise OR of two numbers',
    params: {
      value1: 'First number',
      value2: 'Second number'
    },
    examples: [
      'BITOR(12, 10)'  // 1100 OR 1010 = 1110 = 14
    ],
    relatedFunctions: ['BITAND', 'BITXOR']
  },
  
  BITXOR: {
    syntax: 'BITXOR(value1, value2)',
    description: 'Bitwise XOR of two numbers',
    params: {
      value1: 'First number',
      value2: 'Second number'
    },
    examples: [
      'BITXOR(12, 10)'  // 1100 XOR 1010 = 0110 = 6
    ],
    relatedFunctions: ['BITAND', 'BITOR']
  },
  
  BITLSHIFT: {
    syntax: 'BITLSHIFT(value, shift_amount)',
    description: 'Shifts bits left',
    params: {
      value: 'Number to shift',
      shift_amount: 'Positions to shift'
    },
    examples: [
      'BITLSHIFT(4, 2)'  // 100 << 2 = 10000 = 16
    ],
    relatedFunctions: ['BITRSHIFT']
  },
  
  BITRSHIFT: {
    syntax: 'BITRSHIFT(value, shift_amount)',
    description: 'Shifts bits right',
    params: {
      value: 'Number to shift',
      shift_amount: 'Positions to shift'
    },
    examples: [
      'BITRSHIFT(16, 2)'  // 10000 >> 2 = 100 = 4
    ],
    relatedFunctions: ['BITLSHIFT']
  }
};

/**
 * Get engineering function info
 */
function getEngineeringFunction(name) {
  return ENGINEERING_FUNCTIONS[name.toUpperCase()] || null;
}

/**
 * Get all engineering function names
 */
function getEngineeringFunctionNames() {
  return Object.keys(ENGINEERING_FUNCTIONS);
}

Logger.log('📦 FormulaCatalog_Engineering loaded - ' + Object.keys(ENGINEERING_FUNCTIONS).length + ' functions');
