/**
 * @file FormulaCatalog_Financial.gs
 * @description Financial Functions for Google Sheets
 * 
 * COMPLETE COVERAGE of financial functions for loans, investments,
 * depreciation, and other financial calculations.
 */

var FINANCIAL_FUNCTIONS = {
  
  // ========== LOAN & PAYMENT CALCULATIONS ==========
  
  PMT: {
    syntax: 'PMT(rate, number_of_periods, present_value, [future_value], [end_or_beginning])',
    description: 'Calculates periodic payment for a loan',
    params: {
      rate: 'Interest rate per period',
      number_of_periods: 'Total number of payments',
      present_value: 'Loan amount (positive)',
      future_value: 'Optional. Remaining balance (default: 0)',
      end_or_beginning: 'Optional. 0=end of period (default), 1=beginning'
    },
    examples: [
      'PMT(0.05/12, 360, 200000)',  // Monthly payment on $200k @ 5% for 30 years
      'PMT(5%/12, 12*30, 200000)',  // Same, different notation
      'PMT(A1/12, B1*12, C1)'  // Using cell references
    ],
    tips: [
      'Rate must match period (monthly rate for monthly payments)',
      'Result is negative (cash outflow)',
      'Use ABS() to get positive number'
    ],
    relatedFunctions: ['PPMT', 'IPMT', 'FV', 'PV', 'NPER', 'RATE']
  },
  
  PPMT: {
    syntax: 'PPMT(rate, period, number_of_periods, present_value, [future_value], [end_or_beginning])',
    description: 'Calculates principal portion of a specific payment',
    params: {
      rate: 'Interest rate per period',
      period: 'Specific payment period (1 to n)',
      number_of_periods: 'Total number of payments',
      present_value: 'Loan amount'
    },
    examples: [
      'PPMT(0.05/12, 1, 360, 200000)',  // Principal in 1st payment
      'PPMT(0.05/12, 360, 360, 200000)'  // Principal in last payment
    ],
    tips: [
      'Principal portion increases over time',
      'PPMT + IPMT = PMT for each period'
    ],
    relatedFunctions: ['IPMT', 'PMT', 'CUMIPMT', 'CUMPRINC']
  },
  
  IPMT: {
    syntax: 'IPMT(rate, period, number_of_periods, present_value, [future_value], [end_or_beginning])',
    description: 'Calculates interest portion of a specific payment',
    params: {
      rate: 'Interest rate per period',
      period: 'Specific payment period',
      number_of_periods: 'Total number of payments',
      present_value: 'Loan amount'
    },
    examples: [
      'IPMT(0.05/12, 1, 360, 200000)',  // Interest in 1st payment
      'IPMT(0.05/12, 360, 360, 200000)'  // Interest in last payment (almost zero)
    ],
    tips: [
      'Interest portion decreases over time',
      'Useful for building amortization schedules'
    ],
    relatedFunctions: ['PPMT', 'PMT', 'CUMIPMT']
  },
  
  CUMIPMT: {
    syntax: 'CUMIPMT(rate, number_of_periods, present_value, start_period, end_period, end_or_beginning)',
    description: 'Calculates cumulative interest paid between two periods',
    params: {
      rate: 'Interest rate per period',
      number_of_periods: 'Total periods',
      present_value: 'Loan amount',
      start_period: 'First period in calculation',
      end_period: 'Last period in calculation',
      end_or_beginning: '0=end, 1=beginning of period'
    },
    examples: [
      'CUMIPMT(0.05/12, 360, 200000, 1, 12, 0)',  // Interest in first year
      'CUMIPMT(0.05/12, 360, 200000, 1, 360, 0)'  // Total interest over loan life
    ],
    tips: [
      'Great for calculating yearly interest',
      'Returns negative (outflow)'
    ],
    relatedFunctions: ['CUMPRINC', 'IPMT', 'PMT']
  },
  
  CUMPRINC: {
    syntax: 'CUMPRINC(rate, number_of_periods, present_value, start_period, end_period, end_or_beginning)',
    description: 'Calculates cumulative principal paid between two periods',
    params: {
      rate: 'Interest rate per period',
      number_of_periods: 'Total periods',
      present_value: 'Loan amount',
      start_period: 'First period',
      end_period: 'Last period',
      end_or_beginning: '0=end, 1=beginning'
    },
    examples: [
      'CUMPRINC(0.05/12, 360, 200000, 1, 12, 0)',  // Principal paid in first year
      'CUMPRINC(0.05/12, 360, 200000, 1, 360, 0)'  // Total principal (should equal loan)
    ],
    relatedFunctions: ['CUMIPMT', 'PPMT', 'PMT']
  },
  
  // ========== TIME VALUE OF MONEY ==========
  
  PV: {
    syntax: 'PV(rate, number_of_periods, payment_amount, [future_value], [end_or_beginning])',
    description: 'Calculates present value of an investment',
    params: {
      rate: 'Interest rate per period',
      number_of_periods: 'Number of periods',
      payment_amount: 'Payment per period',
      future_value: 'Optional. Future value',
      end_or_beginning: 'Optional. 0=end, 1=beginning'
    },
    examples: [
      'PV(0.05/12, 360, -1000)',  // Loan amount from $1000/month payments
      'PV(0.08, 5, 0, 10000)',    // Present value of $10k in 5 years @ 8%
      'PV(0.1, 10, -1000, -10000)'  // With both payments and lump sum
    ],
    tips: [
      'Negative payments = cash outflows',
      'Useful for comparing investment options'
    ],
    relatedFunctions: ['FV', 'PMT', 'NPV']
  },
  
  FV: {
    syntax: 'FV(rate, number_of_periods, payment_amount, [present_value], [end_or_beginning])',
    description: 'Calculates future value of an investment',
    params: {
      rate: 'Interest rate per period',
      number_of_periods: 'Number of periods',
      payment_amount: 'Payment per period',
      present_value: 'Optional. Initial investment',
      end_or_beginning: 'Optional. 0=end, 1=beginning'
    },
    examples: [
      'FV(0.07/12, 360, -500)',        // 30 years of $500/month @ 7%
      'FV(0.07/12, 360, -500, -10000)', // Same with $10k initial
      'FV(0.1, 5, 0, -1000)'           // $1000 invested for 5 years @ 10%
    ],
    tips: [
      'Negative inputs = cash outflows',
      'Great for retirement planning'
    ],
    relatedFunctions: ['PV', 'PMT', 'FVSCHEDULE']
  },
  
  FVSCHEDULE: {
    syntax: 'FVSCHEDULE(principal, schedule)',
    description: 'Future value with variable interest rate schedule',
    params: {
      principal: 'Initial investment',
      schedule: 'Array of interest rates'
    },
    examples: [
      'FVSCHEDULE(1000, {0.05, 0.06, 0.07})',  // Variable rates over 3 years
      'FVSCHEDULE(A1, B1:B10)'  // Rates in cells
    ],
    tips: [
      'Each rate applies to one period',
      'Useful for variable rate investments'
    ],
    relatedFunctions: ['FV', 'IRR']
  },
  
  NPER: {
    syntax: 'NPER(rate, payment_amount, present_value, [future_value], [end_or_beginning])',
    description: 'Calculates number of periods for investment/loan',
    params: {
      rate: 'Interest rate per period',
      payment_amount: 'Payment per period',
      present_value: 'Present value/loan amount',
      future_value: 'Optional. Target future value',
      end_or_beginning: 'Optional. 0=end, 1=beginning'
    },
    examples: [
      'NPER(0.05/12, -1000, 200000)',  // Months to pay off $200k @ 5%
      'NPER(0.07/12, -500, 0, 100000)'  // Months to save $100k
    ],
    tips: [
      'Returns number of periods (may be fractional)',
      'Useful for "how long until..." questions'
    ],
    relatedFunctions: ['RATE', 'PMT', 'PV', 'FV']
  },
  
  RATE: {
    syntax: 'RATE(number_of_periods, payment_amount, present_value, [future_value], [end_or_beginning], [guess])',
    description: 'Calculates interest rate per period',
    params: {
      number_of_periods: 'Total number of periods',
      payment_amount: 'Payment per period',
      present_value: 'Present value',
      future_value: 'Optional. Future value',
      end_or_beginning: 'Optional. 0=end, 1=beginning',
      guess: 'Optional. Initial guess for iteration'
    },
    examples: [
      'RATE(360, -1074, 200000)*12',  // Annual rate from monthly payment
      'RATE(5, 0, -1000, 1500)'       // Rate to grow $1000 to $1500 in 5 years
    ],
    tips: [
      'Returns rate per period - multiply by periods/year for annual',
      'May need guess for complex scenarios'
    ],
    relatedFunctions: ['NPER', 'PMT', 'IRR']
  },
  
  // ========== NET PRESENT VALUE & IRR ==========
  
  NPV: {
    syntax: 'NPV(discount_rate, cashflow1, [cashflow2], ...)',
    description: 'Net present value of cash flows (starting at period 1)',
    params: {
      discount_rate: 'Discount rate per period',
      cashflow1: 'Cash flow in period 1',
      cashflow2: 'Additional cash flows'
    },
    examples: [
      'NPV(0.1, -10000, 3000, 4000, 5000, 6000)',  // Project cash flows
      'NPV(0.08, B2:B10)',  // Cash flows in range
      '-A1 + NPV(0.1, B2:B10)'  // Include initial investment (period 0)
    ],
    tips: [
      'Assumes first cash flow is END of period 1',
      'For period 0 investment, add separately: -Investment + NPV(...)',
      'Positive NPV = good investment at that rate'
    ],
    relatedFunctions: ['IRR', 'XNPV', 'PV']
  },
  
  XNPV: {
    syntax: 'XNPV(discount_rate, cashflows, dates)',
    description: 'NPV for irregular cash flow dates',
    params: {
      discount_rate: 'Annual discount rate',
      cashflows: 'Array of cash flow amounts',
      dates: 'Array of dates for each cash flow'
    },
    examples: [
      'XNPV(0.1, B2:B10, A2:A10)',
      'XNPV(10%, {-10000, 3000, 4000, 5000}, {DATE(2024,1,1), DATE(2024,6,15), DATE(2025,1,1), DATE(2025,7,1)})'
    ],
    tips: [
      'Dates can be irregular',
      'First cash flow is usually negative (investment)',
      'Uses actual/365 day convention'
    ],
    relatedFunctions: ['NPV', 'XIRR']
  },
  
  IRR: {
    syntax: 'IRR(cashflows, [guess])',
    description: 'Internal rate of return for regular cash flows',
    params: {
      cashflows: 'Array of cash flows',
      guess: 'Optional. Initial guess (default: 0.1)'
    },
    examples: [
      'IRR({-10000, 3000, 4000, 5000, 6000})',
      'IRR(B1:B10)',
      'IRR(B1:B10, 0.2)'  // With initial guess
    ],
    tips: [
      'Must have at least one positive and one negative',
      'Higher IRR = better investment',
      'May need guess if doesn\'t converge'
    ],
    relatedFunctions: ['NPV', 'XIRR', 'MIRR']
  },
  
  XIRR: {
    syntax: 'XIRR(cashflows, dates, [guess])',
    description: 'IRR for irregular cash flow dates',
    params: {
      cashflows: 'Array of cash flow amounts',
      dates: 'Array of dates',
      guess: 'Optional. Initial guess'
    },
    examples: [
      'XIRR(B2:B10, A2:A10)',
      'XIRR({-10000, 3000, 4000, 5000}, {DATE(2024,1,1), DATE(2024,6,15), DATE(2025,1,1), DATE(2025,7,1)})'
    ],
    tips: [
      'Returns annualized rate',
      'Handles irregular timing'
    ],
    relatedFunctions: ['IRR', 'XNPV']
  },
  
  MIRR: {
    syntax: 'MIRR(cashflows, finance_rate, reinvest_rate)',
    description: 'Modified IRR with specified financing/reinvestment rates',
    params: {
      cashflows: 'Array of cash flows',
      finance_rate: 'Rate paid on negative cash flows',
      reinvest_rate: 'Rate earned on positive cash flows'
    },
    examples: [
      'MIRR({-10000, 3000, 4000, 5000, 6000}, 0.05, 0.08)',
      'MIRR(B1:B10, 0.06, 0.1)'
    ],
    tips: [
      'More realistic than IRR',
      'Addresses IRR multiple solution problem'
    ],
    relatedFunctions: ['IRR', 'NPV']
  },
  
  // ========== DEPRECIATION ==========
  
  SLN: {
    syntax: 'SLN(cost, salvage, life)',
    description: 'Straight-line depreciation per period',
    params: {
      cost: 'Initial cost of asset',
      salvage: 'Salvage value at end of life',
      life: 'Number of periods'
    },
    examples: [
      'SLN(100000, 10000, 10)',  // $100k asset, $10k salvage, 10 years
      'SLN(A1, B1, C1)'
    ],
    tips: [
      'Same depreciation each period',
      'Simplest depreciation method'
    ],
    relatedFunctions: ['DB', 'DDB', 'SYD', 'VDB']
  },
  
  DB: {
    syntax: 'DB(cost, salvage, life, period, [month])',
    description: 'Fixed declining balance depreciation',
    params: {
      cost: 'Initial cost',
      salvage: 'Salvage value',
      life: 'Life in periods',
      period: 'Period to calculate',
      month: 'Optional. Months in first year (default: 12)'
    },
    examples: [
      'DB(100000, 10000, 10, 1)',  // Year 1 depreciation
      'DB(100000, 10000, 10, 5)'   // Year 5 depreciation
    ],
    tips: [
      'Accelerated depreciation',
      'Higher depreciation in early years'
    ],
    relatedFunctions: ['DDB', 'SLN', 'VDB']
  },
  
  DDB: {
    syntax: 'DDB(cost, salvage, life, period, [factor])',
    description: 'Double declining balance depreciation',
    params: {
      cost: 'Initial cost',
      salvage: 'Salvage value',
      life: 'Life in periods',
      period: 'Period to calculate',
      factor: 'Optional. Depreciation factor (default: 2)'
    },
    examples: [
      'DDB(100000, 10000, 10, 1)',      // 200% declining balance
      'DDB(100000, 10000, 10, 1, 1.5)'  // 150% declining balance
    ],
    tips: [
      'Factor of 2 = 200% declining balance',
      'Never depreciates below salvage value'
    ],
    relatedFunctions: ['DB', 'SLN', 'VDB']
  },
  
  SYD: {
    syntax: 'SYD(cost, salvage, life, period)',
    description: 'Sum of years digits depreciation',
    params: {
      cost: 'Initial cost',
      salvage: 'Salvage value',
      life: 'Life in periods',
      period: 'Period to calculate'
    },
    examples: [
      'SYD(100000, 10000, 10, 1)',  // Year 1
      'SYD(100000, 10000, 10, 10)' // Year 10
    ],
    tips: [
      'Accelerated method',
      'Total depreciation = cost - salvage'
    ],
    relatedFunctions: ['DDB', 'DB', 'SLN']
  },
  
  VDB: {
    syntax: 'VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])',
    description: 'Variable declining balance depreciation',
    params: {
      cost: 'Initial cost',
      salvage: 'Salvage value',
      life: 'Life in periods',
      start_period: 'Start period',
      end_period: 'End period',
      factor: 'Optional. Declining balance factor',
      no_switch: 'Optional. FALSE to switch to SLN when advantageous'
    },
    examples: [
      'VDB(100000, 10000, 10, 0, 1)',     // First year
      'VDB(100000, 10000, 10, 4, 5)',     // Fifth year
      'VDB(100000, 10000, 10, 0, 0.5)'    // First half year
    ],
    tips: [
      'Flexible - can handle partial periods',
      'Can switch to straight-line automatically'
    ],
    relatedFunctions: ['DDB', 'DB', 'SLN']
  },
  
  AMORLINC: {
    syntax: 'AMORLINC(cost, purchase_date, first_period_end, salvage, period, rate, [basis])',
    description: 'Linear depreciation for French accounting',
    relatedFunctions: ['AMORDEGRC', 'SLN']
  },
  
  AMORDEGRC: {
    syntax: 'AMORDEGRC(cost, purchase_date, first_period_end, salvage, period, rate, [basis])',
    description: 'Declining depreciation for French accounting',
    relatedFunctions: ['AMORLINC', 'DDB']
  },
  
  // ========== BOND CALCULATIONS ==========
  
  PRICE: {
    syntax: 'PRICE(settlement, maturity, rate, yield, redemption, frequency, [day_count_convention])',
    description: 'Price per $100 face value of a bond',
    params: {
      settlement: 'Settlement date',
      maturity: 'Maturity date',
      rate: 'Annual coupon rate',
      yield: 'Annual yield',
      redemption: 'Redemption value per $100 face',
      frequency: 'Payments per year (1, 2, or 4)',
      day_count_convention: 'Optional. Day count basis'
    },
    examples: [
      'PRICE(DATE(2024,1,15), DATE(2034,1,15), 0.05, 0.06, 100, 2)'
    ],
    relatedFunctions: ['YIELD', 'DURATION', 'MDURATION']
  },
  
  YIELD: {
    syntax: 'YIELD(settlement, maturity, rate, price, redemption, frequency, [day_count_convention])',
    description: 'Yield of a security',
    params: {
      settlement: 'Settlement date',
      maturity: 'Maturity date',
      rate: 'Annual coupon rate',
      price: 'Price per $100 face',
      redemption: 'Redemption value',
      frequency: 'Payments per year'
    },
    examples: [
      'YIELD(DATE(2024,1,15), DATE(2034,1,15), 0.05, 98, 100, 2)'
    ],
    relatedFunctions: ['PRICE', 'YIELDMAT', 'YIELDDISC']
  },
  
  DURATION: {
    syntax: 'DURATION(settlement, maturity, coupon, yield, frequency, [day_count_convention])',
    description: 'Macaulay duration of a bond',
    params: {
      settlement: 'Settlement date',
      maturity: 'Maturity date',
      coupon: 'Annual coupon rate',
      yield: 'Annual yield',
      frequency: 'Payments per year'
    },
    examples: [
      'DURATION(DATE(2024,1,15), DATE(2034,1,15), 0.05, 0.06, 2)'
    ],
    tips: [
      'Measures price sensitivity to yield changes',
      'Expressed in years'
    ],
    relatedFunctions: ['MDURATION', 'PRICE']
  },
  
  MDURATION: {
    syntax: 'MDURATION(settlement, maturity, coupon, yield, frequency, [day_count_convention])',
    description: 'Modified duration of a bond',
    tips: [
      'Modified = Macaulay / (1 + yield/frequency)',
      'Better for estimating price change'
    ],
    relatedFunctions: ['DURATION', 'PRICE']
  },
  
  ACCRINT: {
    syntax: 'ACCRINT(issue, first_interest, settlement, rate, par, frequency, [day_count_convention], [calc_method])',
    description: 'Accrued interest for a security with periodic payments',
    relatedFunctions: ['ACCRINTM', 'PRICE']
  },
  
  ACCRINTM: {
    syntax: 'ACCRINTM(issue, settlement, rate, par, [day_count_convention])',
    description: 'Accrued interest for a security paid at maturity',
    relatedFunctions: ['ACCRINT']
  },
  
  // ========== INTEREST RATE CONVERSION ==========
  
  EFFECT: {
    syntax: 'EFFECT(nominal_rate, npery)',
    description: 'Converts nominal to effective annual rate',
    params: {
      nominal_rate: 'Nominal annual interest rate',
      npery: 'Number of compounding periods per year'
    },
    examples: [
      'EFFECT(0.10, 12)',  // 10% nominal monthly compounding
      'EFFECT(0.10, 365)' // 10% daily compounding
    ],
    tips: [
      'Effective rate > nominal rate (for npery > 1)',
      'APY calculation'
    ],
    relatedFunctions: ['NOMINAL']
  },
  
  NOMINAL: {
    syntax: 'NOMINAL(effect_rate, npery)',
    description: 'Converts effective to nominal annual rate',
    params: {
      effect_rate: 'Effective annual interest rate',
      npery: 'Number of compounding periods per year'
    },
    examples: [
      'NOMINAL(0.1047, 12)'  // ~10% nominal from 10.47% effective
    ],
    relatedFunctions: ['EFFECT']
  },
  
  // ========== OTHER FINANCIAL ==========
  
  INTRATE: {
    syntax: 'INTRATE(settlement, maturity, investment, redemption, [day_count_convention])',
    description: 'Interest rate for fully invested security',
    relatedFunctions: ['RECEIVED', 'DISC']
  },
  
  RECEIVED: {
    syntax: 'RECEIVED(settlement, maturity, investment, discount, [day_count_convention])',
    description: 'Amount received at maturity for discounted security',
    relatedFunctions: ['INTRATE', 'DISC']
  },
  
  DISC: {
    syntax: 'DISC(settlement, maturity, price, redemption, [day_count_convention])',
    description: 'Discount rate for a security',
    relatedFunctions: ['PRICEDISC', 'YIELDDISC']
  },
  
  PRICEDISC: {
    syntax: 'PRICEDISC(settlement, maturity, discount, redemption, [day_count_convention])',
    description: 'Price of discounted security',
    relatedFunctions: ['DISC', 'YIELDDISC']
  },
  
  YIELDDISC: {
    syntax: 'YIELDDISC(settlement, maturity, price, redemption, [day_count_convention])',
    description: 'Annual yield of discounted security',
    relatedFunctions: ['DISC', 'PRICEDISC']
  },
  
  TBILLPRICE: {
    syntax: 'TBILLPRICE(settlement, maturity, discount)',
    description: 'Price of Treasury bill',
    relatedFunctions: ['TBILLYIELD', 'TBILLEQ']
  },
  
  TBILLYIELD: {
    syntax: 'TBILLYIELD(settlement, maturity, price)',
    description: 'Yield of Treasury bill',
    relatedFunctions: ['TBILLPRICE', 'TBILLEQ']
  },
  
  TBILLEQ: {
    syntax: 'TBILLEQ(settlement, maturity, discount)',
    description: 'Bond-equivalent yield of Treasury bill',
    relatedFunctions: ['TBILLPRICE', 'TBILLYIELD']
  },
  
  SRI: {
    syntax: 'DOLLARDE(fractional_dollar, fraction)',
    description: 'Converts fractional dollar to decimal',
    params: {
      fractional_dollar: 'Number with fractional part',
      fraction: 'Denominator'
    },
    examples: [
      'DOLLARDE(1.02, 16)'  // 1 and 2/16 = 1.125
    ],
    relatedFunctions: ['DOLLARFR']
  },
  
  DOLLARDE: {
    syntax: 'DOLLARDE(fractional_dollar, fraction)',
    description: 'Converts fractional dollar notation to decimal',
    params: {
      fractional_dollar: 'Number with fractional part',
      fraction: 'Denominator for fraction'
    },
    examples: [
      'DOLLARDE(1.02, 16)'  // 1 and 2/16 = 1.125
    ],
    relatedFunctions: ['DOLLARFR']
  },
  
  DOLLARFR: {
    syntax: 'DOLLARFR(decimal_dollar, fraction)',
    description: 'Converts decimal dollar to fractional notation',
    params: {
      decimal_dollar: 'Decimal number',
      fraction: 'Denominator for fraction'
    },
    examples: [
      'DOLLARFR(1.125, 16)'  // Returns 1.02 (1 and 2/16)
    ],
    relatedFunctions: ['DOLLARDE']
  },
  
  PDURATION: {
    syntax: 'PDURATION(rate, present_value, future_value)',
    description: 'Number of periods for investment to reach target',
    params: {
      rate: 'Interest rate per period',
      present_value: 'Current value',
      future_value: 'Target value'
    },
    examples: [
      'PDURATION(0.07, 1000, 2000)'  // Years to double at 7%
    ],
    tips: [
      'Rule of 72: 72/rate ≈ years to double',
      'Useful for quick growth calculations'
    ],
    relatedFunctions: ['NPER', 'FV']
  },
  
  RRI: {
    syntax: 'RRI(number_of_periods, present_value, future_value)',
    description: 'Interest rate for investment to reach target',
    params: {
      number_of_periods: 'Number of periods',
      present_value: 'Starting value',
      future_value: 'Ending value'
    },
    examples: [
      'RRI(10, 1000, 2000)'  // Rate to double in 10 years
    ],
    tips: [
      'Calculates compound annual growth rate (CAGR)',
      'Useful for performance analysis'
    ],
    relatedFunctions: ['RATE', 'IRR']
  }
};

/**
 * Get financial function info
 */
function getFinancialFunction(name) {
  return FINANCIAL_FUNCTIONS[name.toUpperCase()] || null;
}

/**
 * Get all financial function names
 */
function getFinancialFunctionNames() {
  return Object.keys(FINANCIAL_FUNCTIONS);
}

Logger.log('📦 FormulaCatalog_Financial loaded - ' + Object.keys(FINANCIAL_FUNCTIONS).length + ' functions');
