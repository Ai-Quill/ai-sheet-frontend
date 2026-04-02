/**
 * @file SheetActions_Chart.gs
 * @version 1.0.0
 * @created 2026-01-30
 * @author AISheeter Team
 * 
 * ============================================
 * SHEET ACTIONS - Chart Creation
 * ============================================
 * 
 * Creates charts from data with comprehensive configuration options.
 * Supports: bar, column, line, pie, area, scatter, combo, histogram.
 */

var SheetActions_Chart = (function() {
  
  // ============================================
  // MAIN CHART CREATION
  // ============================================
  
  /**
   * Create a chart from data with full configuration
   * @param {Object} step - { chartType, dataRange, config }
   * @return {Object} Result with chart details
   * 
   * ============================================
   * COMPREHENSIVE CONFIG OPTIONS
   * ============================================
   * 
   * COMMON OPTIONS (all chart types):
   * - title: Chart title text
   * - seriesNames: Array of legend labels for each data series
   * - colors: Array of hex colors for each series (or use defaults)
   * - legendPosition: 'right', 'bottom', 'top', 'left', 'in', 'none'
   * - width, height: Chart dimensions in pixels (default 600x400)
   * - backgroundColor: Chart background color
   * - fontName: Font family for all text
   * 
   * BAR/COLUMN CHART OPTIONS:
   * - stacked: boolean - Stack bars/columns
   * - stackedPercent: boolean - Stack as 100% (percentage)
   * - showDataLabels: boolean - Show values on bars
   * - barGroupWidth: string - Width of bar groups (e.g., '75%')
   * 
   * LINE CHART OPTIONS:
   * - curveType: 'none' (straight) or 'function' (smooth curves)
   * - lineWidth: number - Thickness of lines (default 2)
   * - pointSize: number - Size of data points (0 = no points)
   * - pointShape: 'circle', 'triangle', 'square', 'diamond', 'star', 'polygon'
   * - interpolateNulls: boolean - Connect line through null points
   * - lineDashStyle: Array - Dash pattern e.g., [4, 4] for dashed
   * 
   * AREA CHART OPTIONS:
   * - areaOpacity: number 0-1 - Transparency of fill (default 0.3)
   * - stacked: boolean - Stack areas
   * - lineWidth: number - Border line thickness
   * 
   * PIE/DONUT CHART OPTIONS:
   * - pieHole: number 0-1 - Size of hole for donut (0 = pie, 0.4 = donut)
   * - pieSliceText: 'percentage', 'value', 'label', 'none'
   * - pieSliceBorderColor: Slice border color
   * - pieStartAngle: number - Starting angle in degrees
   * - is3D: boolean - 3D pie chart
   * - sliceVisibilityThreshold: number - Hide slices below this %
   * 
   * SCATTER CHART OPTIONS:
   * - pointSize: number - Size of points (default 7)
   * - pointShape: 'circle', 'triangle', 'square', 'diamond', 'star'
   * - trendlines: boolean - Add linear trendline
   * - trendlineColor: Trendline color
   * - trendlineOpacity: number 0-1
   * - annotationColumn: string - Column letter with text labels for data points (e.g., company names)
   * 
   * AXIS OPTIONS (bar, column, line, area, scatter):
   * - xAxisTitle: Title for X-axis
   * - yAxisTitle: Title for Y-axis
   * - xAxisFormat: Number format for X-axis
   * - yAxisFormat: Number format for Y-axis
   * - xAxisMin, xAxisMax: Axis range
   * - yAxisMin, yAxisMax: Axis range
   * - gridlines: boolean - Show gridlines (default true)
   * - gridlineColor: Color of gridlines
   * - logScale: boolean - Use logarithmic scale for Y-axis
   * - secondaryAxis: Array of series indices to use secondary Y-axis
   */
  function createChart(step) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var config = step.config || {};
    
    Logger.log('[SheetActions_Chart] Config received: ' + JSON.stringify(config).substring(0, 500));
    Logger.log('[SheetActions_Chart] domainColumn=' + (config.domainColumn || 'NONE') + 
              ', dataColumns=' + JSON.stringify(config.dataColumns || 'NONE') +
              ', chartType=' + (config.chartType || 'NONE'));
    
    // SAFETY NET: If AI didn't provide domainColumn/dataColumns,
    // auto-detect them from the actual sheet data BEFORE building ranges.
    // This prevents the fallback path from using the full inputRange (A2:J16)
    // which includes text columns and produces empty charts.
    if (!config.domainColumn || !config.dataColumns || config.dataColumns.length === 0) {
      Logger.log('[SheetActions_Chart] ⚠️ Missing domainColumn/dataColumns — auto-detecting from sheet data');
      var autoNumeric = _autoDetectNumericColumns(sheet, null);
      if (autoNumeric.length > 0) {
        var lastCol = sheet.getLastColumn();
        var sampleRows = Math.min(5, sheet.getLastRow() - 1);
        for (var ac = 1; ac <= lastCol; ac++) {
          var colL = SheetActions_Utils.columnToLetter(ac);
          if (autoNumeric.indexOf(colL) === -1) {
            var vals = sheet.getRange(2, ac, sampleRows, 1).getValues();
            var hasData = vals.some(function(r) { return r[0] !== '' && r[0] !== null; });
            if (hasData) {
              config.domainColumn = colL;
              config.dataColumns = autoNumeric;
              // Generate series names from headers
              var headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
              config.seriesNames = autoNumeric.map(function(dc) {
                var idx = SheetActions_Utils.letterToColumn(dc) - 1;
                return headerValues[idx] || dc;
              });
              Logger.log('[SheetActions_Chart] Auto-detected: domain=' + colL + ', data=' + autoNumeric.join(',') + ', series=' + config.seriesNames.join(','));
              break;
            }
          }
        }
      }
    }
    
    // Build data range from AI-provided columns
    var rangeResult = _buildDataRangeFromColumns(sheet, config, step);
    Logger.log('[SheetActions_Chart] rangeResult: ' + (typeof rangeResult === 'string' ? rangeResult : JSON.stringify(rangeResult).substring(0, 300)));
    
    // Handle multi-range (non-contiguous columns) vs single range
    var isMultiRange = rangeResult && rangeResult.multiRange === true;
    var dataRange;
    var rangeStr;
    
    if (isMultiRange) {
      // For non-contiguous columns, use the domain range as the primary reference
      rangeStr = rangeResult.fullRange;
      dataRange = sheet.getRange(rangeResult.domainRange);
      Logger.log('[SheetActions_Chart] Multi-range mode: domain=' + rangeResult.domainRange + 
                 ', data=' + rangeResult.dataRanges.join(', '));
    } else {
      rangeStr = rangeResult;
      dataRange = sheet.getRange(rangeStr);
      Logger.log('[SheetActions_Chart] Single range mode: ' + rangeStr);
    }
    
    // Determine chart type
    var chartType = (config.chartType || step.chartType || 'column').toLowerCase();
    
    // Chart dimensions (reasonable defaults, larger for pie)
    var defaultWidth = chartType === 'pie' ? 500 : 600;
    var defaultHeight = chartType === 'pie' ? 400 : 400;
    var chartWidth = config.width || defaultWidth;
    var chartHeight = config.height || defaultHeight;
    
    // Position: place chart to the right of data, or at specified position
    var posRow = config.positionRow || 2;
    var posCol;
    if (isMultiRange) {
      // Use the rightmost column of all involved columns + 2
      // Extract column letter from range string (e.g., "E1:E16" → "E", "AA2:AA20" → "AA")
      var maxColIdx = 0;
      var allColRanges = [rangeResult.domainRange].concat(rangeResult.dataRanges);
      allColRanges.forEach(function(r) {
        var colMatch = r.match(/^([A-Z]+)/);
        if (colMatch) {
          var idx = SheetActions_Utils.letterToColumn(colMatch[1]);
          if (idx > maxColIdx) maxColIdx = idx;
        }
      });
      posCol = config.positionColumn || (maxColIdx + 2);
    } else {
      posCol = config.positionColumn || (dataRange.getLastColumn() + 2);
    }
    
    // Build chart — add ranges in correct order (domain FIRST for proper X/Y mapping)
    var chartBuilder = sheet.newChart()
      .setChartType(SheetActions_Utils.getChartType(chartType));
    
    if (isMultiRange) {
      // Add domain (X-axis) range FIRST
      chartBuilder.addRange(sheet.getRange(rangeResult.domainRange));
      // Then add each data (Y-axis) range
      rangeResult.dataRanges.forEach(function(dr) {
        chartBuilder.addRange(sheet.getRange(dr));
      });
      // Add annotation range for scatter point labels (e.g., company names)
      if (rangeResult.annotationRange) {
        chartBuilder.addRange(sheet.getRange(rangeResult.annotationRange));
        Logger.log('[SheetActions_Chart] Added annotation range: ' + rangeResult.annotationRange);
      }
      Logger.log('[SheetActions_Chart] Added ' + (1 + rangeResult.dataRanges.length) + ' separate ranges (domain first)');
    } else {
      chartBuilder.addRange(dataRange);
    }
    
    // Tell the chart that headers are present.
    // Use -1 (auto-detect) rather than a fixed number so the chart correctly
    // identifies headers regardless of whether the range starts at row 1 or 2.
    chartBuilder.setNumHeaders(-1);
    
    chartBuilder.setPosition(posRow, posCol, 0, 0);
    
    // Set chart dimensions for better visibility
    chartBuilder.setOption('width', chartWidth);
    chartBuilder.setOption('height', chartHeight);
    
    // ===== TITLE =====
    if (config.title) {
      chartBuilder.setOption('title', config.title);
      chartBuilder.setOption('titleTextStyle', { 
        fontSize: 16, 
        bold: true,
        color: config.titleColor || '#333333'
      });
    }
    
    // Font family
    if (config.fontName) {
      chartBuilder.setOption('fontName', config.fontName);
    }
    
    // ===== COLORS =====
    var colors = config.colors || SheetActions_Utils.DEFAULT_COLORS;
    chartBuilder.setOption('colors', colors);
    
    // ===== LEGEND =====
    var legendPosition = config.legendPosition || 'top';
    chartBuilder.setOption('legend', { 
      position: legendPosition,
      textStyle: { fontSize: 12, color: '#666666' },
      alignment: 'center'
    });
    
    // ===== SERIES CONFIGURATION =====
    if (config.seriesNames && config.seriesNames.length > 0) {
      var seriesConfig = {};
      config.seriesNames.forEach(function(name, idx) {
        seriesConfig[idx] = { 
          labelInLegend: name
        };
        if (config.colors && config.colors[idx]) {
          seriesConfig[idx].color = config.colors[idx];
        }
      });
      chartBuilder.setOption('series', seriesConfig);
      Logger.log('[SheetActions_Chart] Series names: ' + config.seriesNames.join(', '));
    }
    
    // ===== CHART-TYPE SPECIFIC OPTIONS =====
    
    if (chartType === 'pie') {
      _applyPieChartOptions(chartBuilder, config);
      
    } else if (chartType === 'line') {
      _applyLineChartOptions(chartBuilder, config, dataRange);
      _applyAxisOptions(chartBuilder, config, dataRange);
      
    } else if (chartType === 'area') {
      _applyAreaChartOptions(chartBuilder, config);
      _applyAxisOptions(chartBuilder, config, dataRange);
      
    } else if (chartType === 'scatter') {
      _applyScatterChartOptions(chartBuilder, config);
      _applyAxisOptions(chartBuilder, config, dataRange);
      
    } else if (chartType === 'bar' || chartType === 'column') {
      _applyBarColumnChartOptions(chartBuilder, config, chartType);
      _applyAxisOptions(chartBuilder, config, dataRange);
      
    } else if (chartType === 'combo') {
      _applyComboChartOptions(chartBuilder, config);
      _applyAxisOptions(chartBuilder, config, dataRange);
    }
    
    // ===== DUAL AXIS SUPPORT =====
    if (config.secondaryAxis && config.secondaryAxis.length > 0) {
      _applyDualAxisOptions(chartBuilder, config);
    }
    
    // ===== COMMON OPTIONS =====
    
    // Tooltip configuration
    chartBuilder.setOption('tooltip', { 
      showColorCode: true,
      textStyle: { fontSize: 12 },
      trigger: config.tooltipTrigger || 'focus'
    });
    
    // Background color
    if (config.backgroundColor) {
      chartBuilder.setOption('backgroundColor', { fill: config.backgroundColor });
    }
    
    // Chart area padding
    chartBuilder.setOption('chartArea', {
      left: config.chartAreaLeft || '10%',
      top: config.chartAreaTop || '15%',
      width: config.chartAreaWidth || '75%',
      height: config.chartAreaHeight || '70%'
    });
    
    // Animation
    chartBuilder.setOption('animation', {
      startup: true,
      duration: 500,
      easing: 'out'
    });
    
    // Build and insert
    var chart = chartBuilder.build();
    sheet.insertChart(chart);
    
    Logger.log('[SheetActions_Chart] ' + chartType.toUpperCase() + ' chart created with ' + 
               (config.seriesNames ? config.seriesNames.length : 'auto-detected') + ' series');
    
    return { 
      chartType: chartType, 
      range: rangeStr,
      title: config.title || 'Chart',
      series: config.seriesNames || [],
      dimensions: { width: chartWidth, height: chartHeight }
    };
  }
  
  // ============================================
  // CHART-TYPE SPECIFIC OPTIONS
  // ============================================
  
  /**
   * Apply PIE/DONUT chart specific options
   */
  function _applyPieChartOptions(chartBuilder, config) {
    // Donut hole (0 = pie, 0.4 = nice donut)
    if (config.pieHole !== undefined) {
      chartBuilder.setOption('pieHole', config.pieHole);
    } else if (config.donut) {
      chartBuilder.setOption('pieHole', 0.4);
    }
    
    // Slice text
    if (config.pieSliceText) {
      chartBuilder.setOption('pieSliceText', config.pieSliceText);
    }
    chartBuilder.setOption('pieSliceTextStyle', {
      fontSize: 12,
      color: config.pieSliceTextColor || '#ffffff',
      bold: true
    });
    
    // Slice border
    if (config.pieSliceBorderColor) {
      chartBuilder.setOption('pieSliceBorderColor', config.pieSliceBorderColor);
    }
    
    // Starting angle
    if (config.pieStartAngle !== undefined) {
      chartBuilder.setOption('pieStartAngle', config.pieStartAngle);
    }
    
    // 3D effect
    if (config.is3D) {
      chartBuilder.setOption('is3D', true);
    }
    
    // Hide small slices
    if (config.sliceVisibilityThreshold) {
      chartBuilder.setOption('sliceVisibilityThreshold', config.sliceVisibilityThreshold);
    }
    
    // Individual slice configuration
    if (config.slices) {
      chartBuilder.setOption('slices', config.slices);
    }
    
    Logger.log('[SheetActions_Chart] Pie options applied: hole=' + (config.pieHole || 0));
  }
  
  /**
   * Apply LINE chart specific options
   */
  function _applyLineChartOptions(chartBuilder, config, dataRange) {
    // Curve type
    var curveType = config.curveType || 'none';
    if (curveType === 'smooth' || curveType === 'curved' || curveType === 'spline') {
      curveType = 'function';
    }
    chartBuilder.setOption('curveType', curveType);
    
    // Line width
    var lineWidth = config.lineWidth !== undefined ? config.lineWidth : 2;
    chartBuilder.setOption('lineWidth', lineWidth);
    
    // Point size
    var pointSize = config.pointSize !== undefined ? config.pointSize : 5;
    chartBuilder.setOption('pointSize', pointSize);
    
    // Point shape
    if (config.pointShape) {
      chartBuilder.setOption('pointShape', config.pointShape);
    }
    
    // Handle null values
    if (config.interpolateNulls) {
      chartBuilder.setOption('interpolateNulls', true);
    }
    
    // Line dash style
    if (config.lineDashStyle) {
      var numSeries = dataRange.getNumColumns() - 1;
      var seriesConfig = {};
      for (var i = 0; i < numSeries; i++) {
        seriesConfig[i] = { lineDashStyle: config.lineDashStyle };
      }
      chartBuilder.setOption('series', seriesConfig);
    }
    
    // Focus target
    chartBuilder.setOption('focusTarget', config.focusTarget || 'datum');
    
    // Crosshair
    if (config.crosshair) {
      chartBuilder.setOption('crosshair', {
        trigger: 'both',
        orientation: 'both',
        color: '#cccccc',
        opacity: 0.7
      });
    }
    
    chartBuilder.setOption('useFirstColumnAsDomain', true);
    
    Logger.log('[SheetActions_Chart] Line options: curve=' + curveType + ', lineWidth=' + lineWidth);
  }
  
  /**
   * Apply AREA chart specific options
   */
  function _applyAreaChartOptions(chartBuilder, config) {
    // Area opacity
    var areaOpacity = config.areaOpacity !== undefined ? config.areaOpacity : 0.3;
    chartBuilder.setOption('areaOpacity', areaOpacity);
    
    // Stacked area
    if (config.stacked) {
      chartBuilder.setOption('isStacked', true);
    }
    if (config.stackedPercent) {
      chartBuilder.setOption('isStacked', 'percent');
    }
    
    // Line width
    var lineWidth = config.lineWidth !== undefined ? config.lineWidth : 2;
    chartBuilder.setOption('lineWidth', lineWidth);
    
    // Point size
    if (config.pointSize !== undefined) {
      chartBuilder.setOption('pointSize', config.pointSize);
    }
    
    chartBuilder.setOption('useFirstColumnAsDomain', true);
    chartBuilder.setOption('focusTarget', config.focusTarget || 'category');
    
    Logger.log('[SheetActions_Chart] Area options: opacity=' + areaOpacity);
  }
  
  /**
   * Apply SCATTER chart specific options
   */
  function _applyScatterChartOptions(chartBuilder, config) {
    // Point size
    var pointSize = config.pointSize !== undefined ? config.pointSize : 7;
    chartBuilder.setOption('pointSize', pointSize);
    
    // Point shape
    var pointShape = config.pointShape || 'circle';
    chartBuilder.setOption('pointShape', pointShape);
    
    // No lines
    chartBuilder.setOption('lineWidth', 0);
    
    // Trendlines
    if (config.trendlines || config.trendline) {
      var trendlineConfig = {};
      
      if (Array.isArray(config.trendlines)) {
        config.trendlines.forEach(function(tl, idx) {
          var seriesIndex = tl.series !== undefined ? tl.series : idx;
          trendlineConfig[seriesIndex] = {
            type: tl.type || 'linear',
            color: tl.color || SheetActions_Utils.DEFAULT_COLORS[seriesIndex % SheetActions_Utils.DEFAULT_COLORS.length],
            lineWidth: tl.lineWidth || 2,
            opacity: tl.opacity || 0.6,
            showR2: tl.showR2 || false,
            visibleInLegend: tl.visibleInLegend !== false,
            labelInLegend: tl.labelInLegend
          };
        });
      } else {
        trendlineConfig = {
          0: {
            type: config.trendlineType || 'linear',
            color: config.trendlineColor || '#888888',
            lineWidth: config.trendlineWidth || 2,
            opacity: config.trendlineOpacity || 0.5,
            showR2: config.showR2 || false,
            visibleInLegend: config.trendlineInLegend !== false
          }
        };
      }
      chartBuilder.setOption('trendlines', trendlineConfig);
    }
    
    // Aggregate duplicate points
    if (config.aggregationTarget) {
      chartBuilder.setOption('aggregationTarget', config.aggregationTarget);
    }
    
    // Crosshair
    if (config.crosshair !== false) {
      chartBuilder.setOption('crosshair', {
        trigger: 'both',
        orientation: 'both',
        color: '#dddddd',
        opacity: 0.5
      });
    }
    
    // CRITICAL: Tell chart to use first column (domain) as X-axis
    chartBuilder.setOption('useFirstColumnAsDomain', true);
    
    // Point labels / annotations (e.g., company names on data points)
    if (config.annotationColumn || config.pointLabelsColumn) {
      chartBuilder.setOption('annotations', {
        textStyle: { 
          fontSize: config.annotationFontSize || 9,
          color: config.annotationColor || '#333333',
          auraColor: 'none'
        },
        stem: { 
          color: '#999999', 
          length: config.annotationStemLength !== undefined ? config.annotationStemLength : 5 
        },
        alwaysOutside: false
      });
      Logger.log('[SheetActions_Chart] Scatter annotations enabled for column: ' + (config.annotationColumn || config.pointLabelsColumn));
    }
    
    Logger.log('[SheetActions_Chart] Scatter options: pointSize=' + pointSize);
  }
  
  /**
   * Apply BAR/COLUMN chart specific options
   */
  function _applyBarColumnChartOptions(chartBuilder, config, chartType) {
    // Stacked
    if (config.stacked) {
      chartBuilder.setOption('isStacked', true);
    }
    if (config.stackedPercent) {
      chartBuilder.setOption('isStacked', 'percent');
    }
    
    // Bar width
    var barWidth = config.barGroupWidth || '75%';
    chartBuilder.setOption('bar', { groupWidth: barWidth });
    
    // Data labels
    if (config.showDataLabels) {
      chartBuilder.setOption('annotations', {
        alwaysOutside: true,
        textStyle: { 
          fontSize: 10,
          color: '#333333',
          auraColor: 'none'
        },
        stem: { color: 'transparent', length: 0 }
      });
    }
    
    chartBuilder.setOption('useFirstColumnAsDomain', true);
    chartBuilder.setOption('focusTarget', config.focusTarget || 'category');
    
    Logger.log('[SheetActions_Chart] ' + chartType + ' options: stacked=' + (config.stacked || false));
  }
  
  /**
   * Apply COMBO chart specific options (NEW)
   */
  function _applyComboChartOptions(chartBuilder, config) {
    // Series types for combo charts
    if (config.seriesTypes) {
      var seriesConfig = {};
      config.seriesTypes.forEach(function(type, idx) {
        seriesConfig[idx] = seriesConfig[idx] || {};
        seriesConfig[idx].type = type; // 'line', 'bars', 'area'
      });
      chartBuilder.setOption('series', seriesConfig);
    }
    
    chartBuilder.setOption('useFirstColumnAsDomain', true);
    chartBuilder.setOption('focusTarget', config.focusTarget || 'category');
    
    Logger.log('[SheetActions_Chart] Combo chart options applied');
  }
  
  /**
   * Apply dual axis support (NEW)
   */
  function _applyDualAxisOptions(chartBuilder, config) {
    var seriesConfig = {};
    config.secondaryAxis.forEach(function(seriesIndex) {
      seriesConfig[seriesIndex] = seriesConfig[seriesIndex] || {};
      seriesConfig[seriesIndex].targetAxisIndex = 1;
    });
    chartBuilder.setOption('series', seriesConfig);
    
    // Configure secondary axis
    chartBuilder.setOption('vAxes', {
      0: { title: config.yAxisTitle || '' },
      1: { title: config.secondaryAxisTitle || '' }
    });
    
    Logger.log('[SheetActions_Chart] Dual axis configured: secondary series=' + config.secondaryAxis.join(','));
  }
  
  /**
   * Apply AXIS options (common for bar, column, line, area, scatter)
   */
  function _applyAxisOptions(chartBuilder, config, dataRange) {
    // ===== HORIZONTAL AXIS (X-axis) =====
    var hAxisConfig = {
      textStyle: { fontSize: 11, color: '#666666' },
      titleTextStyle: { fontSize: 12, italic: true, color: '#333333' }
    };
    
    if (config.xAxisTitle) {
      hAxisConfig.title = config.xAxisTitle;
    }
    
    if (config.xAxisFormat) {
      hAxisConfig.format = config.xAxisFormat;
    }
    
    if (config.xAxisMin !== undefined) {
      hAxisConfig.minValue = config.xAxisMin;
    }
    if (config.xAxisMax !== undefined) {
      hAxisConfig.maxValue = config.xAxisMax;
    }
    
    // Log scale (NEW)
    if (config.xAxisLogScale) {
      hAxisConfig.logScale = true;
    }
    
    // Slant labels
    if (config.slantedTextAngle) {
      hAxisConfig.slantedText = true;
      hAxisConfig.slantedTextAngle = config.slantedTextAngle;
    } else if (dataRange.getNumRows() > 8) {
      hAxisConfig.slantedText = true;
      hAxisConfig.slantedTextAngle = 45;
    }
    
    if (dataRange.getNumRows() > 12) {
      hAxisConfig.textStyle = { fontSize: 10, color: '#666666' };
    }
    
    if (config.gridlines !== false) {
      hAxisConfig.gridlines = { 
        color: config.gridlineColor || '#f0f0f0',
        count: config.xGridlineCount || -1
      };
    }
    
    chartBuilder.setOption('hAxis', hAxisConfig);
    
    // ===== VERTICAL AXIS (Y-axis) =====
    var vAxisConfig = {
      textStyle: { fontSize: 11, color: '#666666' },
      titleTextStyle: { fontSize: 12, italic: true, color: '#333333' }
    };
    
    if (config.yAxisTitle) {
      vAxisConfig.title = config.yAxisTitle;
    }
    
    if (config.yAxisFormat || config.valueFormat) {
      vAxisConfig.format = config.yAxisFormat || config.valueFormat;
    }
    
    if (config.yAxisMin !== undefined) {
      vAxisConfig.minValue = config.yAxisMin;
    }
    if (config.yAxisMax !== undefined) {
      vAxisConfig.maxValue = config.yAxisMax;
    }
    
    // Log scale (NEW)
    if (config.logScale || config.yAxisLogScale) {
      vAxisConfig.logScale = true;
    }
    
    if (config.gridlines !== false) {
      vAxisConfig.gridlines = { 
        color: config.gridlineColor || '#f0f0f0',
        count: config.yGridlineCount || 5
      };
      vAxisConfig.minorGridlines = {
        color: config.minorGridlineColor || '#f8f8f8',
        count: config.minorGridlineCount || 0
      };
    }
    
    if (config.baseline !== undefined) {
      vAxisConfig.baseline = config.baseline;
    }
    
    chartBuilder.setOption('vAxis', vAxisConfig);
    
    Logger.log('[SheetActions_Chart] Axes configured');
  }
  
  // ============================================
  // HELPER FUNCTIONS
  // ============================================
  
  /**
   * Build data range from AI-provided columns.
   * 
   * Returns either a single range string (contiguous columns) or an object
   * with individual column ranges for non-contiguous columns:
   *   { multiRange: true, domainRange: 'E1:E16', dataRanges: ['C1:C16'], fullRange: 'C1:E16' }
   * 
   * This ensures the chart correctly maps domain=X-axis and data=Y-axis
   * even when columns are non-adjacent (e.g., domainColumn=E, dataColumns=["C"]).
   */
  function _buildDataRangeFromColumns(sheet, config, step) {
    var domainColumn = config.domainColumn || config.xAxisColumn || config.labelsColumn;
    var dataColumns = config.dataColumns;
    if (!dataColumns && config.valuesColumn) {
      dataColumns = [config.valuesColumn];
    }
    
    // Validate dataColumns: filter out non-numeric columns
    // AI may send text columns (Website, Email, Description) which produce empty charts
    if (domainColumn && dataColumns && dataColumns.length > 0) {
      var validatedDataColumns = [];
      var lastRow = sheet.getLastRow();
      var sampleStartRow = 2; // Skip header
      var sampleEndRow = Math.min(sampleStartRow + 4, lastRow); // Check 5 rows
      
      for (var dc = 0; dc < dataColumns.length; dc++) {
        var col = dataColumns[dc];
        try {
          var colIdx = SheetActions_Utils.letterToColumn(col);
          if (sampleEndRow >= sampleStartRow) {
            var sampleVals = sheet.getRange(sampleStartRow, colIdx, sampleEndRow - sampleStartRow + 1, 1).getValues();
            var numCount = 0;
            var totalNonEmpty = 0;
            for (var sv = 0; sv < sampleVals.length; sv++) {
              var val = sampleVals[sv][0];
              if (val === '' || val === null) continue;
              totalNonEmpty++;
              if (typeof val === 'number') {
                numCount++;
              } else if (typeof val === 'string') {
                // Check if it's a number with formatting (e.g., "$45.2", "1,200")
                var cleaned = val.replace(/[$€£¥₹%,\s]/g, '');
                if (cleaned !== '' && !isNaN(parseFloat(cleaned))) {
                  numCount++;
                }
              }
            }
            if (totalNonEmpty > 0 && numCount >= totalNonEmpty / 2) {
              validatedDataColumns.push(col);
            } else {
              Logger.log('[SheetActions_Chart] ⚠️ Skipping non-numeric column ' + col + ' from dataColumns (' + numCount + '/' + totalNonEmpty + ' numeric)');
            }
          } else {
            validatedDataColumns.push(col); // Can't validate, assume OK
          }
        } catch (e) {
          Logger.log('[SheetActions_Chart] Error validating column ' + col + ': ' + e.message);
          validatedDataColumns.push(col); // Don't block on error
        }
      }
      
      if (validatedDataColumns.length === 0 && dataColumns.length > 0) {
        Logger.log('[SheetActions_Chart] ⚠️ All dataColumns were non-numeric! Attempting auto-detection...');
        // Auto-detect numeric columns from the sheet
        var autoDetected = _autoDetectNumericColumns(sheet, domainColumn);
        if (autoDetected.length > 0) {
          validatedDataColumns = autoDetected;
          Logger.log('[SheetActions_Chart] Auto-detected numeric columns: ' + validatedDataColumns.join(', '));
        } else {
          // Last resort: use original columns
          validatedDataColumns = dataColumns;
          Logger.log('[SheetActions_Chart] No numeric columns found, using original dataColumns');
        }
      } else if (validatedDataColumns.length < dataColumns.length) {
        Logger.log('[SheetActions_Chart] Filtered dataColumns: ' + dataColumns.join(',') + ' → ' + validatedDataColumns.join(','));
      }
      
      dataColumns = validatedDataColumns;
    }
    
    // If we have validated dataColumns, proceed to build ranges
    if (domainColumn && dataColumns && dataColumns.length > 0) {
      var chartType = (config.chartType || 'column').toLowerCase();
      Logger.log('[SheetActions_Chart] Using unified column approach for ' + chartType);
      Logger.log('[SheetActions_Chart] Domain: ' + domainColumn + ', Data: ' + dataColumns.join(','));
      
      // Build ordered column list: domain first, then data columns
      var allColumns = [domainColumn].concat(dataColumns);
      var colIndices = allColumns.map(function(col) {
        return SheetActions_Utils.letterToColumn(col);
      });
      
      // Detect data boundaries using min/max of all columns
      var minCol = Math.min.apply(null, colIndices);
      var maxCol = Math.max.apply(null, colIndices);
      var dataInfo = SheetActions_Utils.detectDataBoundaries(sheet, minCol, maxCol);
      var startRow = dataInfo.startRow;
      var endRow = dataInfo.endRow;
      
      // Check if columns are contiguous and in the right order (domain first)
      var isContiguous = (maxCol - minCol + 1) === allColumns.length;
      var domainIsFirst = colIndices[0] === minCol;
      
      // Check for annotation column (scatter point labels like company names)
      var annotationCol = config.annotationColumn || config.pointLabelsColumn;
      
      if (isContiguous && domainIsFirst && !annotationCol) {
        // Simple case: columns are adjacent and domain is first — use single range
        var rangeStr = SheetActions_Utils.columnToLetter(minCol) + startRow + ':' + 
                       SheetActions_Utils.columnToLetter(maxCol) + endRow;
        Logger.log('[SheetActions_Chart] Contiguous range: ' + rangeStr);
        return rangeStr;
      }
      
      // If contiguous but has annotation column, force multi-range mode so we can add annotation
      if (isContiguous && domainIsFirst && annotationCol) {
        Logger.log('[SheetActions_Chart] Contiguous columns but annotation column present - using multi-range mode');
      }
      
      // Non-contiguous or wrong order: build individual ranges
      // Domain column range (X-axis) — must be added FIRST to chart
      var domainRangeStr = domainColumn + startRow + ':' + domainColumn + endRow;
      
      // Data column ranges (Y-axis series)
      var dataRangeStrs = dataColumns.map(function(col) {
        return col + startRow + ':' + col + endRow;
      });
      
      // Full range for fallback/logging
      var fullRangeStr = SheetActions_Utils.columnToLetter(minCol) + startRow + ':' + 
                         SheetActions_Utils.columnToLetter(maxCol) + endRow;
      
      // Annotation column for scatter point labels (e.g., company names)
      var annotationColumn = config.annotationColumn || config.pointLabelsColumn;
      var annotationRangeStr = null;
      if (annotationColumn && chartType === 'scatter') {
        annotationRangeStr = annotationColumn + startRow + ':' + annotationColumn + endRow;
        Logger.log('[SheetActions_Chart]   Annotation (labels): ' + annotationColumn + ' → ' + annotationRangeStr);
      }
      
      Logger.log('[SheetActions_Chart] Non-contiguous columns detected!');
      Logger.log('[SheetActions_Chart]   Domain (X): ' + domainColumn + ' → ' + domainRangeStr);
      Logger.log('[SheetActions_Chart]   Data (Y): ' + dataColumns.join(',') + ' → ' + dataRangeStrs.join(', '));
      
      return {
        multiRange: true,
        domainRange: domainRangeStr,
        dataRanges: dataRangeStrs,
        annotationRange: annotationRangeStr,
        fullRange: fullRangeStr,
        startRow: startRow,
        endRow: endRow
      };
    }
    
    // Fallback: If no domainColumn/dataColumns provided, try to auto-detect them
    // This prevents charting ALL columns (including text) which produces empty charts
    if (!domainColumn || !dataColumns || dataColumns.length === 0) {
      Logger.log('[SheetActions_Chart] No domainColumn/dataColumns provided — attempting auto-detection');
      var autoNumeric = _autoDetectNumericColumns(sheet, null);
      if (autoNumeric.length > 0) {
        // Find first text column for domain
        var lastCol = sheet.getLastColumn();
        var lastRow = sheet.getLastRow();
        var sampleRows = Math.min(5, lastRow - 1);
        for (var ac = 1; ac <= lastCol; ac++) {
          var colL = SheetActions_Utils.columnToLetter(ac);
          if (autoNumeric.indexOf(colL) === -1) {
            var vals = sheet.getRange(2, ac, sampleRows, 1).getValues();
            var hasData = vals.some(function(r) { return r[0] !== '' && r[0] !== null; });
            if (hasData) {
              domainColumn = colL;
              dataColumns = autoNumeric;
              config.domainColumn = domainColumn;
              config.dataColumns = dataColumns;
              // Generate series names from headers
              var headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
              config.seriesNames = dataColumns.map(function(dc) {
                var idx = SheetActions_Utils.letterToColumn(dc) - 1;
                return headerValues[idx] || dc;
              });
              Logger.log('[SheetActions_Chart] Auto-detected: domain=' + domainColumn + ', data=' + dataColumns.join(','));
              // Re-enter the column-based path
              return _buildDataRangeFromColumns(sheet, config, step);
            }
          }
        }
      }
    }
    
    // Fallback: Legacy approach
    var rangeStr = config.dataRange || step.dataRange || step.inputRange;
    if (rangeStr) {
      var dataRange = sheet.getRange(rangeStr);
      rangeStr = _ensureDateColumnIncluded(sheet, dataRange, rangeStr, config.title || '');
      return rangeStr;
    }
    
    // Last resort: detect full sheet data
    return SheetActions_Utils.detectDataRange(sheet);
  }
  
  /**
   * Auto-detect numeric columns from the sheet data.
   * Used as fallback when AI sends only text columns as dataColumns.
   * 
   * @param {Sheet} sheet - Active sheet
   * @param {string} excludeColumn - Column to exclude (domain column)
   * @return {string[]} Array of column letters that contain numeric data
   */
  function _autoDetectNumericColumns(sheet, excludeColumn) {
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    if (lastCol < 1 || lastRow < 2) return [];
    
    var numericCols = [];
    var sampleRows = Math.min(5, lastRow - 1);
    
    for (var c = 1; c <= lastCol; c++) {
      var colLetter = SheetActions_Utils.columnToLetter(c);
      if (colLetter === excludeColumn) continue;
      
      var vals = sheet.getRange(2, c, sampleRows, 1).getValues();
      var numCount = 0;
      var totalNonEmpty = 0;
      
      for (var r = 0; r < vals.length; r++) {
        var val = vals[r][0];
        if (val === '' || val === null) continue;
        totalNonEmpty++;
        if (typeof val === 'number') {
          numCount++;
        } else if (typeof val === 'string') {
          var cleaned = val.replace(/[$€£¥₹%,\s]/g, '');
          if (cleaned !== '' && !isNaN(parseFloat(cleaned))) {
            numCount++;
          }
        }
      }
      
      if (totalNonEmpty > 0 && numCount >= totalNonEmpty / 2) {
        numericCols.push(colLetter);
      }
    }
    
    // Limit to first 3 numeric columns to keep chart clean
    return numericCols.slice(0, 3);
  }
  
  /**
   * Ensure the date/label column is included in the chart range
   */
  function _ensureDateColumnIncluded(sheet, range, rangeStr, title) {
    try {
      var startCol = range.getColumn();
      var startRow = range.getRow();
      var numRows = range.getNumRows();
      var numCols = range.getNumColumns();
      
      if (startCol <= 1) {
        return rangeStr;
      }
      
      var titleLower = (title || '').toLowerCase();
      var isTimeBased = titleLower.indexOf('trend') !== -1 || 
                        titleLower.indexOf('time') !== -1 || 
                        titleLower.indexOf('month') !== -1 ||
                        titleLower.indexOf('year') !== -1 ||
                        titleLower.indexOf('revenue') !== -1;
      
      var firstColValues = range.getValues();
      var numericCount = 0;
      var dateOrTextCount = 0;
      
      for (var i = 1; i < Math.min(firstColValues.length, 6); i++) {
        var val = firstColValues[i][0];
        if (val === null || val === '') continue;
        
        if (val instanceof Date) {
          dateOrTextCount++;
          continue;
        }
        
        if (typeof val === 'number') {
          numericCount++;
        } else if (typeof val === 'string') {
          if (val.match(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/) || 
              val.match(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i) ||
              val.match(/^(Tháng|thang)/i)) {
            dateOrTextCount++;
          } else {
            var cleanVal = val.replace(/[$€£¥₫đ%,.\s]/g, '');
            if (!isNaN(parseFloat(cleanVal)) && cleanVal.length > 0) {
              numericCount++;
            } else {
              dateOrTextCount++;
            }
          }
        }
      }
      
      if (numericCount > dateOrTextCount && isTimeBased) {
        var prevCol = startCol - 1;
        var prevColRange = sheet.getRange(startRow, prevCol, numRows, 1);
        var prevColValues = prevColRange.getValues();
        
        var prevHasDates = false;
        for (var j = 1; j < Math.min(prevColValues.length, 5); j++) {
          var pv = prevColValues[j][0];
          if (pv instanceof Date) {
            prevHasDates = true;
            break;
          }
          if (typeof pv === 'string' && 
              (pv.match(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/) ||
               pv.match(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i) ||
               pv.match(/^(Tháng|thang)/i))) {
            prevHasDates = true;
            break;
          }
        }
        
        if (prevHasDates) {
          var newRange = SheetActions_Utils.columnToLetter(prevCol) + startRow + ':' + 
                         SheetActions_Utils.columnToLetter(startCol + numCols - 1) + (startRow + numRows - 1);
          return newRange;
        }
      }
      
      return rangeStr;
    } catch (e) {
      Logger.log('[SheetActions_Chart] Error in _ensureDateColumnIncluded: ' + e.message);
      return rangeStr;
    }
  }
  
  // ============================================
  // PUBLIC API
  // ============================================
  
  return {
    createChart: createChart
  };
  
})();
