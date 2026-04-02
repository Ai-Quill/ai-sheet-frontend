/**
 * @file AgentTaskChain_Execute.gs
 * @version 2.0.0
 * @updated 2026-02-10
 * @author AISheeter Team
 * 
 * ============================================
 * AGENT TASK CHAIN - Execution Flow
 * ============================================
 * 
 * Handles the main execution flow of task chains:
 * - Initialize chain state and start execution
 * - Execute steps sequentially (sheet actions, formulas, AI jobs, analyze)
 * - Handle step completion callbacks
 * 
 * Part of the AgentTaskChain module suite:
 * - AgentTaskChain_Parse.gs    (parsing & suggestions)
 * - AgentTaskChain_Execute.gs  (this file)
 * - AgentTaskChain_Plan.gs     (plan building)
 * - AgentTaskChain_State.gs    (state management)
 * - AgentTaskChain_Analyze.gs  (analyze/chat steps)
 * 
 * Depends on:
 * - Agent.gs (executeAgentPlan, getAgentModel, getAgentContext)
 * - AgentTaskChain_Plan.gs (buildPlanForChainStep)
 * - AgentTaskChain_State.gs (saveChainState, getChainState)
 * - AgentTaskChain_Analyze.gs (_executeAnalyzeStep)
 * - SheetActions_Main.gs (AgentSheetActions)
 * - Jobs.gs (writeJobResults)
 */

// ============================================
// TASK CHAIN EXECUTION
// ============================================

/**
 * Execute a multi-step task chain
 * 
 * @param {Object} chain - Parsed task chain { steps, ... }
 * @param {Object} context - Spreadsheet context
 * @return {Object} Execution result with chain job ID
 */
/**
 * Execute a task chain
 * REUSES the same context/selection logic as single-step commands
 * 
 * @param {Object} chain - Parsed chain with steps
 * @param {Object} context - Selection context from frontend
 * @return {Object} Result with chainId
 */
function executeTaskChain(chain, context) {
  if (!chain || !chain.steps || chain.steps.length === 0) {
    throw new Error('Invalid task chain');
  }
  
  Logger.log('[TaskChain] 🔗 Starting chain with ' + chain.steps.length + ' steps');
  
  // DEBUG: Log the chain object as received
  Logger.log('[TaskChain] 📥 Chain received from frontend:');
  Logger.log('  chain.inputRange: ' + chain.inputRange);
  Logger.log('  chain.inputColumn: ' + chain.inputColumn);
  Logger.log('  chain.inputColumns: ' + JSON.stringify(chain.inputColumns));
  Logger.log('  chain.inputColumns type: ' + typeof chain.inputColumns);
  Logger.log('  chain.inputColumns isArray: ' + Array.isArray(chain.inputColumns));
  Logger.log('  chain.hasMultipleInputColumns: ' + chain.hasMultipleInputColumns);
  
  // Log ALL steps to see what we received
  Logger.log('[TaskChain] 📋 All steps received:');
  chain.steps.forEach(function(step, idx) {
    Logger.log('  Step ' + (idx + 1) + ': ' + step.action + ' | desc: ' + (step.description || 'N/A').substring(0, 40));
    Logger.log('    inputColumns: ' + JSON.stringify(step.inputColumns));
    Logger.log('    outputColumn: ' + step.outputColumn);
  });
  
  // DEFENSIVE: Ensure inputColumns is an array (google.script.run can mangle arrays)
  if (chain.inputColumns && !Array.isArray(chain.inputColumns)) {
    Logger.log('[TaskChain] ⚠️ inputColumns is not an array! Converting...');
    if (typeof chain.inputColumns === 'object') {
      chain.inputColumns = Object.values(chain.inputColumns);
    } else if (typeof chain.inputColumns === 'string') {
      chain.inputColumns = chain.inputColumns.split(',');
    }
    Logger.log('[TaskChain] Converted to: ' + JSON.stringify(chain.inputColumns));
  }
  
  // === GET CONTEXT IF NOT PROVIDED ===
  if (!context || Object.keys(context).length === 0) {
    Logger.log('[TaskChain] No context provided, getting fresh context');
    context = getAgentContext();
  }
  
  Logger.log('[TaskChain] Context keys: ' + JSON.stringify(Object.keys(context)));
  
  // === EXTRACT INPUT INFORMATION ===
  var inputRange = null;
  var inputColumn = null;
  var inputColumns = [];
  var inputColumnHeaders = {};
  var hasMultipleInputColumns = false;
  var rowCount = 0;
  
  // PRIORITY 0: Chain-level configuration (set by backend)
  if (chain.inputRange) {
    Logger.log('[TaskChain] ✨ Using chain-level input config from backend');
    inputRange = chain.inputRange;
    inputColumn = chain.inputColumn || (chain.inputColumns && chain.inputColumns[0]) || inputRange.charAt(0);
    inputColumns = chain.inputColumns || [inputColumn];
    hasMultipleInputColumns = chain.hasMultipleInputColumns || inputColumns.length > 1;
    rowCount = chain.rowCount || 0;
    Logger.log('[TaskChain] Chain config: inputRange=' + inputRange + ', columns=' + inputColumns.join(','));
    
    if ((!chain.inputColumns || chain.inputColumns.length === 0) && context.selectionInfo && context.selectionInfo.columnsWithData) {
      inputColumns = context.selectionInfo.columnsWithData;
      inputColumn = inputColumns[0];
      hasMultipleInputColumns = inputColumns.length > 1;
      Logger.log('[TaskChain] ⚡ Enhanced columns from context: ' + inputColumns.join(','));
    }
  }
  // Priority 1: Use selectionInfo if available
  else if (context.selectionInfo) {
    var info = context.selectionInfo;
    inputRange = info.dataRange || info.range;
    
    if (info.columnsWithData && info.columnsWithData.length > 0) {
      inputColumns = info.columnsWithData;
      inputColumn = inputColumns[0];
      hasMultipleInputColumns = inputColumns.length > 1;
      
      if (info.selectedHeaders) {
        info.selectedHeaders.forEach(function(h) {
          if (h.column && h.name) {
            inputColumnHeaders[h.column] = h.name;
          }
        });
      }
    }
    
    rowCount = info.numRows || info.rowCount || 0;
    Logger.log('[TaskChain] Using selectionInfo: ' + inputRange + ' (' + (hasMultipleInputColumns ? 'multi-column' : 'single-column') + ')');
  }
  // Priority 2: Use selected range
  else if (context.selectedRange) {
    inputRange = context.selectedRange;
    inputColumn = inputRange.charAt(0);
    Logger.log('[TaskChain] Using selectedRange: ' + inputRange);
  }
  // Priority 3: Use columnDataRanges
  else if (context.columnDataRanges && typeof context.columnDataRanges === 'object') {
    var colKeys = Object.keys(context.columnDataRanges);
    if (colKeys.length > 0) {
      inputColumn = colKeys[0];
      inputRange = context.columnDataRanges[inputColumn].dataRange || context.columnDataRanges[inputColumn].range;
      rowCount = context.columnDataRanges[inputColumn].rowCount || 0;
      Logger.log('[TaskChain] Using columnDataRanges: ' + inputRange);
    }
  }
  
  // Priority 4: Fallback to current selection
  if (!inputRange) {
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var selection = sheet.getActiveRange();
      if (selection) {
        inputRange = selection.getA1Notation();
        inputColumn = inputRange.charAt(0);
        rowCount = selection.getNumRows();
        Logger.log('[TaskChain] Using active selection fallback: ' + inputRange);
      }
    } catch (e) {
      Logger.log('[TaskChain] Could not get active selection: ' + e.message);
    }
  }
  
  if (!inputRange) {
    throw new Error('No input range found. Please select a range with data before running the task chain.');
  }
  
  Logger.log('[TaskChain] 📊 Input config:');
  Logger.log('  inputRange: ' + inputRange);
  Logger.log('  inputColumn: ' + inputColumn);
  Logger.log('  inputColumns: ' + JSON.stringify(inputColumns));
  Logger.log('  hasMultipleInputColumns: ' + hasMultipleInputColumns);
  Logger.log('  rowCount: ' + rowCount);
  
  // === CREATE CHAIN STATE ===
  var chainId = 'chain_' + Utilities.getUuid().substring(0, 8);
  
  var model = chain.model || getAgentModel();
  Logger.log('[TaskChain] 🤖 Using model: ' + model);
  
  var chainState = {
    chainId: chainId,
    totalSteps: chain.steps.length,
    currentStep: 0,
    status: 'running',
    model: model,
    addHeaders: chain.addHeaders || false,
    inputRange: inputRange,
    inputColumn: inputColumn,
    inputColumns: inputColumns,
    inputColumnHeaders: inputColumnHeaders,
    hasMultipleInputColumns: hasMultipleInputColumns,
    rowCount: rowCount,
    sheetName: context.sheetName || SpreadsheetApp.getActiveSheet().getName(),
    // Preserve chain-level sheetConfig (used as fallback for chart/format steps)
    sheetConfig: chain.sheetConfig || null,
    steps: chain.steps.map(function(step, idx) {
      Logger.log('[TaskChain] Copying step ' + (idx + 1) + ':');
      Logger.log('  action: ' + step.action);
      Logger.log('  inputColumns: ' + JSON.stringify(step.inputColumns));
      Logger.log('  outputColumn: ' + step.outputColumn);
      Logger.log('  formula: ' + (step.formula ? step.formula.substring(0, 50) + '...' : 'none'));
      Logger.log('  config: ' + (step.config ? JSON.stringify(step.config).substring(0, 100) + '...' : 'none'));
      
      return {
        id: step.id || ('step_' + (idx + 1)),
        order: step.order || (idx + 1),
        action: step.action || 'process',
        description: step.description,
        prompt: step.prompt || step.description,
        dependsOn: step.dependsOn,
        usesResultOf: step.usesResultOf,
        inputColumns: step.inputColumns || null,
        outputColumn: step.outputColumn || null,
        outputFormat: step.outputFormat || null,
        customFormatHint: step.customFormatHint || null,
        formula: step.formula || null,
        startRow: step.startRow || null,
        endRow: step.endRow || null,
        config: step.config || null,
        range: step.range || null,
        formatType: step.formatType || null,
        options: step.options || step.formatOptions || null,
        chartType: step.chartType || null,
        dataRange: step.dataRange || null,
        title: step.title || null,
        rules: step.rules || null,
        criteria: step.criteria || null,
        operations: step.operations || null,
        validationType: step.validationType || null,
        startCell: step.startCell || null,
        data: step.data || null,
        status: 'pending',
        jobId: null,
        allJobIds: [],
        outputRange: null,
        result: null,
        error: null
      };
    }),
    startedAt: new Date().toISOString()
  };
  
  saveChainState(chainState);
  
  Logger.log('[TaskChain] 🚀 Starting execution...');
  executeNextStep(chainState.chainId, chainState);
  
  return {
    success: true,
    chainId: chainId,
    totalSteps: chain.steps.length,
    message: 'Task chain started'
  };
}

/**
 * Execute the next step in a chain
 * REUSES the same logic as single-step execution via executeAgentPlan
 * 
 * @param {string} chainId - Chain ID or chainState object
 * @param {Object} chainStateOrContext - Chain state (if passed from executeTaskChain) or context
 */
function executeNextStep(chainId, chainStateOrContext) {
  Logger.log('[executeNextStep] Called with chainId type: ' + typeof chainId);
  
  // Handle both chainId string and chainState object
  var chainState;
  if (typeof chainId === 'string') {
    // If chainState was passed as second arg, use it directly (avoids re-fetching)
    if (chainStateOrContext && typeof chainStateOrContext === 'object' && chainStateOrContext.chainId) {
      Logger.log('[executeNextStep] Using passed chainState directly');
      chainState = chainStateOrContext;
    } else {
      Logger.log('[executeNextStep] Fetching chainState from storage');
      chainState = getChainState(chainId);
    }
  } else if (chainId && chainId.chainId) {
    chainState = chainId;
    chainId = chainState.chainId;
  }
  
  if (!chainState) {
    Logger.log('[TaskChain] Chain not found: ' + chainId);
    return;
  }
  
  // Find next pending step
  var nextStepIndex = chainState.steps.findIndex(function(s) {
    return s.status === 'pending';
  });
  
  if (nextStepIndex === -1) {
    // All steps complete
    chainState.status = 'completed';
    chainState.completedAt = new Date().toISOString();
    saveChainState(chainState);
    Logger.log('[TaskChain] ✅ Chain completed: ' + chainState.chainId);
    
    // Post-completion review hook: record implicit positive feedback for analysis steps
    _recordImplicitFeedback(chainState);
    
    // Auto-cleanup: remove old completed chain properties to free quota
    // Keep this chain's state briefly so the frontend poll can read it
    try {
      var props = PropertiesService.getUserProperties();
      var activeChains = JSON.parse(props.getProperty('activeChains') || '[]');
      // Remove all chains except the current one and the most recent
      var toKeep = [chainState.chainId];
      activeChains.forEach(function(id) {
        if (id !== chainState.chainId) {
          props.deleteProperty('chain_' + id);
        }
      });
      // Update active list to only current chain
      props.setProperty('activeChains', JSON.stringify(toKeep));
      Logger.log('[TaskChain] Auto-cleanup: kept only current chain, freed ' + (activeChains.length - 1) + ' old chains');
    } catch (cleanupErr) {
      Logger.log('[TaskChain] Auto-cleanup error (non-fatal): ' + cleanupErr.message);
    }
    
    return;
  }
  
  // SAFEGUARD: If this isn't step 1, ensure previous steps are completed
  if (nextStepIndex > 0) {
    var prevStep = chainState.steps[nextStepIndex - 1];
    if (prevStep.status !== 'completed') {
      Logger.log('[TaskChain] ⚠️ SAFEGUARD: Trying to start step ' + (nextStepIndex + 1) + ' but step ' + nextStepIndex + ' is ' + prevStep.status);
      Logger.log('[TaskChain] Previous step jobId: ' + prevStep.jobId);
      // If previous step is still running, don't start next step
      if (prevStep.status === 'running') {
        Logger.log('[TaskChain] Waiting for previous step to complete...');
        return;
      }
    }
  }
  
  var step = chainState.steps[nextStepIndex];
  
  Logger.log('[TaskChain] ▶️ Executing step ' + (nextStepIndex + 1) + '/' + chainState.totalSteps + ': ' + step.action);
  Logger.log('[TaskChain] Step description: ' + step.description);
  Logger.log('[TaskChain] Step status before: ' + step.status);
  Logger.log('[TaskChain] All steps status: ' + chainState.steps.map(function(s) { return s.status; }).join(', '));
  
  // Mark step as running
  chainState.currentStep = nextStepIndex + 1;
  chainState.steps[nextStepIndex].status = 'running';
  saveChainState(chainState);
  
  try {
    // === BUILD A PLAN FOR THIS STEP (reusing same logic as single-step) ===
    var plan = buildPlanForChainStep(step, chainState, nextStepIndex);
    
    if (!plan) {
      throw new Error('Failed to build plan for step');
    }
    
    Logger.log('[TaskChain] 📋 Built plan for step ' + (nextStepIndex + 1) + ':');
    Logger.log('  Step description: ' + step.description);
    Logger.log('  inputRange: ' + plan.inputRange);
    Logger.log('  inputColumn: ' + plan.inputColumn);
    Logger.log('  inputColumns: ' + JSON.stringify(plan.inputColumns));
    Logger.log('  inputColumns length: ' + (plan.inputColumns ? plan.inputColumns.length : 'null'));
    Logger.log('  hasMultipleInputColumns: ' + plan.hasMultipleInputColumns);
    Logger.log('  outputColumns: ' + plan.outputColumns.join(', '));
    Logger.log('  prompt: ' + plan.prompt.substring(0, 80) + '...');
    Logger.log('  model: ' + plan.model);
    
    // Verify we have the right data
    if (plan.hasMultipleInputColumns && (!plan.inputColumns || plan.inputColumns.length === 0)) {
      Logger.log('[TaskChain] ⚠️ WARNING: hasMultipleInputColumns is true but inputColumns is empty!');
    }
    
    // === CHECK IF THIS IS A SHEET ACTION (chart, format, validation, filter) ===
    // Sheet actions execute instantly without AI jobs
    if (AgentSheetActions && AgentSheetActions.isSheetAction(step.action)) {
      Logger.log('[TaskChain] ⚡ Sheet action detected: ' + step.action);
      
      // Build step config for sheet action
      // IMPORTANT: Step-level properties from the AI tool call take precedence over 
      // chain-level sheetConfig, since each step has its own distinct parameters.
      var stepConfig = step.config || {};
      
      // ROBUSTNESS: Also check chain-level sheetConfig as fallback
      // The SDK Agent stores config in step.config, but legacy path stores in chainState.sheetConfig
      var chainSheetConfig = chainState.sheetConfig || {};
      
      var sheetStep = {
        action: step.action,
        inputRange: plan.inputRange,
        dataRange: plan.inputRange,
        // Pass through range from AI tool call (critical for conditionalFormat, filter, etc.)
        range: step.range || stepConfig.range || chainSheetConfig.range || null,
        // Config object with step-level details merged in
        // CRITICAL: If step.config is empty but chainState.sheetConfig has data, use that
        config: (step.config && Object.keys(step.config).length > 0) ? step.config : (chainSheetConfig || {}),
        chartType: step.chartType || stepConfig.chartType || chainSheetConfig.chartType,
        formatType: step.formatType || stepConfig.formatType || chainSheetConfig.formatType,
        validationType: step.validationType || stepConfig.validationType || chainSheetConfig.validationType,
        rules: step.rules || stepConfig.rules || chainSheetConfig.rules,
        criteria: step.criteria || stepConfig.criteria || chainSheetConfig.criteria,
        options: step.options || stepConfig.options || chainSheetConfig.options || {},
        // Pass through other properties that specific actions need
        operations: step.operations || stepConfig.operations || chainSheetConfig.operations,
        startCell: step.startCell || stepConfig.startCell || chainSheetConfig.startCell,
        data: step.data || stepConfig.data || chainSheetConfig.data,
        title: step.title || stepConfig.title || chainSheetConfig.title,
        description: step.description
      };
      
      Logger.log('[TaskChain] Sheet step config: ' + JSON.stringify(sheetStep).substring(0, 500));
      Logger.log('[TaskChain] Step config keys: ' + (step.config ? Object.keys(step.config).join(',') : 'null'));
      Logger.log('[TaskChain] Chain sheetConfig keys: ' + Object.keys(chainSheetConfig).join(','));
      Logger.log('[TaskChain] Step operations: ' + (sheetStep.operations ? JSON.stringify(sheetStep.operations).substring(0, 200) : 'null'));
      Logger.log('[TaskChain] Step options: ' + (sheetStep.options ? JSON.stringify(sheetStep.options).substring(0, 200) : 'null'));
      Logger.log('[TaskChain] Step range: ' + (sheetStep.range || 'null'));
      
      // Chart-specific debug: log domainColumn/dataColumns so we can verify
      // the AI tool call provided proper column config for charting
      if (step.action === 'chart') {
        var chartConfig = sheetStep.config || {};
        Logger.log('[TaskChain] 📊 CHART DEBUG:');
        Logger.log('[TaskChain]   domainColumn: ' + (chartConfig.domainColumn || 'MISSING'));
        Logger.log('[TaskChain]   dataColumns: ' + JSON.stringify(chartConfig.dataColumns || 'MISSING'));
        Logger.log('[TaskChain]   chartType: ' + (chartConfig.chartType || sheetStep.chartType || 'MISSING'));
        Logger.log('[TaskChain]   title: ' + (chartConfig.title || 'MISSING'));
        Logger.log('[TaskChain]   seriesNames: ' + JSON.stringify(chartConfig.seriesNames || 'MISSING'));
        Logger.log('[TaskChain]   inputRange: ' + (sheetStep.inputRange || 'MISSING'));
        Logger.log('[TaskChain]   config keys: ' + Object.keys(chartConfig).join(','));
      }
      
      // Execute sheet action
      var sheetResult = AgentSheetActions.execute(sheetStep);
      
      if (sheetResult.success) {
        Logger.log('[TaskChain] ✅ Sheet action completed: ' + sheetResult.action);
        
        // Mark step as completed immediately (sheet actions execute instantly)
        chainState.steps[nextStepIndex].status = 'completed';
        chainState.steps[nextStepIndex].completedAt = new Date().toISOString();
        chainState.steps[nextStepIndex].result = sheetResult.result;
        saveChainState(chainState);
        
        // Execute next step immediately
        executeNextStep(chainState.chainId, chainState);
      } else {
        Logger.log('[TaskChain] ❌ Sheet action failed: ' + sheetResult.error);
        
        chainState.steps[nextStepIndex].status = 'error';
        chainState.steps[nextStepIndex].error = sheetResult.error;
        chainState.status = 'error';
        saveChainState(chainState);
      }
      
      return;
    }
    
    // === CHECK IF THIS IS A FORMULA STEP ===
    // Support both step.formula (SDK Agent) and step.prompt (legacy) as formula source
    var formulaSource = null;
    Logger.log('[TaskChain] Checking for formula: outputFormat=' + step.outputFormat + 
               ', hasFormula=' + !!step.formula + 
               ', formulaValue=' + (step.formula ? step.formula.substring(0, 50) : 'none'));
    
    if (step.formula) {
      // SDK Agent path - formula from tool call
      formulaSource = step.formula.startsWith('=') ? step.formula : '=' + step.formula;
      Logger.log('[TaskChain] Using step.formula: ' + formulaSource.substring(0, 80));
    } else if (step.prompt && step.prompt.startsWith('=')) {
      // Legacy path - formula in prompt
      formulaSource = step.prompt;
      Logger.log('[TaskChain] Using step.prompt as formula');
    }
    
    if (step.outputFormat === 'formula' && formulaSource) {
      Logger.log('[TaskChain] ⚡ Formula step detected - applying formula directly');
      Logger.log('[TaskChain] Formula template: ' + formulaSource);
      
      // Extract row numbers from inputRange
      var rowNums = plan.inputRange.match(/\d+/g);
      var startRow = step.startRow || (rowNums ? parseInt(rowNums[0]) : 2);
      var endRow = step.endRow || (rowNums && rowNums.length > 1 ? parseInt(rowNums[rowNums.length - 1]) : startRow + 29);
      
      // Apply formula to each row
      var sheet = SpreadsheetApp.getActiveSheet();
      var formulaTemplate = formulaSource; // e.g., "=IF(D{{ROW}}>E{{ROW}}, D{{ROW}}*0.05, 0)"
      var successCount = 0;
      
      for (var row = startRow; row <= endRow; row++) {
        try {
          var formula = formulaTemplate.replace(/\{\{ROW\}\}/g, row);
          var cell = sheet.getRange(step.outputColumn + row);
          cell.setFormula(formula);
          successCount++;
        } catch (formulaError) {
          Logger.log('[TaskChain] Formula application error at row ' + row + ': ' + formulaError.message);
        }
      }
      
      // Write header if requested
      if (chainState.addHeaders && step.outputColumn) {
        var headerRow = startRow - 1;
        if (headerRow >= 1) {
          // Use step.description as header, fallback to generic name
          var headerText = step.description || step.action || 'Result';
          var colNumber = step.outputColumn.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
          var newHeaderCell = sheet.getRange(headerRow, colNumber);
          
          // Copy formatting from existing header (column A) to maintain consistency
          var existingHeaderCell = sheet.getRange(headerRow, 1);
          existingHeaderCell.copyFormatToRange(sheet, colNumber, colNumber, headerRow, headerRow);
          
          // Set the header text
          newHeaderCell.setValue(headerText);
          Logger.log('[TaskChain] 📝 Header written to ' + step.outputColumn + headerRow + ': ' + headerText + ' (with format copied from A' + headerRow + ')');
        }
      }
      
      Logger.log('[TaskChain] ✅ Formula applied to ' + successCount + ' rows');
      
      // Mark step as completed immediately (formulas execute instantly)
      chainState.steps[nextStepIndex].status = 'completed';
      chainState.steps[nextStepIndex].completedAt = new Date().toISOString();
      chainState.steps[nextStepIndex].outputRange = step.outputColumn + startRow + ':' + step.outputColumn + endRow;
      // CRITICAL: Set writeResult so frontend shows correct row count
      chainState.steps[nextStepIndex].writeResult = { 
        successCount: successCount, 
        errorCount: 0,
        isFormula: true 
      };
      saveChainState(chainState);
      
      // Execute next step immediately
      executeNextStep(chainState.chainId, chainState);
      return;
    }
    
    // === CHECK IF THIS IS A CHAT/ANALYZE STEP (NO CELL WRITING) ===
    if (step.outputFormat === 'chat' || step.action === 'analyze') {
      Logger.log('[TaskChain] 💬 Chat/Analyze step detected - will NOT write to cells');
      
      // For chat output, we need to get AI insights but not write to cells
      // The insights will be displayed in the sidebar chat
      try {
        var chatResult = _executeAnalyzeStep(step, plan, chainState);
        
        chainState.steps[nextStepIndex].status = 'completed';
        chainState.steps[nextStepIndex].completedAt = new Date().toISOString();
        chainState.steps[nextStepIndex].chatResponse = chatResult.response || 'Analysis complete';
        // NOTE: Don't duplicate response in writeResult to save PropertiesService space (9KB limit)
        chainState.steps[nextStepIndex].writeResult = { 
          successCount: 0, 
          errorCount: 0,
          isChat: true
        };
        
        Logger.log('[TaskChain] Chat response length: ' + (chatResult.response || '').length + ' chars');
        saveChainState(chainState);
        
        Logger.log('[TaskChain] ✅ Chat/Analyze step completed (no cells written)');
        
        // Execute next step immediately
        executeNextStep(chainState.chainId, chainState);
        return;
      } catch (chatError) {
        Logger.log('[TaskChain] ❌ Chat/Analyze step failed: ' + chatError.message);
        chainState.steps[nextStepIndex].status = 'completed'; // Still mark as completed
        chainState.steps[nextStepIndex].chatResponse = 'Analysis could not be completed: ' + chatError.message;
        chainState.steps[nextStepIndex].writeResult = { 
          successCount: 0, 
          errorCount: 0,
          isChat: true,
          error: chatError.message
        };
        saveChainState(chainState);
        executeNextStep(chainState.chainId, chainState);
        return;
      }
    }
    
    // === EXECUTE USING THE SAME executeAgentPlan FUNCTION ===
    Logger.log('[TaskChain] Calling executeAgentPlan with full plan...');
    Logger.log('[TaskChain] Full plan JSON: ' + JSON.stringify(plan).substring(0, 500));
    var result;
    try {
      result = executeAgentPlan(plan);
      Logger.log('[TaskChain] executeAgentPlan returned: ' + JSON.stringify({
        success: result && result.success,
        jobCount: result && result.jobs && result.jobs.length,
        firstJobId: result && result.jobs && result.jobs[0] && result.jobs[0].jobId,
        firstJobStatus: result && result.jobs && result.jobs[0] && result.jobs[0].status
      }));
    } catch (planError) {
      Logger.log('[TaskChain] ❌ executeAgentPlan threw error: ' + planError.message);
      throw planError;
    }
    
    if (!result || !result.jobs || result.jobs.length === 0) {
      Logger.log('[TaskChain] ❌ No jobs in result!');
      throw new Error('No jobs created for step');
    }
    
    // Store job IDs in chain state
    // Note: executeAgentPlan returns jobs with 'jobId' property (not 'id')
    var jobId = result.jobs[0].jobId;
    if (!jobId) {
      Logger.log('[TaskChain] ❌ Job object missing jobId! Full job: ' + JSON.stringify(result.jobs[0]));
      throw new Error('Job created but missing jobId property');
    }
    Logger.log('[TaskChain] 📝 Storing jobId: ' + jobId);
    chainState.steps[nextStepIndex].jobId = jobId;
    chainState.steps[nextStepIndex].allJobIds = result.jobs.map(function(j) { return j.jobId; }).filter(Boolean);
    
    // CRITICAL: Store multi-aspect metadata for result splitting
    chainState.steps[nextStepIndex].outputColumns = plan.outputColumns;
    chainState.steps[nextStepIndex].outputFormat = step.outputFormat;
    var isMultiAspect = plan.outputColumns.length > 1 && step.outputFormat && step.outputFormat.indexOf('|') !== -1;
    if (isMultiAspect) {
      Logger.log('[TaskChain] 📝 Storing multi-aspect metadata: ' + plan.outputColumns.length + ' columns, format: ' + step.outputFormat);
    }
    
    // Calculate output range
    // For single column: G8:G15
    // For multi-column (aspect analysis): G8:J15 (e.g., 4 aspects)
    var rowNums = plan.inputRange.match(/\d+/g);
    var startRow = rowNums ? rowNums[0] : '1';
    var endRow = rowNums && rowNums.length > 1 ? rowNums[rowNums.length - 1] : startRow;
    var firstCol = plan.outputColumns[0];
    var lastCol = plan.outputColumns[plan.outputColumns.length - 1];
    chainState.steps[nextStepIndex].outputRange = firstCol + startRow + ':' + lastCol + endRow;
    
    Logger.log('[TaskChain] Output range for step: ' + chainState.steps[nextStepIndex].outputRange);
    saveChainState(chainState);
    
    Logger.log('[TaskChain] ✅ Step ' + (nextStepIndex + 1) + ' started with job: ' + jobId);
    
  } catch (e) {
    Logger.log('[TaskChain] ❌ Step execution failed: ' + e.message);
    Logger.log('[TaskChain] Failed step index: ' + nextStepIndex);
    Logger.log('[TaskChain] Failed step description: ' + step.description);
    Logger.log('[TaskChain] Failed step inputColumns: ' + JSON.stringify(step.inputColumns));
    Logger.log('[TaskChain] Stack: ' + (e.stack || 'N/A'));
    chainState.steps[nextStepIndex].status = 'failed';
    chainState.steps[nextStepIndex].error = e.message;
    chainState.status = 'failed';
    saveChainState(chainState);
  }
}

/**
 * Handle step completion - called when a job finishes
 * 
 * @param {string} chainId - Chain ID
 * @param {string} jobId - Completed job ID
 * @param {Array} results - Job results
 */
function onStepComplete(chainId, jobId, results) {
  var chainState = getChainState(chainId);
  if (!chainState) return;
  
  // Find the step with this job
  var stepIndex = chainState.steps.findIndex(function(s) {
    return s.jobId === jobId;
  });
  
  if (stepIndex === -1) return;
  
  Logger.log('[TaskChain] Step ' + (stepIndex + 1) + ' completed');
  
  // Update step status
  chainState.steps[stepIndex].status = 'completed';
  chainState.steps[stepIndex].result = results;
  saveChainState(chainState);
  
  // Execute next step
  executeNextStep(chainId, {});
}

/**
 * Record implicit positive feedback for completed analysis steps.
 * A completed chain = user accepted the analysis without re-running.
 * This data feeds into user_analysis_preferences over time.
 */
function _recordImplicitFeedback(chainState) {
  try {
    var analyzeSteps = (chainState.steps || []).filter(function(s) {
      return s.status === 'completed' && 
             (s.action === 'analyze' || s.action === 'chat');
    });
    
    if (analyzeSteps.length === 0) return;
    
    for (var i = 0; i < analyzeSteps.length; i++) {
      var step = analyzeSteps[i];
      var queryText = step.question || step.prompt || step.description || '';
      if (!queryText) continue;
      
      submitAnalysisFeedback(queryText, 'up', null, null);
    }
    
    Logger.log('[TaskChain] Recorded implicit positive feedback for ' + analyzeSteps.length + ' analysis step(s)');
  } catch (e) {
    Logger.log('[TaskChain] Implicit feedback error (non-fatal): ' + e.message);
  }
}

// ============================================
// MODULE LOADED
// ============================================

Logger.log('📦 AgentTaskChain_Execute module loaded');
