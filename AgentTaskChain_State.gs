/**
 * @file AgentTaskChain_State.gs
 * @version 2.0.0
 * @updated 2026-02-10
 * @author AISheeter Team
 * 
 * ============================================
 * AGENT TASK CHAIN - State Management
 * ============================================
 * 
 * Manages chain state persistence and retrieval:
 * - Save/load chain state to UserProperties
 * - Automatic cleanup of old chain data
 * - Job status checking and chain advancement
 * - Property storage quota management
 * 
 * Part of the AgentTaskChain module suite:
 * - AgentTaskChain_Parse.gs    (parsing & suggestions)
 * - AgentTaskChain_Execute.gs  (execution flow)
 * - AgentTaskChain_Plan.gs     (plan building)
 * - AgentTaskChain_State.gs    (this file)
 * - AgentTaskChain_Analyze.gs  (analyze/chat steps)
 * 
 * Depends on:
 * - AgentTaskChain_Execute.gs (executeNextStep - for chain advancement)
 * - AgentTaskChain_Plan.gs (generateSmartHeader - for result writing)
 * - Jobs.gs (getJobStatus, writeJobResults)
 */

// ============================================
// CHAIN STATE MANAGEMENT
// ============================================

/**
 * Save chain state to user properties with retry logic
 * Automatically cleans up old chains to prevent quota overflow.
 * 
 * @param {Object} chainState - Chain state object
 * @param {number} [retries] - Number of retries (default 3)
 */
function saveChainState(chainState, retries) {
  retries = retries || 3;
  var lastError = null;
  
  for (var attempt = 0; attempt < retries; attempt++) {
    try {
      var props = PropertiesService.getUserProperties();
      var key = 'chain_' + chainState.chainId;
      
      // Clean up debug info before saving (reduce size)
      delete chainState._debug;
      for (var i = 0; i < chainState.steps.length; i++) {
        delete chainState.steps[i]._lastJobStatus;
        delete chainState.steps[i]._lastJobProgress;
      }
      
      // On first attempt, run lightweight cleanup to prevent quota overflow
      if (attempt === 0) {
        _cleanupOldChains(props);
      }
      
      props.setProperty(key, JSON.stringify(chainState));
      
      // Also save to list of active chains (only on first save)
      var activeChains = JSON.parse(props.getProperty('activeChains') || '[]');
      if (activeChains.indexOf(chainState.chainId) === -1) {
        activeChains.push(chainState.chainId);
        // Keep only last 5 chains (reduced from 10 to save space)
        if (activeChains.length > 5) {
          // Delete properties of chains being removed from the list
          var removed = activeChains.slice(0, activeChains.length - 5);
          removed.forEach(function(oldId) {
            props.deleteProperty('chain_' + oldId);
          });
          activeChains = activeChains.slice(-5);
        }
        props.setProperty('activeChains', JSON.stringify(activeChains));
      }
      
      return; // Success
      
    } catch (e) {
      lastError = e;
      Logger.log('[TaskChain] Save attempt ' + (attempt + 1) + ' failed: ' + e.message);
      
      // If quota exceeded, force a deep cleanup and retry
      if (e.message && e.message.indexOf('quota') !== -1) {
        Logger.log('[TaskChain] Quota exceeded — running deep cleanup...');
        _deepCleanupProperties();
      }
      
      // Wait before retry (exponential backoff)
      if (attempt < retries - 1) {
        Utilities.sleep(100 * Math.pow(2, attempt)); // 100ms, 200ms, 400ms
      }
    }
  }
  
  // All retries failed
  Logger.log('[TaskChain] ❌ Failed to save chain state after ' + retries + ' attempts: ' + lastError.message);
  throw lastError;
}

/**
 * Lightweight cleanup: remove completed chain properties not in active list
 * Called on every save to prevent gradual accumulation.
 */
function _cleanupOldChains(props) {
  try {
    var allKeys = props.getKeys();
    var activeChains = JSON.parse(props.getProperty('activeChains') || '[]');
    var activeSet = {};
    activeChains.forEach(function(id) { activeSet['chain_' + id] = true; });
    
    var removed = 0;
    allKeys.forEach(function(key) {
      if (key.startsWith('chain_') && key !== 'activeChains' && !activeSet[key]) {
        props.deleteProperty(key);
        removed++;
      }
    });
    
    if (removed > 0) {
      Logger.log('[TaskChain] Cleanup: removed ' + removed + ' orphaned chain properties');
    }
  } catch (e) {
    Logger.log('[TaskChain] Cleanup error (non-fatal): ' + e.message);
  }
}

/**
 * Deep cleanup: aggressively free property storage when quota is hit.
 * Removes all chain data, old analytics, and trims history.
 */
function _deepCleanupProperties() {
  try {
    var props = PropertiesService.getUserProperties();
    var allKeys = props.getKeys();
    var removedCount = 0;
    
    // Remove ALL chain properties
    allKeys.forEach(function(key) {
      if (key.startsWith('chain_') || key === 'activeChains') {
        props.deleteProperty(key);
        removedCount++;
      }
    });
    
    // Remove analytics
    props.deleteProperty('pendingAnalytics');
    removedCount++;
    
    // Trim task history to last 5
    try {
      var history = JSON.parse(props.getProperty('agentRecentTasks') || '[]');
      if (history.length > 5) {
        props.setProperty('agentRecentTasks', JSON.stringify(history.slice(0, 5)));
      }
    } catch (e2) { /* ignore */ }
    
    Logger.log('[TaskChain] Deep cleanup: removed ' + removedCount + ' properties');
  } catch (e) {
    Logger.log('[TaskChain] Deep cleanup error: ' + e.message);
  }
}

/**
 * Manual cleanup function — can be called from frontend or script editor
 * to free up property storage quota.
 * @return {Object} Cleanup summary
 */
function cleanupPropertyStorage() {
  var props = PropertiesService.getUserProperties();
  var allKeys = props.getKeys();
  var before = allKeys.length;
  
  _deepCleanupProperties();
  
  var after = props.getKeys().length;
  var result = {
    before: before,
    after: after,
    removed: before - after
  };
  Logger.log('[Cleanup] Result: ' + JSON.stringify(result));
  return result;
}

/**
 * Get chain state from user properties
 * Also checks and updates job status for running steps (enables progress tracking)
 * 
 * PERFORMANCE OPTIMIZATION (2026-01-21):
 * - Added minimum interval between job status checks (2 seconds)
 * - Prevents excessive API calls when frontend polls frequently
 * - Saves ~200-300ms per redundant check
 * 
 * @param {string} chainId - Chain ID
 * @return {Object|null} Chain state or null
 */
function getChainState(chainId) {
  var props = PropertiesService.getUserProperties();
  var key = 'chain_' + chainId;
  var data = props.getProperty(key);
  
  if (!data) return null;
  
  var chainState = JSON.parse(data);
  
  // If chain is running, check status of running steps
  if (chainState.status === 'running') {
    try {
      // PERFORMANCE: Only check job status if enough time has passed
      // This prevents excessive API calls when frontend polls frequently
      var now = new Date().getTime();
      var lastCheck = chainState._lastStatusCheck || 0;
      var MIN_CHECK_INTERVAL = 2000; // 2 seconds minimum between checks
      
      if (now - lastCheck >= MIN_CHECK_INTERVAL) {
        chainState._lastStatusCheck = now;
        var result = checkRunningStepsStatus(chainState);
        // Only save if modified AND executeNextStep didn't already save
        if (result.modified && !result.savedByNextStep) {
          saveChainState(chainState);
        }
      } else {
        // Skip check - too soon since last check
        Logger.log('[TaskChain] Skipping job status check (last check ' + (now - lastCheck) + 'ms ago)');
      }
    } catch (e) {
      // Add debug info to chain state for frontend visibility
      chainState._debug = chainState._debug || [];
      chainState._debug.push({
        time: new Date().toISOString(),
        error: 'checkRunningStepsStatus failed: ' + e.message
      });
      // Keep only last 5 debug entries
      if (chainState._debug.length > 5) {
        chainState._debug = chainState._debug.slice(-5);
      }
    }
  }
  
  // NOTE: Job status debug info is now set by checkRunningStepsStatus
  // to avoid duplicate API calls. Don't call getJobStatus here again!
  
  return chainState;
}

/**
 * Check status of running steps and advance chain if needed
 * This enables the polling mechanism to track progress
 * 
 * @param {Object} chainState - Chain state
 * @return {Object} { modified: boolean, savedByNextStep: boolean }
 */
function checkRunningStepsStatus(chainState) {
  var modified = false;
  var savedByNextStep = false; // Prevent double-save when executeNextStep already saved
  
  for (var i = 0; i < chainState.steps.length; i++) {
    var step = chainState.steps[i];
    
    if (step.status === 'running') {
      // Check if step has a job ID
      if (!step.jobId) {
        Logger.log('[TaskChain] ⚠️ Step ' + (i + 1) + ' is running but has NO jobId!');
        // This means the job creation failed silently - mark as failed
        step.status = 'failed';
        step.error = 'No job was created for this step';
        chainState.status = 'failed';
        modified = true;
        continue;
      }
      
      // Check job status (only ONE API call per step per poll!)
      try {
        Logger.log('[TaskChain] Checking job status for step ' + (i + 1) + ': ' + step.jobId);
        Logger.log('[TaskChain] Step outputRange: ' + step.outputRange);
        var jobStatus = getJobStatus(step.jobId);
        Logger.log('[TaskChain] Job results count: ' + (jobStatus.results ? jobStatus.results.length : 0));
        
        // Store for frontend visibility (this was previously done in getChainState too - now only here)
        step._lastJobStatus = jobStatus ? jobStatus.status : 'unknown';
        step._lastJobProgress = jobStatus ? (jobStatus.progress || 0) : 0;
        
        Logger.log('[TaskChain] Job ' + step.jobId + ' status: ' + JSON.stringify({
          status: jobStatus && jobStatus.status,
          progress: jobStatus && jobStatus.progress,
          error: jobStatus && jobStatus.error
        }));
        
        if (jobStatus && jobStatus.status === 'completed') {
          Logger.log('[TaskChain] ✅ Step ' + (i + 1) + ' job completed!');
          
          // CRITICAL: Write job results to the sheet (this was missing!)
          // Without this, jobs complete but no data is written!
          var results = jobStatus.results || [];
          if (results.length > 0 && step.outputRange) {
            Logger.log('[TaskChain] 📝 Writing ' + results.length + ' results to ' + step.outputRange);
            
            // Extract output column and start row from outputRange (e.g., "G8:G15")
            var rangeMatch = step.outputRange.match(/([A-Z]+)(\d+)/i);
            var outputColumn = rangeMatch ? rangeMatch[1] : 'G';
            var startRow = rangeMatch ? parseInt(rangeMatch[2]) : 2;
            
            try {
              var sheet = SpreadsheetApp.getActiveSheet();
              
              // CRITICAL: Check if multi-aspect (split results into multiple columns)
              var isMultiAspect = step.outputColumns && step.outputColumns.length > 1 && 
                                  step.outputFormat && step.outputFormat.indexOf('|') !== -1;
              
              if (isMultiAspect) {
                Logger.log('[TaskChain] 🔀 Multi-aspect detected: ' + step.outputColumns.length + ' aspects');
                Logger.log('[TaskChain] 🔀 Output format: ' + step.outputFormat);
                Logger.log('[TaskChain] 🔀 Columns: ' + step.outputColumns.join(', '));
                
                // Split results
                var aspectCount = step.outputColumns.length;
                var columnResults = [];
                for (var colIdx = 0; colIdx < aspectCount; colIdx++) {
                  columnResults.push([]);
                }
                
                results.forEach(function(result, rowIdx) {
                  var output = String(result.output || '').trim();
                  var parts = output.split('|').map(function(p) { return p.trim(); });
                  
                  if (rowIdx < 3) {
                    Logger.log('[TaskChain] 🔀 Row ' + rowIdx + ': "' + output + '" → [' + parts.join(', ') + ']');
                  }
                  
                  for (var colIdx = 0; colIdx < aspectCount; colIdx++) {
                    var value = parts[colIdx] || '';
                    columnResults[colIdx].push({
                      index: result.index,
                      output: value,
                      input: result.input
                    });
                  }
                });
                
                // Write each aspect to its column
                step.outputColumns.forEach(function(col, colIdx) {
                  Logger.log('[TaskChain] 📝 Writing aspect ' + (colIdx + 1) + ' to column ' + col);
                  
                  // Write header if requested
                  if (chainState.addHeaders) {
                    var headerRow = startRow - 1;
                    if (headerRow >= 1) {
                      var aspects = step.outputFormat.split('|').map(function(a) { return a.trim(); });
                      var headerText = aspects[colIdx] || ('Aspect ' + (colIdx + 1));
                      var colNumber = col.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
                      sheet.getRange(headerRow, colNumber).setValue(headerText);
                      Logger.log('[TaskChain] 📝 Header written to ' + col + headerRow + ': "' + headerText + '"');
                    }
                  }
                  
                  var writeResult = writeJobResults(columnResults[colIdx], col, startRow);
                  Logger.log('[TaskChain] 📝 Column ' + col + ' write result: ' + writeResult.successCount + ' success');
                });
                
                step.writeResult = { successCount: results.length, errorCount: 0 };
              } else {
                // Single-column write
                // Write column header if requested (row above data)
                if (chainState.addHeaders && outputColumn) {
                  var headerRow = startRow - 1;
                  if (headerRow >= 1) {
                    // Generate smart header from description (NOT outputFormat!)
                    var headerText = generateSmartHeader(step.description, step.action);
                    // Truncate if too long
                    if (headerText.length > 40) {
                      headerText = headerText.substring(0, 37) + '...';
                    }
                    // Convert column letter to column number
                    var colNumber = outputColumn.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
                    sheet.getRange(headerRow, colNumber).setValue(headerText);
                    Logger.log('[TaskChain] 📝 Header written to ' + outputColumn + headerRow + ': "' + headerText + '"');
                  }
                }
                
                var writeResult = writeJobResults(results, outputColumn, startRow);
                Logger.log('[TaskChain] 📝 Write result: ' + writeResult.successCount + ' success, ' + writeResult.errorCount + ' errors');
                step.writeResult = writeResult;
              }
            } catch (writeError) {
              Logger.log('[TaskChain] ❌ Failed to write results: ' + writeError.message);
              step.writeError = writeError.message;
            }
          } else {
            Logger.log('[TaskChain] ⚠️ No results to write or no outputRange! results=' + results.length + ', outputRange=' + step.outputRange);
          }
          
          step.status = 'completed';
          step.completedAt = new Date().toISOString();
          modified = true;
          
          // Trigger next step (this saves chain state internally)
          executeNextStep(chainState.chainId, chainState);
          savedByNextStep = true;
          
        } else if (jobStatus && jobStatus.status === 'failed') {
          Logger.log('[TaskChain] ❌ Step ' + (i + 1) + ' job failed: ' + (jobStatus.error || 'Unknown error'));
          step.status = 'failed';
          step.error = jobStatus.error || 'Job failed';
          chainState.status = 'failed';
          modified = true;
        } else if (jobStatus && (jobStatus.status === 'queued' || jobStatus.status === 'processing')) {
          // Job is still being processed - check if worker is running
          Logger.log('[TaskChain] Job ' + step.jobId + ' still ' + jobStatus.status + ', waiting...');
        } else {
          Logger.log('[TaskChain] ⚠️ Unknown job status: ' + (jobStatus && jobStatus.status || 'null'));
        }
        
      } catch (e) {
        Logger.log('[TaskChain] Error checking job status: ' + e.message);
        Logger.log('[TaskChain] Error stack: ' + (e.stack || 'N/A'));
        
        // Store error in step for frontend visibility
        step._lastJobStatus = 'error: ' + e.message;
        
        // If quota error, stop the chain to prevent burning more quota
        if (e.message && (e.message.indexOf('too many times') > -1 || 
                          e.message.indexOf('urlfetch') > -1 ||
                          e.message.indexOf('quota') > -1)) {
          Logger.log('[TaskChain] ❌ QUOTA ERROR - stopping chain to prevent further API calls');
          step.status = 'failed';
          step.error = 'Google API quota exceeded. Please try again later.';
          chainState.status = 'failed';
          chainState._quotaError = true;
          modified = true;
        }
      }
    }
  }
  
  return { modified: modified, savedByNextStep: savedByNextStep };
}

/**
 * Get all active chains for current user
 * 
 * @return {Array} List of chain states
 */
function getActiveChains() {
  var props = PropertiesService.getUserProperties();
  var activeChains = JSON.parse(props.getProperty('activeChains') || '[]');
  
  return activeChains.map(function(chainId) {
    return getChainState(chainId);
  }).filter(function(state) {
    return state !== null;
  });
}

/**
 * Clear a chain state
 * 
 * @param {string} chainId - Chain ID
 */
function clearChainState(chainId) {
  var props = PropertiesService.getUserProperties();
  props.deleteProperty('chain_' + chainId);
  
  var activeChains = JSON.parse(props.getProperty('activeChains') || '[]');
  var index = activeChains.indexOf(chainId);
  if (index !== -1) {
    activeChains.splice(index, 1);
    props.setProperty('activeChains', JSON.stringify(activeChains));
  }
}

// ============================================
// MODULE LOADED
// ============================================

Logger.log('📦 AgentTaskChain_State module loaded');
