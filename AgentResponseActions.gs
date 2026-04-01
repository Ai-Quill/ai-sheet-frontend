/**
 * @file AgentResponseActions.gs
 * @version 1.0.0
 * @updated 2026-01-24
 * 
 * ============================================
 * AGENT RESPONSE ACTIONS
 * ============================================
 * 
 * Handles user actions on agent chat responses:
 * - Insert as cell note
 * - Insert in new column
 * - Insert in new sheet
 */

/**
 * Insert agent response into sheet
 * 
 * @param {string} content - The response text to insert
 * @param {string} action - 'note', 'column', or 'sheet'
 * @return {Object} { success: boolean, message: string }
 */
function insertAgentResponse(content, action) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeCell = sheet.getActiveCell();
    
    if (action === 'note') {
      // Insert as cell note (recommended - non-intrusive)
      if (!activeCell) {
        return { success: false, message: 'No cell selected' };
      }
      
      activeCell.setNote(content);
      
      // Visual feedback - highlight the cell briefly
      const currentBackground = activeCell.getBackground();
      activeCell.setBackground('#fef3c7'); // Amber highlight
      
      // Reset after 1 second
      Utilities.sleep(1000);
      activeCell.setBackground(currentBackground);
      
      return { 
        success: true, 
        message: `Inserted as note in ${activeCell.getA1Notation()}`
      };
      
    } else if (action === 'column') {
      // Insert in new column (creates "AI Summary" column)
      const lastColumn = sheet.getLastColumn();
      const newColumn = lastColumn + 1;
      
      // Set header
      const headerCell = sheet.getRange(1, newColumn);
      headerCell.setValue('AI Summary');
      headerCell.setFontWeight('bold');
      headerCell.setBackground('#dbeafe');
      
      // Insert content in first row below header
      const contentCell = sheet.getRange(2, newColumn);
      contentCell.setValue(content);
      contentCell.setWrap(true);
      
      // Auto-resize column width
      sheet.autoResizeColumn(newColumn);
      
      // Select the new cell
      sheet.setActiveRange(contentCell);
      
      return { 
        success: true, 
        message: `Inserted in new column ${String.fromCharCode(64 + newColumn)}`
      };
      
    } else if (action === 'sheet') {
      // Insert in new sheet (creates "Analysis" sheet)
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Generate unique sheet name
      let sheetName = 'AI Analysis';
      let counter = 1;
      while (spreadsheet.getSheetByName(sheetName)) {
        counter++;
        sheetName = `AI Analysis ${counter}`;
      }
      
      // Create new sheet
      const newSheet = spreadsheet.insertSheet(sheetName);
      
      // Format header
      const headerCell = newSheet.getRange(1, 1);
      headerCell.setValue('Analysis Report');
      headerCell.setFontSize(14);
      headerCell.setFontWeight('bold');
      headerCell.setBackground('#dbeafe');
      
      // Add timestamp
      const timestampCell = newSheet.getRange(2, 1);
      timestampCell.setValue(`Generated: ${new Date().toLocaleString()}`);
      timestampCell.setFontSize(9);
      timestampCell.setFontColor('#6b7280');
      
      // Insert content starting from row 4
      const contentCell = newSheet.getRange(4, 1);
      contentCell.setValue(content);
      contentCell.setWrap(true);
      contentCell.setVerticalAlignment('top');
      
      // Set column width for readability
      newSheet.setColumnWidth(1, 600);
      
      // Activate the new sheet
      spreadsheet.setActiveSheet(newSheet);
      
      return { 
        success: true, 
        message: `Created new sheet "${sheetName}"`
      };
      
    } else {
      return { success: false, message: 'Invalid action' };
    }
    
  } catch (error) {
    Logger.log('[InsertResponse] Error: ' + error.message);
    return { 
      success: false, 
      message: 'Failed to insert: ' + error.message 
    };
  }
}
