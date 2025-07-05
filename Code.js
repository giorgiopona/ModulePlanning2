function doGet() {
  return HtmlService.createHtmlOutputFromFile('index.html')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Debug function to test configuration and spreadsheet access
 */
function debugConfig() {
  try {
    const result = {
      configLoaded: typeof CONFIG !== 'undefined',
      configDetails: null,
      spreadsheetAccess: false,
      sheetAccess: false,
      error: null
    };
    
    if (typeof CONFIG !== 'undefined') {
      result.configDetails = {
        spreadsheetId: CONFIG.SPREADSHEET_ID,
        sheetName: CONFIG.SHEET_NAME,
        dataRange: CONFIG.DATA_RANGE
      };
      
      try {
        const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        result.spreadsheetAccess = true;
        
        const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
        if (sheet) {
          result.sheetAccess = true;
          result.rowCount = sheet.getLastRow();
          result.columnCount = sheet.getLastColumn();
        }
      } catch (spreadsheetError) {
        result.error = 'Spreadsheet access error: ' + spreadsheetError.toString();
      }
    }
    
    return result;
  } catch (error) {
    return {
      configLoaded: false,
      error: 'Debug function error: ' + error.toString()
    };
  }
}

/**
 * Get all data from the Timetable sheet
 */
function getTimetableData() {
  try {
    // Check if CONFIG is available
    if (typeof CONFIG === 'undefined') {
      return {
        success: false,
        error: 'Configuration not loaded'
      };
    }
    
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      return {
        success: false,
        error: 'Sheet not found: ' + CONFIG.SHEET_NAME
      };
    }
    
    const data = sheet.getRange(CONFIG.DATA_RANGE).getValues();
    
    // Filter out empty rows
    const filteredData = data.filter(row => row.some(cell => cell !== ''));
    
    return {
      success: true,
      data: filteredData
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get unique modules from the Timetable sheet
 */
function getUniqueModules() {
  try {
    // Check if CONFIG is available
    if (typeof CONFIG === 'undefined') {
      return {
        success: false,
        error: 'Configuration not loaded'
      };
    }
    
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      return {
        success: false,
        error: 'Sheet not found: ' + CONFIG.SHEET_NAME
      };
    }
    
    const data = sheet.getRange(CONFIG.DATA_RANGE).getValues();
    
    // Debug: Log data processing steps
    const debugInfo = {
      totalRows: data.length,
      nonEmptyRows: 0,
      moduleColumnData: [],
      uniqueModules: []
    };
    
    // Filter out empty rows and get unique modules (column E, index 4)
    const modules = data
      .filter(row => {
        const hasData = row.some(cell => cell !== '');
        if (hasData) debugInfo.nonEmptyRows++;
        return hasData;
      })
      .map(row => {
        const moduleValue = row[4]; // Module column
        debugInfo.moduleColumnData.push({
          value: moduleValue,
          type: typeof moduleValue,
          trimmed: moduleValue ? moduleValue.toString().trim() : ''
        });
        return moduleValue;
      })
      .filter(module => {
        const isValid = module && module.toString().trim() !== '';
        if (isValid) debugInfo.uniqueModules.push(module.toString().trim());
        return isValid;
      });
    
    // Get unique modules - fix the deduplication
    const uniqueModulesSet = new Set();
    modules.forEach(module => {
      if (module && module.toString().trim() !== '') {
        uniqueModulesSet.add(module.toString().trim());
      }
    });
    
    const uniqueModules = Array.from(uniqueModulesSet).sort();
    
    // Add debug info to response
    const response = {
      success: true,
      modules: uniqueModules,
      debug: {
        ...debugInfo,
        uniqueModulesCount: uniqueModules.length,
        uniqueModulesList: uniqueModules
      }
    };
    
    return response;
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

/**
 * Get data for a specific module
 */
function getModuleData(moduleName) {
  try {
    // Check if CONFIG is available
    if (typeof CONFIG === 'undefined') {
      return {
        success: false,
        error: 'Configuration not loaded'
      };
    }
    
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      return {
        success: false,
        error: 'Sheet not found: ' + CONFIG.SHEET_NAME
      };
    }
    
    const data = sheet.getRange(CONFIG.DATA_RANGE).getValues();
    
    // Debug: Log data processing steps
    const debugInfo = {
      requestedModule: moduleName,
      totalRows: data.length,
      matchingRows: 0,
      sampleData: []
    };
    
    // Filter data for the specific module (column E, index 4)
    const moduleData = data.filter(row => {
      const hasData = row.some(cell => cell !== '');
      if (hasData) {
        const rowModule = row[4]; // Module column
        const matches = rowModule && rowModule.toString().trim() === moduleName;
        if (matches) {
          debugInfo.matchingRows++;
          if (debugInfo.sampleData.length < 3) {
            debugInfo.sampleData.push({
              period: row[0] ? row[0].toString() : '',
              week: row[1] ? row[1].toString() : '',
              topic: row[5] ? row[5].toString() : '',
              location: row[6] ? row[6].toString() : '',
              hours: row[8] ? row[8].toString() : '',
              staff: row[11] ? row[11].toString() : ''
            });
          }
        }
        return matches;
      }
      return false;
    });
    
    // Convert all data to strings to ensure serialization works
    const serializedData = moduleData.map(row => 
      row.map(cell => {
        if (cell === null || cell === undefined) return '';
        if (cell instanceof Date) return cell.toISOString();
        return cell.toString();
      })
    );
    
    const response = {
      success: true,
      data: serializedData,
      debug: debugInfo
    };
    
    return response;
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}
