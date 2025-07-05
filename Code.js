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
 * Get unique modules with their teaching periods from the Timetable sheet
 */
function getUniqueModules() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getRange(CONFIG.DATA_RANGE).getValues();
    
    // Filter out empty rows and get unique {period, module} pairs (A: period, E: module)
    const pairs = data
      .filter(row => row.some(cell => cell !== ''))
      .map(row => ({ period: row[0], module: row[4] }))
      .filter(pair => pair.period && pair.module && pair.period.toString().trim() !== '' && pair.module.toString().trim() !== '');
    
    // Get unique pairs
    const uniquePairs = [];
    const seen = new Set();
    pairs.forEach(pair => {
      const key = pair.period + '||' + pair.module;
      if (!seen.has(key)) {
        uniquePairs.push(pair);
        seen.add(key);
      }
    });
    
    return {
      success: true,
      modules: uniquePairs
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get data for a specific module and period
 */
function getModuleData(moduleName, period) {
  try {
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
    const debugInfo = {
      requestedModule: moduleName,
      requestedPeriod: period,
      totalRows: data.length,
      matchingRows: 0,
      sampleData: []
    };
    // Filter data for the specific module (E, index 4) and period (A, index 0)
    const moduleData = data.filter(row => {
      const hasData = row.some(cell => cell !== '');
      if (hasData) {
        const rowModule = row[4];
        const rowPeriod = row[0];
        const location = row[6];
        const matches = rowModule && rowModule.toString().trim() === moduleName && rowPeriod && rowPeriod.toString().trim() === period;
        const isNotIndividual = !location || location.toString().trim() !== 'Individual (1-2-1)';
        if (matches && isNotIndividual) {
          debugInfo.matchingRows++;
          if (debugInfo.sampleData.length < 3) {
            debugInfo.sampleData.push({
              period: row[0],
              week: row[1],
              topic: row[5],
              location: row[6],
              hours: row[8],
              staff: row[11]
            });
          }
        }
        return matches && isNotIndividual;
      }
      return false;
    });
    // Convert all data to strings to ensure serialization works
    const serializedData = moduleData.map(row =>
      row.map((cell, cellIndex) => {
        if (cell === null || cell === undefined) return '';
        if (cell instanceof Date) {
          if (cellIndex === 14) {
            const hours = cell.getHours().toString().padStart(2, '0');
            const minutes = cell.getMinutes().toString().padStart(2, '0');
            const seconds = cell.getSeconds().toString().padStart(2, '0');
            return `${hours}:${minutes}:${seconds}`;
          }
          return cell.toISOString();
        }
        return cell.toString();
      })
    );
    // Group data by groups (column H, index 7)
    const groupedData = {};
    const allGroups = new Set();
    serializedData.forEach(row => {
      const group = row[7] || 'No Group';
      allGroups.add(group);
      if (!groupedData[group]) {
        groupedData[group] = [];
      }
      groupedData[group].push(row);
    });
    // Sort groups and sessions within each group
    const sortedGroups = Array.from(allGroups).sort((a, b) => {
      const numA = parseInt(a.toString().match(/\d+/)) || 0;
      const numB = parseInt(b.toString().match(/\d+/)) || 0;
      if (numA !== numB) {
        return numA - numB;
      }
      return a.toString().localeCompare(b.toString());
    });
    const organizedData = {};
    sortedGroups.forEach(group => {
      organizedData[group] = groupedData[group].sort((a, b) => {
        const weekA = parseInt(a[1]) || 0;
        const weekB = parseInt(b[1]) || 0;
        if (weekA !== weekB) return weekA - weekB;
        const periodA = a[0] || '';
        const periodB = b[0] || '';
        return periodA.localeCompare(periodB);
      });
    });
    const response = {
      success: true,
      data: serializedData,
      groupedData: organizedData,
      groups: sortedGroups,
      debug: {
        ...debugInfo,
        totalGroups: sortedGroups.length,
        groupsList: sortedGroups
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
 * Get staff list from HR sheet
 */
function getStaffList() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const hrSheet = spreadsheet.getSheetByName('HR');
    
    if (!hrSheet) {
      return {
        success: false,
        error: 'HR sheet not found'
      };
    }
    
    const staffData = hrSheet.getRange('A2:A').getValues();
    const staffList = staffData
      .filter(row => row[0] && row[0].toString().trim() !== '')
      .map(row => row[0].toString().trim());
    
    return {
      success: true,
      staff: staffList
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get room list from Facility sheet
 */
function getRoomList() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const facilitySheet = spreadsheet.getSheetByName('Facility');
    
    if (!facilitySheet) {
      return {
        success: false,
        error: 'Facility sheet not found'
      };
    }
    
    const roomData = facilitySheet.getRange('A2:A').getValues();
    const roomList = roomData
      .filter(row => row[0] && row[0].toString().trim() !== '')
      .map(row => row[0].toString().trim());
    
    return {
      success: true,
      rooms: roomList
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Save edited data to the spreadsheet using UID
 */
function saveEditedData(uid, columnIndex, newValue) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    
    // Find the row with the matching UID (column Q, index 16)
    const data = sheet.getRange(CONFIG.DATA_RANGE).getValues();
    let targetRow = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][16] && data[i][16].toString().trim() === uid.toString().trim()) {
        targetRow = i;
        break;
      }
    }
    
    if (targetRow === -1) {
      return {
        success: false,
        error: 'Record with UID ' + uid + ' not found'
      };
    }
    
    // Convert from 0-based index to 1-based and add header row offset
    const actualRow = targetRow + 2; // +2 because we start from A2 and targetRow is 0-based
    const actualColumn = columnIndex + 1; // +1 because columnIndex is 0-based
    
    sheet.getRange(actualRow, actualColumn).setValue(newValue);
    
    return {
      success: true,
      message: 'Data saved successfully for UID: ' + uid
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get unique periods from the Timetable sheet
 */
function getUniquePeriods() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      return {
        success: false,
        error: 'Sheet not found: ' + CONFIG.SHEET_NAME
      };
    }
    
    const data = sheet.getRange(CONFIG.DATA_RANGE).getValues();
    
    // Get unique periods (column A, index 0)
    const periods = data
      .filter(row => row.some(cell => cell !== ''))
      .map(row => row[0]) // Period column
      .filter(period => period && period.toString().trim() !== '');
    
    // Get unique periods and sort them
    const uniquePeriods = [...new Set(periods)].sort();
    
    return {
      success: true,
      periods: uniquePeriods
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}
