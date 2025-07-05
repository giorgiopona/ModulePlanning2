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

/**
 * Get Academic Calendar data (Teaching Periods and Start Dates)
 */
function getAcademicCalendar() {
  try {
    Logger.log('=== getAcademicCalendar() started ===');
    
    const debugInfo = {
      configLoaded: typeof CONFIG !== 'undefined',
      spreadsheetId: typeof CONFIG !== 'undefined' ? CONFIG.SPREADSHEET_ID : 'undefined',
      sheetFound: false,
      dataRetrieved: false,
      dataLength: 0,
      filteredDataLength: 0,
      errors: []
    };
    
    Logger.log('Config loaded: ' + debugInfo.configLoaded);
    Logger.log('Spreadsheet ID: ' + debugInfo.spreadsheetId);
    
    if (typeof CONFIG === 'undefined') {
      Logger.log('ERROR: Configuration not loaded');
      return {
        success: false,
        error: 'Configuration not loaded',
        debug: debugInfo
      };
    }
    
    Logger.log('Opening spreadsheet...');
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    Logger.log('Spreadsheet opened successfully');
    
    Logger.log('Looking for Academic Calendar sheet...');
    const calendarSheet = spreadsheet.getSheetByName('Academic Calendar');
    
    if (!calendarSheet) {
      Logger.log('ERROR: Academic Calendar sheet not found');
      Logger.log('Available sheets: ' + spreadsheet.getSheets().map(s => s.getName()).join(', '));
      debugInfo.errors.push('Academic Calendar sheet not found');
      return {
        success: false,
        error: 'Academic Calendar sheet not found',
        debug: debugInfo
      };
    }
    
    Logger.log('Academic Calendar sheet found');
    debugInfo.sheetFound = true;
    
    // Get data from Academic Calendar sheet (A: Teaching Period, B: Start Date)
    Logger.log('Getting data range...');
    const data = calendarSheet.getDataRange().getValues();
    Logger.log('Data retrieved. Rows: ' + data.length);
    debugInfo.dataRetrieved = true;
    debugInfo.dataLength = data.length;
    
    // Log first few rows for debugging
    Logger.log('First 3 rows of data:');
    for (let i = 0; i < Math.min(3, data.length); i++) {
      Logger.log('Row ' + i + ': ' + JSON.stringify(data[i]));
    }
    
    // Filter out empty rows and create period-to-date mapping
    Logger.log('Filtering data...');
    const calendarData = data
      .filter(row => {
        const hasPeriod = row[0] && row[0].toString().trim() !== '';
        const hasDate = row[1] && row[1] instanceof Date;
        
        Logger.log('Row check - Period: "' + row[0] + '" (hasPeriod: ' + hasPeriod + '), Date: ' + row[1] + ' (hasDate: ' + hasDate + ')');
        
        if (!hasPeriod) debugInfo.errors.push('Row missing period: ' + JSON.stringify(row));
        if (!hasDate) debugInfo.errors.push('Row missing valid date: ' + JSON.stringify(row));
        return hasPeriod && hasDate;
      })
      .map(row => ({
        period: row[0].toString().trim(),
        startDate: row[1]
      }));
    
    Logger.log('Filtered data length: ' + calendarData.length);
    Logger.log('Final calendar data: ' + JSON.stringify(calendarData));
    debugInfo.filteredDataLength = calendarData.length;
    
    // Convert Date objects to ISO strings for frontend serialization
    const serializedCalendarData = calendarData.map(item => ({
      period: item.period,
      startDate: item.startDate.toISOString()
    }));
    
    Logger.log('Serialized calendar data: ' + JSON.stringify(serializedCalendarData));
    Logger.log('=== getAcademicCalendar() completed successfully ===');
    
    return {
      success: true,
      calendar: serializedCalendarData,
      debug: debugInfo
    };
  } catch (error) {
    Logger.log('ERROR in getAcademicCalendar(): ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    return {
      success: false,
      error: error.toString(),
      debug: { errors: [error.toString()] }
    };
  }
}

/**
 * Calculate date based on teaching period, week number, and day of week
 */
function calculateDate(period, weekNumber, dayOfWeek) {
  try {
    // Get academic calendar data
    const calendarResponse = getAcademicCalendar();
    if (!calendarResponse.success) {
      return {
        success: false,
        error: 'Failed to get academic calendar: ' + calendarResponse.error
      };
    }
    
    // Find the start date for the given period
    const periodData = calendarResponse.calendar.find(item => item.period === period);
    if (!periodData) {
      return {
        success: false,
        error: 'Teaching period not found in academic calendar: ' + period
      };
    }
    
    // Convert string date back to Date object
    const startDate = new Date(periodData.startDate);
    const weekNum = parseInt(weekNumber);
    
    if (isNaN(weekNum) || weekNum < 1) {
      return {
        success: false,
        error: 'Invalid week number: ' + weekNumber
      };
    }
    
    // Calculate the date for the specified week and day
    // Week 1 starts from the start date
    const targetDate = new Date(startDate);
    targetDate.setDate(startDate.getDate() + (weekNum - 1) * 7);
    
    // Adjust to the specified day of the week
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const targetDayIndex = dayNames.indexOf(dayOfWeek);
    
    if (targetDayIndex === -1) {
      return {
        success: false,
        error: 'Invalid day of week: ' + dayOfWeek
      };
    }
    
    const currentDayIndex = targetDate.getDay();
    const daysToAdd = (targetDayIndex - currentDayIndex + 7) % 7;
    targetDate.setDate(targetDate.getDate() + daysToAdd);
    
    return {
      success: true,
      date: targetDate,
      formattedDate: targetDate.toISOString().split('T')[0] // YYYY-MM-DD format
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Save edited data to the spreadsheet using UID and calculate date if day is changed
 */
function saveEditedDataWithDate(uid, columnIndex, newValue, period, weekNumber) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
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
    const actualRow = targetRow + 2;
    const actualColumn = columnIndex + 1;
    sheet.getRange(actualRow, actualColumn).setValue(newValue);
    let calculatedDate = null;
    if (columnIndex === 13) {
      if (!newValue) {
        // If day is cleared, also clear the date
        sheet.getRange(actualRow, 16).setValue('');
        calculatedDate = '';
      } else if (period && weekNumber) {
        const dateResponse = calculateDate(period, weekNumber, newValue);
        if (dateResponse.success) {
          sheet.getRange(actualRow, 16).setValue(dateResponse.formattedDate);
          calculatedDate = dateResponse.formattedDate;
        }
      }
    }
    return {
      success: true,
      message: 'Data saved successfully for UID: ' + uid,
      calculatedDate: calculatedDate
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Test function to debug date calculation
 */
function testDateCalculation(period, weekNumber, dayOfWeek) {
  try {
    Logger.log('=== testDateCalculation() started ===');
    Logger.log('Input parameters: period=' + period + ', weekNumber=' + weekNumber + ', dayOfWeek=' + dayOfWeek);
    
    const debugInfo = {
      input: { period, weekNumber, dayOfWeek },
      calendarAccess: null,
      dateCalculation: null,
      errors: []
    };
    
    // First test academic calendar access
    try {
      Logger.log('Calling getAcademicCalendar()...');
      const calendarResponse = getAcademicCalendar();
      Logger.log('getAcademicCalendar() response: ' + JSON.stringify(calendarResponse));
      debugInfo.calendarAccess = calendarResponse;
      
      if (!calendarResponse.success) {
        Logger.log('ERROR: Academic calendar access failed: ' + calendarResponse.error);
        debugInfo.errors.push('Academic calendar access failed: ' + calendarResponse.error);
        return {
          success: false,
          error: 'Academic calendar access failed: ' + calendarResponse.error,
          debug: debugInfo
        };
      }
    } catch (calendarError) {
      Logger.log('ERROR: Calendar access exception: ' + calendarError.toString());
      debugInfo.errors.push('Calendar access exception: ' + calendarError.toString());
      return {
        success: false,
        error: 'Calendar access exception: ' + calendarError.toString(),
        debug: debugInfo
      };
    }
    
    // Test date calculation
    try {
      Logger.log('Calling calculateDate()...');
      const dateResponse = calculateDate(period, weekNumber, dayOfWeek);
      Logger.log('calculateDate() response: ' + JSON.stringify(dateResponse));
      debugInfo.dateCalculation = dateResponse;
      
      if (!dateResponse.success) {
        Logger.log('ERROR: Date calculation failed: ' + dateResponse.error);
        debugInfo.errors.push('Date calculation failed: ' + dateResponse.error);
        return {
          success: false,
          error: 'Date calculation failed: ' + dateResponse.error,
          debug: debugInfo
        };
      }
    } catch (dateError) {
      Logger.log('ERROR: Date calculation exception: ' + dateError.toString());
      debugInfo.errors.push('Date calculation exception: ' + dateError.toString());
      return {
        success: false,
        error: 'Date calculation exception: ' + dateError.toString(),
        debug: debugInfo
      };
    }
    
    Logger.log('=== testDateCalculation() completed successfully ===');
    return {
      success: true,
      calendarData: debugInfo.calendarAccess.calendar,
      dateCalculation: debugInfo.dateCalculation,
      debug: debugInfo,
      testResult: {
        period: period,
        weekNumber: weekNumber,
        dayOfWeek: dayOfWeek,
        calculatedDate: debugInfo.dateCalculation.date,
        formattedDate: debugInfo.dateCalculation.formattedDate
      }
    };
  } catch (error) {
    Logger.log('ERROR: Test function exception: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    return {
      success: false,
      error: 'Test function exception: ' + error.toString(),
      debug: { errors: [error.toString()] }
    };
  }
}

/**
 * Batch update staff, room, day, time for multiple records by UID
 * changes: [{ uid, staff, room, day, time, period, week }]
 */
function batchUpdateFields(changes) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getRange(CONFIG.DATA_RANGE).getValues();
    let updated = 0;
    changes.forEach(change => {
      const { uid, staff, room, day, time, period, week } = change;
      let targetRow = -1;
      for (let i = 0; i < data.length; i++) {
        if (data[i][16] && data[i][16].toString().trim() === uid.toString().trim()) {
          targetRow = i;
          break;
        }
      }
      if (targetRow === -1) return;
      const actualRow = targetRow + 2;
      if (staff !== undefined) sheet.getRange(actualRow, 12).setValue(staff);
      if (room !== undefined) sheet.getRange(actualRow, 13).setValue(room);
      if (day !== undefined) sheet.getRange(actualRow, 14).setValue(day);
      if (time !== undefined) sheet.getRange(actualRow, 15).setValue(time);
      // If day is set, calculate and set date; if day is cleared, clear date
      if (day !== undefined) {
        if (!day) {
          sheet.getRange(actualRow, 16).setValue('');
        } else if (period && week) {
          const dateResponse = calculateDate(period, week, day);
          if (dateResponse.success) {
            sheet.getRange(actualRow, 16).setValue(dateResponse.formattedDate);
          }
        }
      }
      updated++;
    });
    return {
      success: true,
      updated: updated,
      total: changes.length
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}
