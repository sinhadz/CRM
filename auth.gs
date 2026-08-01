// ============================================
// AUTH.GS - Full Authentication
// ============================================

function authenticateUser(employeeId, pin) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    
    if (!sheet) {
      return { success: false, message: 'Employees sheet not found. Please create it.' };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var empIdCol = headers.indexOf('Employee ID');
    var pinCol = headers.indexOf('PIN');
    var nameCol = headers.indexOf('Name');
    var roleCol = headers.indexOf('Role');
    var statusCol = headers.indexOf('Status');
    
    if (empIdCol === -1 || pinCol === -1) {
      return { success: false, message: 'Invalid sheet structure. Missing Employee ID or PIN columns.' };
    }
    
    for (var i = 1; i < data.length; i++) {
      var empIdFromSheet = String(data[i][empIdCol]).trim();
      var empIdFromInput = String(employeeId).trim();
      
      var isMatch = empIdFromSheet === empIdFromInput || 
                    empIdFromSheet === 'Employee ' + empIdFromInput ||
                    empIdFromInput === empIdFromSheet.replace('Employee ', '');
      
      if (isMatch && String(data[i][pinCol]) === String(pin)) {
        if (data[i][statusCol] && data[i][statusCol].toLowerCase() !== 'active') {
          return { success: false, message: 'Account is inactive.' };
        }
        
        return {
          success: true,
          message: 'Login successful',
          data: {
            employeeId: empIdFromSheet,
            name: data[i][nameCol] || 'Employee',
            role: data[i][roleCol] || 'Employee'
          }
        };
      }
    }
    
    return { success: false, message: 'Invalid Employee ID or PIN' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function logoutUser(employeeId) {
  return { success: true };
}

function getDashboardStats(employeeId) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.LEADS);
    
    if (!sheet) {
      return { success: true, data: getDefaultStats() };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var assignedCol = headers.indexOf('Assigned To');
    var statusCol = headers.indexOf('Call Status');
    var followupCol = headers.indexOf('Follow-up Date');
    var bookingCol = headers.indexOf('Booking Date & Time');
    
    var total = 0;
    var callsMade = 0;
    var pending = 0;
    var disposed = 0;
    var bookedHome = 0;
    var bookedWarehouse = 0;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][assignedCol]) === String(employeeId)) {
        total++;
        var status = String(data[i][statusCol] || '');
        var hasFollowup = data[i][followupCol] ? true : false;
        var hasBooking = data[i][bookingCol] ? true : false;
        
        if (status === 'Follow-up' && hasFollowup) {
          callsMade++;
        } else if (status === 'Booked - Home Inspection' && hasBooking) {
          callsMade++;
          bookedHome++;
        } else if (status === 'Booked - Warehouse Inspection' && hasBooking) {
          callsMade++;
          bookedWarehouse++;
        } else if (['Duplicate', 'Incorrect Number', 'Not Interested', 'Unattended'].indexOf(status) > -1) {
          callsMade++;
          disposed++;
        } else if (status === 'Pending' || !status) {
          pending++;
        } else {
          pending++;
        }
      }
    }
    
    return {
      success: true,
      data: {
        assigned: total,
        callsMade: callsMade,
        pending: pending,
        disposed: disposed,
        bookedHome: bookedHome,
        bookedWarehouse: bookedWarehouse,
        workingTime: 0,
        breakTime: 0,
        inactiveTime: 0,
        netProductiveTime: 0
      }
    };
  } catch (error) {
    return { success: true, data: getDefaultStats() };
  }
}

function getDefaultStats() {
  return {
    assigned: 0,
    callsMade: 0,
    pending: 0,
    disposed: 0,
    bookedHome: 0,
    bookedWarehouse: 0,
    workingTime: 0,
    breakTime: 0,
    inactiveTime: 0,
    netProductiveTime: 0
  };
}