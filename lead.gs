// ============================================
// LEAD.GS - Lead Management
// ============================================

function getAssignedLeads(employeeId) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.LEADS);
    
    if (!sheet) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var leads = [];
    var assignedCol = headers.indexOf('Assigned To');
    var leadNumCol = headers.indexOf('Lead Number');
    var nameCol = headers.indexOf('Customer Name');
    var phoneCol = headers.indexOf('Phone');
    var vehicleCol = headers.indexOf('Vehicle');
    var statusCol = headers.indexOf('Call Status');
    var remarksCol = headers.indexOf('Remarks');
    var followupCol = headers.indexOf('Follow-up Date');
    var bookingCol = headers.indexOf('Booking Date & Time');
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][assignedCol]) === String(employeeId)) {
        leads.push({
          leadNumber: data[i][leadNumCol] || '',
          customerName: data[i][nameCol] || '',
          phone: data[i][phoneCol] || '',
          vehicle: data[i][vehicleCol] || '',
          callStatus: data[i][statusCol] || '',
          remarks: data[i][remarksCol] || '',
          followupDate: data[i][followupCol] || '',
          bookingDate: data[i][bookingCol] || ''
        });
      }
    }
    
    return { success: true, data: leads };
  } catch (error) {
    return { success: true, data: [] };
  }
}

function getLeadDetails(employeeId, leadNumber) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.LEADS);
    
    if (!sheet) {
      return { success: false, message: 'Leads sheet not found' };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var leadNumCol = headers.indexOf('Lead Number');
    var assignedCol = headers.indexOf('Assigned To');
    var nameCol = headers.indexOf('Customer Name');
    var phoneCol = headers.indexOf('Phone');
    var addressCol = headers.indexOf('Address');
    var vehicleCol = headers.indexOf('Vehicle');
    var statusCol = headers.indexOf('Call Status');
    var remarksCol = headers.indexOf('Remarks');
    var followupCol = headers.indexOf('Follow-up Date');
    var bookingCol = headers.indexOf('Booking Date & Time');
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][leadNumCol]) === String(leadNumber)) {
        if (String(data[i][assignedCol]) !== String(employeeId)) {
          return { success: false, message: 'Unauthorized access' };
        }
        
        return {
          success: true,
          data: {
            leadNumber: data[i][leadNumCol] || '',
            customerName: data[i][nameCol] || '',
            phone: data[i][phoneCol] || '',
            address: data[i][addressCol] || '',
            vehicle: data[i][vehicleCol] || '',
            callStatus: data[i][statusCol] || '',
            remarks: data[i][remarksCol] || '',
            followupDate: data[i][followupCol] || '',
            bookingDate: data[i][bookingCol] || '',
            assignedTo: data[i][assignedCol] || ''
          }
        };
      }
    }
    
    return { success: false, message: 'Lead not found' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}