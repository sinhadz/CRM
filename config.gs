// ============================================
// CONFIG.GS - System Configuration
// ============================================

const CONFIG = {
  // Your Sheet ID (already correct)
  SPREADSHEET_ID: '1eYJyL-_4aKwcuPIuz_A8VIwJoSCNok9NS6GpvoVozgA',
  
  SHEETS: {
    EMPLOYEES: 'Employees',
    LEADS: 'Leads',
    ATTENDANCE: 'Attendance',
    AUDIT_LOG: 'Audit_Log',
    BREAK_LOG: 'Break_Log',
    INACTIVE_LOG: 'Inactive_Log',
    SESSION_LOG: 'Session_Log',
    EMPLOYEE_AUDIT: 'Employee_Audit'
  },
  
  INACTIVE_TIMEOUT: 10,
  SESSION_TIMEOUT: 480
};

function getConfig() {
  return CONFIG;
}

function getSpreadsheetId() {
  return CONFIG.SPREADSHEET_ID;
}