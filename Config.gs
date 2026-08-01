// ============================================================================
// Config.gs
// Write4U CRM v2.0
// ============================================================================

const CONFIG={
SPREADSHEET_ID:'REPLACE_WITH_SPREADSHEET_ID',
SHEETS:{EMPLOYEES:'Employees',LEADS:'Leads',TIME_LOG:'Time Log',AUDIT:'Audit Log'},
TIMEOUTS:{AUTO_BREAK_MINUTES:10,INACTIVITY_MINUTES:10,LEAD_LOCK_MINUTES:15},
CACHE:{DEFAULT_SECONDS:300},
FEATURES:{AUTO_BREAK:true,LEAD_LOCKING:true,AUDIT_LOGGING:true}
};
