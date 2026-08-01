// ============================================================================
// Logger.gs
// Write4U CRM v2.0
// Initial logging scaffold
// ============================================================================

function logInfo(message,data){Logger.log('[INFO] '+message+(data?' '+JSON.stringify(data):''));}
function logWarn(message,data){Logger.log('[WARN] '+message+(data?' '+JSON.stringify(data):''));}
function logError(message,error){Logger.log('[ERROR] '+message+' '+(error&&error.stack?error.stack:error||''));}
function audit(action,employeeId,details){logInfo('AUDIT',{action:action,employeeId:employeeId,details:details,timestamp:new Date()});}
