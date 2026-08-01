// ============================================================================
// Security.gs
// Initial security scaffold
// ============================================================================

function sanitizeText(value){return String(value||'').replace(/[<>]/g,'').trim();}
function validateEmployeeId(id){return /^[A-Za-z0-9 _-]{1,50}$/.test(String(id||''));}
function validatePin(pin){return /^\d{4,8}$/.test(String(pin||''));}
function requireRole(currentRole,allowed){if(allowed.indexOf(currentRole)===-1)throw new Error('Unauthorized');return true;}
