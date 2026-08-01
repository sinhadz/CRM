// ============================================
// TIME.GS - Time Tracking (Simplified)
// ============================================


function endBreak(employeeId) {
  return { success: true, message: 'Break ended' };
}

function logInactive(employeeId, startTime, endTime) {
  return { success: true };
}