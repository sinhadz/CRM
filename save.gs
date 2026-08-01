// ============================================
// SAVE.GS - Save Operations
// ============================================


function saveAndNext(employeeId, leadData) {
  var saveResult = saveLead(employeeId, leadData);
  if (!saveResult.success) {
    return saveResult;
  }
  
  var leads = getAssignedLeads(employeeId);
  var nextLead = null;
  
  if (leads.success && leads.data.length > 0) {
    // Find next lead that needs work
    for (var i = 0; i < leads.data.length; i++) {
      var lead = leads.data[i];
      if (!lead.callStatus || lead.callStatus === 'Pending' || 
          (lead.callStatus === 'Follow-up' && !lead.followupDate) ||
          (lead.callStatus.includes('Booked') && !lead.bookingDate)) {
        if (lead.leadNumber !== leadData.leadNumber) {
          nextLead = lead;
          break;
        }
      }
    }
  }
  
  return { 
    success: true, 
    message: 'Lead saved successfully', 
    nextLead: nextLead 
  };
}