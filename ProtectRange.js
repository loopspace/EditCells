function protectRange() {
  // get active sheet
  var ss = SpreadsheetApp.getActive();
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Things get complicated with multiple protections; our goal is to limit a user's editing to one specific range so we need to clear all other protection first
  // Remove all range protections in the spreadsheet that the user has permission to edit.
  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    if (protection.canEdit()) {
      protection.remove();
    }
  }
  
  // Now that we have a blank slate, we'll put in place what we want to achieve.
  // We start by putting protection in place for the entire sheet except for the selected range
  var protection = sheet.protect().setDescription('Protect entire sheet');
  // only editable by the current user
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  // remove all editors (other than owner and current user)
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  // now get the selected range
  var range = sheet.getActiveRange();
  // add this as an exception
  protection.setUnprotectedRanges([range]);
    
  // get the data and number of rows of the range
  var n = range.getNumRows();
  var data = range.getValues();
  
  // the first cell of each row sets the user for that row (and subsequent rows until a new user is specified)
  var user;
  var userrow;
  var row;
  var protectrow;
  for (var i = 0; i < n; i++) {
    if (typeof data[i][0] === 'string' && data[i][0] !== '') {
      user = data[i][0];
      userrow = 0;
    }
    if (typeof user !== 'undefined') {
      userrow++;
      // get current row
      row = range.offset(i,0,1);
      // create a protection object for the row
      protectrow = row.protect();
      protectrow.setDescription('Row ' + userrow + ' for ' + user);
      // add specified user to protection
      protectrow.addEditor(user);
      // make user an editor on the spreadsheet (otherwise they can't edit even their row)
      ss.addEditor(user);
    }
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Protect')
      .addItem('Protect Selection', 'protectRange')
      .addToUi();
}
