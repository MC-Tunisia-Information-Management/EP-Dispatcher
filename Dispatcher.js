function distributeEPs() {
  var mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var crmSheet = mainSpreadsheet.getSheetByName("CRM");
  var contactListSheet = mainSpreadsheet.getSheetByName("Contact List oGV");

  if (!crmSheet || !contactListSheet) {
    Logger.log("Required sheets not found.");
    return;
  }

  var managersDataRange = contactListSheet.getRange("D5:E"); // Update range if needed
  var managersDataValues = managersDataRange.getValues();

  var managers = {};
  managersDataValues.forEach(function (row) {
    var manager = row[0];
    if (manager && !managers[manager]) {
      managers[manager] = { email: row[1], assignedEPs: 0 };
    }
  });

  var dataRange = crmSheet.getRange("B6:G");
  var data = dataRange.getValues();

  for (var i = 0; i < data.length; i++) {
    var epName = data[i][0];
    var currentManager = data[i][5];

    if (!epName || currentManager) {
      continue;
    }

    var assignedManager = null;
    for (var manager in managers) {
      if (managers[manager].assignedEPs < 2) {
        assignedManager = manager;
        break;
      }
    }

    if (!assignedManager) {
      // Find the manager with the least assigned EPs
      var leastAssignedManager = null;
      var leastAssignedEPs = Infinity;
      for (var manager in managers) {
        if (managers[manager].assignedEPs < leastAssignedEPs) {
          leastAssignedManager = manager;
          leastAssignedEPs = managers[manager].assignedEPs;
        }
      }
      assignedManager = leastAssignedManager;
    }

    if (assignedManager) {
      data[i][5] = assignedManager;
      managers[assignedManager].assignedEPs++;
      sendEmailNotification(managers[assignedManager].email, data[i]);
      Logger.log(`Assigned EP ${epName} to Manager: ${assignedManager}`);
    }
  }

  dataRange.setValues(data);
  Logger.log("Script executed successfully.");
}

function sendEmailNotification(managerEmail, epData) {
  var subject = "EP Assignment Notification";
  var body =
    "Hello Manager,\n\nYou have been assigned this EP:\n\n" +
    "EP Name: " +
    epData[0] +
    "\n" +
    "EP Email: " +
    epData[1] +
    "\n" +
    "EP Phone Number: " +
    epData[2] +
    "\n" +
    // Include other EP data fields as needed
    "\n\nPlease update your tracking tool accordingly.\n\nRegards,\nYour Allocation System";

  // Send notification email
  GmailApp.sendEmail(managerEmail, subject, body);
}
