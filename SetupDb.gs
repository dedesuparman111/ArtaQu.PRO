/* Rahaza Digital Premium Webapps Builder by ddsuparman */

/**
 * Run this function ONCE from the Apps Script Editor to initialize the database.
 * It will create all necessary sheets and insert default data safely.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let message = "Setup Database Completed:\n";

  // 1. Setup USERS Sheet
  let usersSheet = ss.getSheetByName("USERS");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("USERS");
    usersSheet.appendRow(["id", "username", "password", "role", "status"]);
    usersSheet.appendRow(["U-001", "admin", "admin123", "Admin", "Active"]);
    usersSheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
    message += "- USERS sheet created (Default Login: admin / admin123)\n";
  } else {
    message += "- USERS sheet already exists\n";
  }

  // 2. Setup TRANSACTIONS Sheet
  let trxSheet = ss.getSheetByName("TRANSACTIONS");
  if (!trxSheet) {
    trxSheet = ss.insertSheet("TRANSACTIONS");
    trxSheet.appendRow(["id", "date", "type", "category", "amount", "description", "status", "timestamp"]);
    trxSheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
    message += "- TRANSACTIONS sheet created\n";
  } else {
    message += "- TRANSACTIONS sheet already exists\n";
  }

  // 3. Setup INSTALLMENTS Sheet
  let instSheet = ss.getSheetByName("INSTALLMENTS");
  if (!instSheet) {
    instSheet = ss.insertSheet("INSTALLMENTS");
    instSheet.appendRow(["id", "name", "creditor", "total_amount", "paid_amount", "remaining", "start_date", "status", "description", "timestamp"]);
    instSheet.getRange("1:1").setFontWeight("bold").setBackground("#e2e8f0");
    message += "- INSTALLMENTS sheet created\n";
  } else {
    message += "- INSTALLMENTS sheet already exists\n";
  }

  // Finalize UI
  SpreadsheetApp.getUi().alert(message);
}