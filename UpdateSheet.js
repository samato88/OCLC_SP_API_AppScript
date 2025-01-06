function updateSheet(form, startingRow, numRows, dataSheet, hathiColumn,hathiIdColumn, hathiTitleColumn, iaColumn, sppColumn, retainersColumn, currentOCLCColumn, usHoldingsColumn,mergedOCLCColumn, worldcatTitleColumn, columnHeaders) {
  // update sheet, would be faster to do as a multidimensional array
  try {
    var lock = LockService.getScriptLock();
    try { 
    lock.tryLock(30000)  //This method has no effect if the lock has already been acquired.https://developers.google.com/apps-script/reference/lock/lock

    if (form.hathi) {
      updateSheetColumn(startingRow, numRows, hathiColumn, "E", dataSheet) ; 
      updateSheetColumn(startingRow, numRows, hathiIdColumn, "I", dataSheet) ; 
      updateSheetColumn(startingRow, numRows, hathiTitleColumn, "J", dataSheet) ; 
    }
    if (form.ia)   {updateSheetColumn(startingRow, numRows, iaColumn, "F", dataSheet) ; }

  if (form.worldcatretentions) { // this is the checkbox for SPP
      updateSheetColumn(startingRow, numRows, sppColumn,        "C", dataSheet) ; 
      updateSheetColumn(startingRow, numRows, retainersColumn,   "D", dataSheet) ; 
      updateSheetColumn(startingRow, numRows, currentOCLCColumn, "B", dataSheet) ; 
      if (form.WCHoldings) {
        updateSheetColumn(startingRow, numRows, usHoldingsColumn, "G", dataSheet) ;
      }
      if (form.WCData) { 
        updateSheetColumn(startingRow, numRows, mergedOCLCColumn, "H", dataSheet) ;
        updateSheetColumn(startingRow, numRows, worldcatTitleColumn, "K", dataSheet) ; }
    } // end if form.worldcatretentions
  
    dataSheet.getRange("B1:K1").setValues(columnHeaders); // set column headers
    dataSheet.getRange("B1:K1").setFontWeight("bold");
    dataSheet.getRange("B1:K1").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);


    lock.releaseLock();
    } catch {
          ui.alert("Failed to sheet lock - reload sheet and try again.") // not sure this is a good idea
          return ;
    }
  } // end try getting a lock and updating sheet
  catch { // catch failed lock
    ui.alert("Failed to lock sheet - reload sheet and try again.") // slightly different error message for debugging where failed
    return ;
  }

 /* removing this Jan 2025 since some universities won't approve add ons that ask for email scope
  if (form.sendEmail) {// send plain generic email if requested ( html alternative https://blog.gsmart.in/google-apps-script-send-html-email/)
    var donetime = new Date();
    var emailAddress = form.emailAddress ;
    var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;  
    if(emailPattern.test(emailAddress) == false) {
      //Logger.log("Invalid Email Address Entered. " + emailAddress);
      ui.alert("Invalid email address entered. No mail will be sent")
    } else {
      var subject = "Shared Print Retentions Spreadsheet Complete";
      var message = "Shared Print Retentions Spreadsheet lookup completed at " + donetime.toUTCString(); 
      message += "\nThis message sent from the sheet: " + SpreadsheetApp.getActive().getUrl(); // not sure if this is right sheet

      try {
        MailApp.sendEmail(emailAddress, subject, message); 
      } catch(e) {
        ui.alert("Error sending email. No mail will be sent")
        Logger.log("Error with email (" + emailAddress + "). " + e);
      } // end catch email
    } // end else is valid email pattern
  } // end if form sendEmail 
*/
}
//===================================================================================================
function updateSheetColumn(startingRow, rows, newValues, column, sheet) {  //https://developers.google.com/apps-script/guides/support/best-practices

numrowsremove = startingRow-2
newValues.splice(0, numrowsremove);

var formatColumn = sheet.getRange(column + ":" + column);
formatColumn.setNumberFormat("@"); // set column to be a text, not number, column - for merged OCLC numbers especially!
   
var rangeRows = (rows-(startingRow-1)+1)  // number of rows to fetch 
var sheetRange = column + startingRow + ":" + column + (rows + 1) ;
var allRange = sheet.getRange(sheetRange);// e.g. sheet.getRange("C2:C600") 
    
var updateValues = [] ; // create new array for update [column][index]
for (counter = 0; counter < newValues.length; ++counter) {   updateValues[counter] = new Array(1); } ;

for (counter = 0; counter < newValues.length; ++counter) {    
  if (typeof newValues[counter] == "undefined") { newValues[counter]= "";}
  updateValues[counter][0] = newValues[counter];  
}  // end for counter < new values 

allRange.setValues(updateValues) ; // actually update the sheet
   
} // end updateSheetColumn
//===================================================================================================
