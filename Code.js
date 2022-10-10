/* 
Check HathiTrust and/or Internet Archives for holdings and OCLC for Shared Print registrations
Searches limited to 1000 row, otherwise gives warning
OCNs must be in column A, overwrites columns B-K

Given column A of OCLC numbers:
1)  Hathi - Access level in column E, ID in I, title in J
2)  IA - Colummn F  
3)  OCLC for SPP retentions put in column C, retained by in column D 
4)  OCLC for current number, retrieve merged numbers  B, H
5)  OCLC holdings in column G, title in K

Possible Enhancements:  
--Translate which libraries retain - symbol -> library, and catalog link -  would need another api call
--Make get oclc from isbn feature? or 2nd column of isbn to test if oclc doesn't match? API doesn't support isbn lookup but could possible perhaps with bib search first
 --Add check for holdings or retentions on symbol - symbol input in sidebar   
 --Add field for holdings in 583$3
*/
/*=====================================================================================================*/
/* Note: API target retained-holdings => current OCLC & who retains, does not return merged numbers */
/* adapted from:  https://github.com/suranofsky/tech-services-g-sheets-addon/blob/master/Code.gs */
 
var HTTP_OPTIONS = {muteHttpExceptions: true}
var apiService ;  // global used in multiple functions
var OCLCurl = 'https://americas.discovery.api.oclc.org/worldcat/search/v2/' ;
var ui = SpreadsheetApp.getUi();   // global scope.
var percentDone = "3" ; // used for progress bar - currently not working, using static 'working' bar

function onOpen(e) { /* What should the add-on do when a document is opened */
  // the 'e' allows it to be closed later - https://developers.google.com/apps-script/guides/html/reference/host#close
      ui.createAddonMenu()
      .addItem('Search by OCLC #s', 'showSidebar')
      .addToUi();
}

function onInstall() {
  onOpen();
} 

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Lookup:')
      .setWidth(500);
      ui.showSidebar(html);    
}

function getTabs() {
    var out = new Array();
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i=0 ; i<sheets.length ; i++) {
       out.push( [ sheets[i].getName() ] );
    }
    return out;
}

function getStoredAPIKey() {
     return PropertiesService.getUserProperties().getProperty('apiKey')
}
function getStoredAPISecret() {
     return PropertiesService.getUserProperties().getProperty('apiSecret')
}

//FUNCTION IS LAUNCHED WHEN THE 'START SEARCH' BUTTON ON THE SIDEBAR IS CLICKED //'form' REPRESENTS THE FORM ON THE SIDEBAR
function startLookup(form) {
   "use strict" ;
   var apiKey = form.apiKey; //MAKE SURE THE OCLC API KEY HAS BEEN ENTERED IF NEEDED
   var apiSecret = form.apiSecret; //MAKE SURE THE OCLC API SECRET HAS BEEN ENTERED IF NEEDED
  
   PropertiesService.getUserProperties().setProperty('apiKey', apiKey);
   PropertiesService.getUserProperties().setProperty('apiSecret', apiSecret);
   
   if (form.worldcatretentions) { // if worldcat search box checked - check for key and secret and is authorized
       if ((apiKey == null || apiKey == "")) {
         ui.alert("OCLC API Key is Required for WorldCat lookups");
         return;
       } else if (apiSecret == null || apiSecret == "")  {
         ui.alert("OCLC API Secret is Required for WorldCat lookups");
         return;
       }
       apiService = getApiService(); 

       if (!apiService.hasAccess()) {
         ui.alert("Invalid API Key or Secret.  Please re-enter or uncheck 'Retentions in OCLC' box") ;
         return
       } 
   } // end check for key and secret and is authorized

   var spreadsheet = SpreadsheetApp.getActive();
   spreadsheet.setFrozenRows(1); // freeze the top row
   var dataTabName = form.searchForTab;
   var dataSheet = spreadsheet.getSheetByName(dataTabName);  

   PropertiesService.getUserProperties().setProperty('percentDone', 2); //start at 2%, currently not using this

   var SPP = form.SPP;  

   var lastRow = dataSheet.getLastRow();   
   if (lastRow > 1000) {
      ui.alert("This script works best with under 1,000 rows. \nPlease try again with a shorter sheet");
      return;
   }

   var oclcsRange = dataSheet.getRange(2,1,lastRow-1); // read from A2 to-> lastRow-1:  getRange(row, column, numRows, numColumns)
   var numRows = oclcsRange.getNumRows();
   var sppColumn = new Array(numRows); // store results for later writing to sheet
   var retainersColumn = new Array(numRows); // store who retains
   var hathiColumn = new Array(numRows); // store results for later writing to sheet
   var hathiTitleColumn = new Array(numRows); // store Hathi Titles
   var hathiIdColumn = new Array(numRows); // store Hathi Id 
   var iaColumn = new Array(numRows); // store results for later writing to sheet
   var currentOCLCColumn = new Array(numRows); // store current oclc
   var mergedOCLCColumn = new Array(numRows); // store merged oclcs
   var usHoldingsColumn = new Array(numRows); // store worldcat same edition holdings
   var worldcatTitleColumn = new Array(numRows); // store worldcat title
   var edition = "" ;  // used to set header column with 'same' or 'any' edition
   var startingRow = form.startRow;

if (startingRow > lastRow) {
      ui.alert("Start search at row number is "+ startingRow + ", but this sheet only has " + lastRow.toString() + " Rows.\nPlease try again with a lower start row number");
      return;
   }
if (startingRow < 2) {
      ui.alert("'Start search at row number' must be greater than 1.\nPlease try again");
      return;
   }

   if (form.WCHoldings == true) {edition = "Any"} else { edition = "Same"} ;
   var columnHeaders = [["WorldCat OCLC", SPP + " Retentions", "Retained By Symbol", "Hathi", "IA", "US Holdings (" + edition + " edition) in WorldCat", "Merged OCLC numbers", 	"Hathi ID", "Hathi Title", "OCLC Title"]] ;

  try {// wondering if this will catch time outs - try/catch around all lookups
   var x = 1;
   if (startingRow != null && startingRow != "" && startingRow !=1) 
    { 
      x = startingRow-1 ;
    }  else { 
      startingRow=2; // start at 2 for sheet update
    } // end if staringRow not null 

   for (x; x <= numRows; x++) { //FOR EACH ITEM TO BE LOOKED UP IN THE DATA SPREADSHEET:
   // for (var x =1; x <= numRows; x++) { //FOR EACH ITEM TO BE LOOKED UP IN THE DATA SPREADSHEET:
      var oclcCell = oclcsRange.getCell(x,1);
      var merged = "" ;
      var htid = "" ;   

      percentDone = parseInt((x/numRows) * 100);
      PropertiesService.getUserProperties().setProperty('percentDone', percentDone);
      //ui.alert(percentDone); // this works, but global variable seems not around at end of script??

      Logger.log("Row: " + x + " OCLC: " + oclcCell.getValue());

      if (!oclcCell.isBlank()) {
          var oclc = oclcCell.getValue() ;
          if (isNaN(oclc)) { //test if NaN, if so, skip
              oclc = oclc.toString().toLowerCase(); 
              if (oclc.startsWith("ocn") || oclc.startsWith("ocm")) { 
                oclc = oclc.replace("ocn", "");
                oclc = oclc.replace("ocm", "");
              } else {
                currentOCLCColumn[x-1] = "Invalid OCN"; 
                continue ;
              } // doesn't start with ocn or ocm
          } // end test if NaN, if so remove prefix or skip
           
          oclc = parseInt(oclc, 10); // trim leading zeros (will round any decimals - should not be any anyways.)
        
          // check here which systems to check and do it!
        if (form.worldcatretentions) {  // this is the checkbox for WC retentions 
          let [numbSPPHoldings, retainedBY, currentOCLC] = getSPPHoldings(oclc, SPP) ;
          // ui.alert(retainedBY);

          if (numbSPPHoldings > 999999) {
            sppColumn[x-1] = "" // API returned wacky number for holdings if invalid OCLC, this appears fixed now
            currentOCLCColumn[x-1] = "Invalid OCN"; 
            continue
          } else if (typeof retainedBY != "undefined") {
            sppColumn[x-1] = numbSPPHoldings //array for updating sheet; x is 1, array index starts at 0
            retainersColumn[x-1] = retainedBY.join(',');
          } else {
            sppColumn[x-1] = numbSPPHoldings //array for updating sheet; x is 1, array index starts at 0
            retainersColumn[x-1] = "";
          }// end else ifs for holdings > 999999

          if (form.WCData || form.WCHoldings) {
            // SHOULD CHECK 'currentOCLC' for validity and use that when valid
            // overwrite currentOCLC cuz above leaves out current if retained is 0
            [usHoldingsColumn[x-1], worldcatTitleColumn[x-1] , currentOCLC, merged] = getWorldCatHoldings(oclc,form.WCHoldingsType) ;
            // check if usHoldings came back as 'invalid oclc'
            if (usHoldingsColumn[x-1] == 'invalid oclc') {
              currentOCLC = "Invalid OCN";
              usHoldingsColumn[x-1] = "";
              sppColumn[x-1] = "" ; // this is 0 from getSPPHoldings by virtue of API reporting 0 not ocn error
            } // end if usHoldings is invalid oclc
          } // end if WCData or WCHoldings

          if (merged) {
              mergedOCLCColumn[x-1] = String(merged);
          } // end if merged
          
          if (currentOCLC == "Invalid OCN") {
            currentOCLCColumn[x-1] =  currentOCLC; 
          } else {
            currentOCLCColumn[x-1] = '=hyperlink("https://worldcat.org/oclc/' + currentOCLC +  '","' + currentOCLC + '")'; 
          } // end if currentOCLC is 'Invalid OCN'
        } // end form.worldcatretentions   
       
        if (form.hathi) { 
          var [hathiHoldings, htid, htitle] = getHathiHoldingsMerged(oclc, merged) ;
          //ui.alert("row (index starts at 0): " + String(x))
          hathiColumn[x-1] = hathiHoldings ;
          hathiIdColumn[x-1] = htid ;
          hathiTitleColumn[x-1] = htitle ;
        } // end form.hathi
       
        if (form.ia){
          var iaHoldings = getIAHoldingsMerged(oclc, merged) ;
          iaColumn[x-1] = iaHoldings ;
        } // end form.ia  
       
      } // end if oclc cell not blank
    } // end for each OCLC
  
  } catch(err) {
     Logger.log(err);
     ui.alert(err); // will this catch Exceeded maximum execution time
     // need to return here? it does write out what it got... but doesn't continue on
   } // end catch

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

   Logger.log('Script done');
   PropertiesService.getScriptProperties().setProperty('run', 'done'); //leap of faith that above is synchronous
 

} // end start lookup
 
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
function getPercentDone() { // never got this working - if any has ideas would love to hear them
  //ui.alert("PD Property: " + PropertiesService.getUserProperties().getProperty('percentDone')); // this seems to always be 100 when we get here
  return PropertiesService.getUserProperties().getProperty('percentDone');

  //return percentDone ;
}
//===================================================================================================  

function getApiService() {  //https://github.com/gsuitedevs/apps-script-oauth2/blob/master/samples/TwitterAppOnly.gs
                            //https://github.com/gsuitedevs/apps-script-oauth2
  myKey = PropertiesService.getUserProperties().getProperty('apiKey')
  mySecret = PropertiesService.getUserProperties().getProperty('apiSecret')  

  //myKey = 'testingForFailure'
  //mySecret = 'testingForFailure'

  return OAuth2.createService('WorldCat Discovery API')
      .setPropertyStore(PropertiesService.getUserProperties()) // use cache as per advice in https://github.com/gsuitedevs/apps-script-oauth2
      .setCache(CacheService.getUserCache())
      .setTokenUrl('https://oauth.oclc.org/token')      // Set the endpoint URLs.

      // Set the client ID and secret.
      .setClientId(myKey)
      .setClientSecret(mySecret)

      // Sets the custom grant type to use.
      .setGrantType('client_credentials')
      .setScope('wcapi')
      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties());
}

//===================================================================================================    
function reset() { // Reset the authorization state, so that it can be re-tested.
  getApiService().reset();
}
//===================================================================================================    

