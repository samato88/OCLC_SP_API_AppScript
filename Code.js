/* 
Check HathiTrust, OCLC, and/or Internet Archives for holdings and Shared Print registrations
Searches limited to 1000 row, otherwise gives warning
OCNs must be in column A, overwrites columns B-K

Given column a of OCLC numbers:
1)  Hathi - Access level in column D
2)  IA - Colummn E  
3)  OCLC for SPP retentions put in column C, retained by in column D 
4)  OCLC for current number, retrieve merged numbers  B, H

To Do:
  more error checking on api calls - try/catch them
  Add toast when done?:  SpreadsheetApp.getActiveSpreadsheet().toast('Complete', 'Status', 3); // so funny!

working on:
Exceeded maximum execution time  - hangs program - not caught :(

sometimes seeing : TypeError: undefined is not a function
                    BUT not limited to a certain OCN, not sure what triggers it, maybe something in IA lookup
IA error (shows in spreadsheet column: Exception: Address unavailable: https://archive.org/advancedsearch.php?q=oclc-id%3A36246&fl%5B%5D=licenseurl,identifier&sort%5B%5D=&sort%5B%5D=&sort%5B%5D=&rows=50&page=1&output=json)

trying try/catch for whole lookup - catch randomerrors and also timeout? = testing

Saw error one time (hung code) Log said failed "We're sorry, a server error occurred while reading from storage. Error code RESOURCE_EXHAUSTED.
"  -- https://cloud.google.com/apis/design/errors says this is a server 429 error

looking ~180
TypeError: retainedBY.join is not a function (one example happened on row 327 of test 400 - ocn: 32625860, though running that OCN alone is okay --- actually its the one after - 43096707)
-- sometimes retainedBY is a string - why???
May have fixed this ~ line 380  retainedBy = [];

IA error: Exception: Address unavailable: (https://archive.org/advancedsearch.php?q=oclc-id%3A1086981674&fl%5B%5D=licenseurl,identifier&sort%5B%5D=&sort%5B%5D=&sort%5B%5D=&rows=50&page=1&output=json
).  --- seems like there might have been a line break after "oclc-"

\
~ line 100 of sidebar.hmtl - try to get return value of getPercentDone 
  problem at line 500 in code.gs - not reading global var percentDone - or startLookup not updating global var
  https://developers.google.com/apps-script/guides/html/reference/run
  Setting above aside for now, make it a indetermiate status bar



Enhancements:  
--Set column widths?
--Translate which libraries retain - symbol -> library 
        would mean mapping to library name - EAST could in theory use spreadsheet, 
        others would need another api call, 
        BUT could get you also opac link. More than I'm willing to think about right now. 
--Make get oclc from isbn feature? or 2nd column of isbn to test if oclc doesn't match? API doesn't support isbn lookup but could possible 
    get from another API search
--Add check for holdings or retentions on symbol - symbol input in sidebar - not sure what the use case is here


*/
/*=====================================================================================================*/
/* Note: API target retained-holdings => current OCLC & who retains, does not return merged numbers */
/* don't forget to turn off all the logging when done */
/* adapted from:  https://github.com/suranofsky/tech-services-g-sheets-addon/blob/master/Code.gs */
 
var HTTP_OPTIONS = {muteHttpExceptions: true}
var apiService ;  // global used in multiple functions
var OCLCurl = 'https://americas.discovery.api.oclc.org/worldcat/search/v2/' ;
var ui = SpreadsheetApp.getUi();   // Or DocumentApp or FormApp. This in global scope.
var percentDone = "3" ; // used for progress bar - at least trying

function onOpen(e) { /* What should the add-on do when a document is opened */
  // the 'e' allows it to be closed later - https://developers.google.com/apps-script/guides/html/reference/host#close
  ui.createMenu('Shared Print Lookup')
      .addItem('Search by OCLC #s', 'showSidebar')
      //.addItem('Find OCLC from ISBN', 'showISBNbar')
      .addToUi();
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
   var dataTabName = form.searchForTab;
   var dataSheet = spreadsheet.getSheetByName(dataTabName);  

   PropertiesService.getUserProperties().setProperty('percentDone', 2); //start at 2%

   var SPP = form.SPP;  

   var lastRow = dataSheet.getLastRow();   
   if (lastRow > 1000) {
      ui.alert("This script works best with under 1,000 rows. \nPlease try again with a shorter sheet");
      return;
   }

   var oclcsRange = dataSheet.getRange(2,1,lastRow-1); // read from A2 to-> lastRow-1:  getRange(row, column, numRows, numColumns)
   var numRows = oclcsRange.getNumRows();
   var eastColumn = new Array(numRows); // store results for later writing to sheet
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
          //ui.alert(oclc)
          // check here which systems to check and do it!
        if (form.worldcatretentions) {  // this is the checkbox for WC retentions 
          let [numbEASTHoldings, retainedBY, currentOCLC] = getEASTHoldings(oclc, SPP) ;
          // ui.alert(retainedBY);

          if (numbEASTHoldings > 999999) {
            eastColumn[x-1] = "" // API returned wacky number for holdings, invalid OCLC
            currentOCLCColumn[x-1] = "Invalid OCN"; 
            continue
          } else if (typeof retainedBY != "undefined") {
            eastColumn[x-1] = numbEASTHoldings //array for updating sheet; x is 1, array index starts at 0
            // check how long retainedBY is - if more than one, can join? Is this the join error spot?
            Logger.log(typeof retainedBY)
            Logger.log(retainedBY)
            retainersColumn[x-1] = retainedBY.join(',');
          } else {
            eastColumn[x-1] = numbEASTHoldings //array for updating sheet; x is 1, array index starts at 0
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
              eastColumn[x-1] = "" ; // this is 0 from getEASTHoldings by virtue of API reporting 0 not ocn error
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
      //ui.alert(hathiColumn);
      //updateSheetColumn(numRows, hathiColumn, "E", dataSheet) ; 
      //updateSheetColumn(numRows, hathiIdColumn, "I", dataSheet) ; 
      //updateSheetColumn(numRows, hathiTitleColumn, "J", dataSheet) ; 
      updateSheetColumn(startingRow, numRows, hathiColumn, "E", dataSheet) ; 
      updateSheetColumn(startingRow, numRows, hathiIdColumn, "I", dataSheet) ; 
      updateSheetColumn(startingRow, numRows, hathiTitleColumn, "J", dataSheet) ; 
    }
    if (form.ia)   {updateSheetColumn(startingRow, numRows, iaColumn, "F", dataSheet) ; }

  if (form.worldcatretentions) { // this is the checkbox for EAST
      updateSheetColumn(startingRow, numRows, eastColumn,        "C", dataSheet) ; 
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

    lock.releaseLock();
    } catch {
          ui.alert("Failed to sheet lock - reload sheet and try again.") // not sure this is a good idea
          return ;
    }
  } // end try getting a lock and updating sheet
  catch { // catch failed lock
    //Logger.log("Failed to get lock") ;
    ui.alert("Failed to lock sheet - reload sheet and try again.") // not sure this is a good idea, slightly different error so I can tell where it happened
    return ;
  }

  // send email if requested (could do html email but have not coded https://blog.gsmart.in/google-apps-script-send-html-email/)
  if (form.sendEmail) {
    var donetime = new Date();
    //var emailAddress = "samato@blc.org";
    var emailAddress = form.emailAddress ;
    var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;  
    if(emailPattern.test(emailAddress) == false) {
      Logger.log("Invalid Email Address Entered. " + emailAddress);
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

} // end start lookup

//===================================================================================================
//function updateSheetColumn(rows, newValues, column, sheet) {  //https://developers.google.com/apps-script/guides/support/best-practices
function updateSheetColumn(startingRow, rows, newValues, column, sheet) {  //https://developers.google.com/apps-script/guides/support/best-practices
  //Logger.log(newValues) ;
  //Logger.log(rows) ;
 
 // ui.alert(newValues)
 //ui.alert(rows)
 //ui.alert(rows-(startingRow-1)) // 4-(4-1) = 4-3 = 1
 numrowsremove = startingRow-2
 //ui.alert(numrowsremove)
newValues.splice(0, numrowsremove);
//ui.alert(newValues)

    var formatColumn = sheet.getRange(column + ":" + column);
    formatColumn.setNumberFormat("@"); // set column to be a text, not number, column - for merged OCLC numbers especially!
   
    //var rangeRows = rows + 1 // number of rows to fetch 
    var rangeRows = (rows-(startingRow-1)+1)  // number of rows to fetch 
    //ui.alert(rangeRows)

    //var sheetRange = column + "2:" + column + rangeRows ;
    var sheetRange = column + startingRow + ":" + column + (rows + 1) ;
         
   // ui.alert(rows)
    //ui.alert(sheetRange);
    var allRange = sheet.getRange(sheetRange);// e.g. sheet.getRange("C2:C600") 
    
    var updateValues = [] ; // create new array for update [column][index]
    for (counter = 0; counter < newValues.length; ++counter) {   updateValues[counter] = new Array(1); } ;
    for (counter = 0; counter < newValues.length; ++counter) {    
      if (typeof newValues[counter] == "undefined") { newValues[counter]= "";}
      updateValues[counter][0] = newValues[counter];  
  }  // end for counter < new values 
    //Logger.log("updateValues: " + updateValues);
    //ui.alert(updateValues)
    allRange.setValues(updateValues) ; // actually update the sheet
   
} // end updateSheetColumn

//===================================================================================================

function getEASTHoldings(oclc, SPP) {
    // https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app
    let ui = SpreadsheetApp.getUi();
    let numberEASTHoldings = 0 ;
    let retainedBy = [] ;
    let currentOCLC = "" ;

    if (!apiService.hasAccess()) { // check if expired.  not sure this is the right way to do that
      getApiService();
    }
  /*
    try {
      var response = UrlFetchApp.fetch(url).getContentText();
    }
    catch(err) {
      Logger.log(response)
      Logger.log(err)
    }
    */
     if (apiService.hasAccess()) { //      
      //Logger.log(service.getAccessToken());
      var url = OCLCurl + 'retained-holdings?oclcNumber=' + oclc + '&spProgram=' + SPP ;
      var response = UrlFetchApp.fetch(url, {
        headers: {
           Authorization: 'Bearer ' + apiService.getAccessToken()
        },
        validateHttpsCertificates: false,
        muteHttpExceptions: true
       });
      //Logger.log(response)

     ////CHECK RESPONSE HEADER NOT 403 or 404  //Logger.log(response.getHeaders()); //Logger.log(response.getContentText());
     //Logger.log(response)
     if(response.getResponseCode() != 200) { // not 200
       //Logger.log(response.getResponseCode());
          numberEASTHoldings = "";
          retainedBy = [];
          currentOCLC = "Server Error: " + response.getResponseCode() ;
     } else { //valid response
          var results = JSON.parse(response.getContentText());
          //Logger.log(results);

          if (results.numberOfHoldings) {
            //Logger.log("results.numberofHoldings: " + results.numberOfHoldings);
            numberEASTHoldings = results.numberOfHoldings;

            for (lib in results.detailedHoldings) {
/*
{detailedHoldings=[{format=zu, lhrLastUpdated=20210215, sharedPrintCommitments=[{actionNote=committed to retain, dateOfAction=20160630, commitmentExpirationDate=20310630, authorization=EAST, institution=MBU}], lhrControlNumber=352397802, lhrDateEntered=20210215, location={sublocationCollection=BOSS, holdingLocation=BOS}, hasSharedPrintCommitment=Y, summary=Local Holdings Available., oclcNumber=123456}], numberOfHoldings=1.0} */
              //Logger.log(results.detailedHoldings[lib].location.holdingLocation);  
              retainedBy.push(results.detailedHoldings[lib].location.holdingLocation); //  holdings symbol
            } // end foreach results.detailedHoldings (LHRs)
            
            //ui.alert("EAST: " + numberEASTHoldings); 
            //Logger.log("results.detailedHoldings[0].oclcNumber: "+ results.detailedHoldings[0].oclcNumber);
            currentOCLC = results.detailedHoldings[0].oclcNumber
          }
     } // end else valid response
    } // end apiService.hasAccess 
       //Logger.log(apiService.getLastError());
    return [numberEASTHoldings, retainedBy, currentOCLC] ;
} // end getEASTHoldings

//===================================================================================================
 
function getHathiHoldingsMerged(oclc, merged) { //https://catalog.hathitrust.org/api/volumes/brief/oclc/14219719.json
  // local variables:
  var hathiurl = "http://catalog.hathitrust.org/api/volumes/brief/oclc/" +  oclc +  ".json"

  try {
    var response = JSON.parse(UrlFetchApp.fetch(hathiurl).getContentText());
    var rights = "" ;
    var recordurl = "" ;
    var r = "";
    var htid = "" ;
    var htitle = "" ;
    //Logger.log(response);
    //Logger.log("type of response: " + typeof response.items[0]) ;

    if (typeof response.items[0] == "undefined") {
      if (merged && merged.length > 0) { // there are alt oclc numbers
        for (altNumb in merged) {
          //Logger.log("altNumb: " + altNumb + " Merged: " + merged + " typeof: " + typeof merged + " len: " + merged.length);
          //Logger.log(merged[altNumb]);
            hathiurl = "http://catalog.hathitrust.org/api/volumes/brief/oclc/" +  merged[altNumb] +  ".json"
            response = JSON.parse(UrlFetchApp.fetch(hathiurl,HTTP_OPTIONS).getContentText());
              if (typeof response.items[0] != "undefined") {
                Logger.log("type of response alt: " + typeof response.items[0]) ;
                break;
              } // if got response break to outer
       } // end for altNumb in merged
     } // end if merged numbers
    }// end if undefined response
  
  
    if   (typeof response.items[0] != "undefined") { // make sure you got something
      r = response.items[0].usRightsString ;
      htid = response.items[0].htid
      //  pubdate   = response.records[Object.keys(response.records)[0]].publishDates;  //yes
      //  recordurl = response.records[Object.keys(response.records)[0]].recordURL // this works too  
     htitle = response.records[Object.keys(response.records)[0]].titles[0];
     //Logger.log(hathiurl);
     //Logger.log(htitle);
     recordurl = "https://catalog.hathitrust.org/Record/" + response.items[0].fromRecord  ;
     rights = '=hyperlink("' + recordurl +  '","' + r + '")';    
    }

  } // end try
  catch(err) {
    rights = err;
    htid = ""
    htitle = ""
  }
 
  //return rights ;
  return [rights, htid, htitle] ;
}// end get HathiHoldings  
//===================================================================================================

function getIAHoldingsMerged(oclc, merged) { 
  // https://archive.org/advancedsearch.php?q=oclc-id%3A31773958&fl%5B%5D=licenseurl&sort%5B%5D=&sort%5B%5D=&sort%5B%5D=&rows=50&page=1&output=json&callback=callback
  // https://archive.org/advancedsearch.php?q=oclc-id:31773958
  // https://archive.org/advancedsearch.php?q=oclc-id:31773958&fl[]=identifier&output=json // this one clean!
  // https://archive.org/advancedsearch.php?q=oclc-id%3A31773958&fl%5B%5D=licenseurl&sort%5B%5D=&sort%5B%5D=&sort%5B%5D=&rows=50&page=1&fl%5B%5D=identifier&output=json
  
  // local variables:
  //var iaurl = "https://archive.org/advancedsearch.php?q=oclc-id%3A" +  oclc +  "&fl%5B%5D=licenseurl,identifier&sort%5B%5D=&sort%5B%5D=&sort%5B%5D=&rows=50&page=1&output=json"
  var iaurl = "https://archive.org/advancedsearch.php?q=oclc-id:" + oclc + "&fl[]=identifier&output=json"
  var iaholdings = ""
 // var response = JSON.parse(UrlFetchApp.fetch(iaurl).getContentText());
  
  try {
   var r =  UrlFetchApp.fetch(iaurl);
    Logger.log("OCLC: " + oclc + "iaurl: " + iaurl);
 
   // Logger.log(r.getResponseCode());

   // PUT TRY CATCH HERE!??
   var response = JSON.parse(r.getContentText()) ;
   Logger.log("IA Response: " + response);
   if (response) {
   if  (response.response.numFound == 0 ) {
     if (typeof merged !== 'undefined') {

      if (merged.length > 0) { // there are alt oclc numbers
       for (altNumb in merged) {
         iaurl = "https://archive.org/advancedsearch.php?q=oclc-id%3A" +  merged[altNumb] +  "&fl%5B%5D=licenseurl,identifier&sort%5B%5D=&sort%5B%5D=&sort%5B%5D=&rows=50&page=1&output=json"
          response = JSON.parse(UrlFetchApp.fetch(iaurl,HTTP_OPTIONS).getContentText());
          // response = JSON.parse(UrlFetchApp.fetch(iaurl,HTTP_OPTIONS).getResponseCode);
          if (response.response.numFound > 0) { // probably should use more error checking here
             break;
          } // if got response break to outer
        } // end for altNumb in merged
      } // end if merged numbers
     } // end if merged defined
   } // end if numFound == 0
  
   if (response.response.numFound > 0) { // make sure you got something
     var id = response.response.docs[0].identifier ;
     var recordurl = "https://archive.org/details/" + id  ;
     var iaholdings = '=hyperlink("' + recordurl +  '","' + id + '")';    
   } // end response > 0
   } // end if response.response
  } catch(err) {
    iaholdings = err
  }
 
  return iaholdings ;
}// end get HathiHoldings  
//===================================================================================================

/* function getCurrentOCLC(oclc) { // I think this can now be deleted
    //https://americas.api.oclc.org/discovery/worldcat/v1/bibs/650 => MergedOclcNumbers
    //var apiService = getApiService();

  if (!apiService.hasAccess()) { // check if expired.  not sure this is the right way to do that
      getApiService();
  }
   
   if (apiService.hasAccess()) {
     //Logger.log(service.getAccessToken());
     var url = 'https://americas.api.oclc.org/discovery/worldcat/v1/bibs/' + oclc;
     var response = UrlFetchApp.fetch(url, {
       headers: {
         Authorization: 'Bearer ' + apiService.getAccessToken()
       },
       validateHttpsCertificates: false,
       muteHttpExceptions: true
     });

     ////NEED TO CHECK RESPONSE HEADER NOT 403 or 404 
     //Logger.log(response.getHeaders()); //Logger.log(response.getContentText());
     
     if(response.getResponseCode() != 200) {
       Logger.log(response.getResponseCode());
       return { 
          currentOCLC: "Server Error: " + response.getResponseCode() , 
          mergedOCLC: ""
         }; 
     } else { //valide response
          var result = JSON.parse(response.getContentText());
          //Logger.log(response.getContentText())
          //ui.alert(result.identifier.oclcNumber)
          //Logger.log(result.identifier.mergedOclcNumbers.join(','))
     } // end else is valid response
   } else {
       Logger.log(apiService.getLastError());
       return { 
          currentOCLC: apiService.getLastError(), 
          mergedOCLC: ""
         }; 
 
  } // end else doesn't have access
 
  if (result.identifier.mergedOclcNumbers) {
    var Merged = result.identifier.mergedOclcNumbers.join(',') 
  }
   
  return { 
          currentOCLC: result.identifier.oclcNumber, 
          mergedOCLC: Merged
         }; 
  } // end function getCurrentOCLC
*/
//===================================================================================================  
function getWorldCatHoldings(oclc, edition) {
//holdingsAllEditions=true
    //https://americas.api.oclc.org/discovery/worldcat/v1/bibs-holdings?oclcNumber=650&heldInCountry=US =>
    //briefRecords -> institutionHolding -> totalHoldingCount
    if (!apiService.hasAccess()) { // check if expired.  not sure this is the right way to do that
      getApiService();
    }
   
   if (apiService.hasAccess()) {
    // Logger.log("has service")
     var url2 = OCLCurl + 'bibs-holdings?oclcNumber=' + oclc + '&holdingsAllEditions=' + edition ;
    // var url2 = OCLCurl + 'bibs-holdings?oclcNumber=' + oclc + '&heldInCountry=US' + '&holdingsAllEditions=' + edition ;
     var response2 = UrlFetchApp.fetch(url2, {
       headers: {
         Authorization: 'Bearer ' + apiService.getAccessToken()
       },
       validateHttpsCertificates: false,
              muteHttpExceptions: true
     });
     //Logger.log(url2)
     //Logger.log(response2)
     if(response2.getResponseCode() != 200) {
       Logger.log(response2.getResponseCode());
       return  response2.getResponseCode() ;
     } else {
       var result2 = JSON.parse(response2.getContentText());
       //Logger.log(result2);
       if (result2.numberOfRecords > 0) {
        var otitle = result2.briefRecords[0].title ;
        var CurrentOCN = result2.briefRecords[0].oclcNumber ;
        var Merged = result2.briefRecords[0].mergedOclcNumbers ;
        holdingsCount = result2.briefRecords[0].institutionHolding.totalHoldingCount ;
       } else { //numberOfRecords is zero, this happens with invalid ocn e.g. 123456789012345
          holdingsCount = "invalid oclc"
          var otitle = "";
          var CurrentOCN = "" ;
          var Merged = "" ;
       } // end else numberOfRecords = 0
     }// end valid response
   } else {
       Logger.log(apiService.getLastError());
       return   "Authorization failed"        
  } // end else doesn't have access
    return [holdingsCount, otitle, CurrentOCN, Merged];
  
  } // end get worldcat holdings

//===================================================================================================  
function getPercentDone() {
  Logger.log("PD:" + percentDone);
  Logger.log("PD property: " + PropertiesService.getUserProperties().getProperty('percentDone'));
  //ui.alert("PD: " + percentDone)
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
      // Set the endpoint URLs.
      .setTokenUrl('https://oauth.oclc.org/token')

      // Set the client ID and secret.
      .setClientId(myKey)
      .setClientSecret(mySecret)

      // Sets the custom grant type to use.
      .setGrantType('client_credentials')
      //.setScope('DISCOVERY')
      .setScope('wcapi')
      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties());
}

//===================================================================================================    
function reset() { // Reset the authorization state, so that it can be re-tested.
  getApiService().reset();
}
//===================================================================================================    

