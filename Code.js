/* 
Check HathiTrust and/or Internet Archives for holdings and OCLC for Shared Print registrations
Sheets limited to 10,000 row, otherwise gives warning. This is arbitrary.
OCNs must be in column A, overwrites columns B-K (will warn if data exists)
Given column A of OCLC numbers:
1)  Hathi - Access level in column E, ID in I, title in J
2)  IA - Colummn F  
3)  OCLC for SPP retentions put in column C, retained by in column D 
4)  OCLC for current number, retrieve merged numbers  B, H
5)  OCLC holdings in column G, title in K

To Do:
--add column for EAST Holders (not yet retained)

April 2024: Added Metadata API key use rather than just discovery search API

Possible Enhancements: 
--Translate which libraries retain - symbol -> library, and catalog link -  would need another api call
--Make get oclc from isbn feature? or 2nd column of isbn to test if oclc doesn't match? API doesn't support isbn lookup but could possible perhaps with bib search first
 --Add check for holdings or retentions on symbol - symbol input in sidebar   
 --Add field for holdings in 583$3
 --Add check for retentions by ALL programs
*/
/*=====================================================================================================*/
/* Note: API target retained-holdings => current OCLC & who retains, does not return merged numbers */
/* adapted from:  https://github.com/suranofsky/tech-services-g-sheets-addon/blob/master/Code.gs */
 
var HTTP_OPTIONS = {muteHttpExceptions: true}
var apiService ;  // global used in multiple functions (actually that might not be true)
var OCLCurl = 'https://americas.discovery.api.oclc.org/worldcat/search/v2/' ; // this is if discovery key - might need to set elsewhere
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
      // if name is the current active sheet, out.unshift rather than push, so it will be the selected tab
      if (SpreadsheetApp.getActiveSheet().getName() == sheets[i].getName()) { // get name of current tab
           out.unshift( [ sheets[i].getName() ] );
       } else {
           out.push( [ sheets[i].getName() ] );
       }
    }
    return out;
}

function getStoredAPIKey(apitype) {
  if (apitype == "discovery") {
    //Logger.log(apitype)
    //Logger.log(PropertiesService.getUserProperties().getProperty('apiKey'))
    return PropertiesService.getUserProperties().getProperty('apiKey')
  } else {
    //Logger.log(apitype)
    //Logger.log(PropertiesService.getUserProperties().getProperty('mapiKey'))
    return PropertiesService.getUserProperties().getProperty('mapiKey')
  }
}

function getStoredAPISecret(apitype) {
  if (apitype == "discovery") {
      return PropertiesService.getUserProperties().getProperty('apiSecret')
  } else {
      return PropertiesService.getUserProperties().getProperty('mapiSecret')
  }
}

//FUNCTION IS LAUNCHED WHEN THE 'START SEARCH' BUTTON ON THE SIDEBAR IS CLICKED //'form' REPRESENTS THE FORM ON THE SIDEBAR
function startLookup(form) {
   "use strict" ;
   //ui.alert(form.apitype);
   var apiKey = form.apiKey; //MAKE SURE THE OCLC API KEY HAS BEEN ENTERED IF NEEDED
   var apiSecret = form.apiSecret; //MAKE SURE THE OCLC API SECRET HAS BEEN ENTERED IF NEEDED
   PropertiesService.getUserProperties().setProperty('cancelscript', 'no'); // set to no, gets set to use if cancel in form hit

   if (form.worldcatretentions) { // if worldcat search box checked - check for key and secret and is authorized
       if ((apiKey == null || apiKey == "")) {
         ui.alert("OCLC API Key is Required for WorldCat lookups");
         return;
       } else if (apiSecret == null || apiSecret == "")  {
         ui.alert("OCLC API Secret is Required for WorldCat lookups");
         return;
       }
      
      if (form.apitype == "discovery") {
        apiService = getDiscoveryApiService(form.apitype); 
      } else if (form.apitype == "metadata") {
        apiService = getDiscoveryApiService(form.apitype); 
      }

      if (form.apitype == "discovery") {
         //Logger.log("setting API Key " + apiKey)
         PropertiesService.getUserProperties().setProperty('apiKey', apiKey);
         PropertiesService.getUserProperties().setProperty('apiSecret', apiSecret);
      } else if (form.apitype == "metadata")  {
         //Logger.log("setting API Secret for " + form.apitype + ": " + apiSecret)
         PropertiesService.getUserProperties().setProperty('mapiKey', apiKey); // save metadata api key
         PropertiesService.getUserProperties().setProperty('mapiSecret', apiSecret); // save metadata api secret
      }

       if (!apiService.hasAccess()) {
         ui.alert("Invalid API Key or Secret.  Please re-enter or uncheck 'Retentions in OCLC' box") ;
         return
       } 
   } // end check for key and secret and is authorized

   var spreadsheet = SpreadsheetApp.getActive();
   spreadsheet.setFrozenRows(1); // freeze the top row
   var dataTabName = form.searchForTab;
   var dataSheet = spreadsheet.getSheetByName(dataTabName);  
    
   if (SpreadsheetApp.getActiveSheet().getName() != dataTabName) { // if active sheet is not the same listed in the search form
      var responset = ui.alert("I noticed you are in the '" + SpreadsheetApp.getActiveSheet().getName() + "' tab, but the lookup tab is set to '" + dataTabName + "'.\nDo you want to continue with the lookup in " + dataTabName + "?", ui.ButtonSet.YES_NO);
      if (responset === ui.Button.NO) { // https://code.luasoftware.com/tutorials/google-apps-script/google-apps-script-confirm-dialog/
          return;
        }
   } // end if active tab not the same as tab in search form

   PropertiesService.getUserProperties().setProperty('percentDone', 2); //start at 2%, currently not using this

   var SPP = form.SPP;  

   var lastRow = dataSheet.getLastRow();   
   if (lastRow > 10000) { // Had this at 1000 before timeouts, now arbitrarily set at 10,000  
      ui.alert("Sheets over 10,000 lines can be unwidely. \nPlease try again with a shorter sheet");
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
   const startTime = new Date();
   const maxRunTime = form.timelimit;
   const columnsToCheck = [] ;

   console.log(startTime);

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

  try {// wondering if this will catch time outs - try/catch around all lookups -Nope.
   var x = 1;
   if (startingRow != null && startingRow != "" && startingRow !=1) 
    { 
      x = startingRow-1 ;
    }  else { 
      startingRow=2; // start at 2 for sheet update
    } // end if staringRow not null 

// Check for data in columns where output will be, alert if anything will be overwritten
  if (form.worldcatretentions) { columnsToCheck.push("B", "C","D") } 
  if (form.hathi) { columnsToCheck.push("E", "I","J")  }
  if (form.ia){ columnsToCheck.push("F") }
  if (checkForData(dataSheet, columnsToCheck, startingRow, lastRow)) {
       var messageColumns = columnsToCheck.slice(0, -1).join(', ')+' and/or '+columnsToCheck.slice(-1);
       var response = ui.alert("There is data at or below row " + startingRow + " in column(s):\n\n" + messageColumns + "\n\nthat may be overwritten. Continue?", ui.ButtonSet.YES_NO);
        if (response === ui.Button.NO) { // https://code.luasoftware.com/tutorials/google-apps-script/google-apps-script-confirm-dialog/
          return;
        }
     } // end if checkForData finds data
      

   for (x; x <= numRows; x++) { //FOR EACH ITEM TO BE LOOKED UP IN THE DATA SPREADSHEET:
      //ui.alert(x + PropertiesService.getUserProperties().getProperty('cancelscript')); // did cancel get hit and make it here?
      if(PropertiesService.getUserProperties().getProperty('cancelscript') =="yes") {
        ui.alert("SCRIPT CANCELLED")
        return
     } // end if cancel

      let elapsed = Date.now() - startTime;
      if (elapsed/1000 > maxRunTime-15) { // check if getting close to timeout
      //if (elapsed/1000 >  15) { // just testing - quick turnaround
        const stopTime = new Date();
        console.log("Script stoped at: " + stopTime + " Total run time: " + elapsed/1000 + " seconds.");
        break; // stop the loop if the maxRunTime has been reached, with a 15 second leeway
        }

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
              if (oclc.startsWith("ocn") || oclc.startsWith("ocm") || oclc.startsWith("on")) { 
                oclc = oclc.replace("ocn", "");
                oclc = oclc.replace("ocm", "");
                oclc = oclc.replace("on", "");
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
     Logger.log("Error at row " + x);
     Logger.log(err);
     Logger.log(err.stack);
     ui.alert(err.name + " , " + err.message); // will this catch Exceeded maximum execution time
     // need to return here?  write out what it got?... but doesn't continue on
   } // end catch

  // update sheet was inline here
  updateSheet(form, startingRow, numRows, dataSheet, hathiColumn,hathiIdColumn, hathiTitleColumn, iaColumn, sppColumn, retainersColumn,       currentOCLCColumn, usHoldingsColumn,mergedOCLCColumn, worldcatTitleColumn, columnHeaders)
   Logger.log('Script done');
   //PropertiesService.getScriptProperties().setProperty('run', 'done'); //leap of faith that above is synchronous 
  return
} // end start lookup
//===================================================================================================  

function getApiService(apitype) {  //https://github.com/gsuitedevs/apps-script-oauth2/blob/master/samples/TwitterAppOnly.gs
                            //https://github.com/gsuitedevs/apps-script-oauth2

  if (apitype == "metadata") {
    OCLCurl = "https://metadata.api.oclc.org/worldcat";
    scope = "WorldCatMetadataAPI"
    myKey = PropertiesService.getUserProperties().getProperty('mapiKey')
    mySecret = PropertiesService.getUserProperties().getProperty('mapiSecret')  
  } else { // discovery
    OCLCurl = 'https://americas.discovery.api.oclc.org/worldcat/search/v2/' ;  
    scope = "wcapi" ;
    myKey = PropertiesService.getUserProperties().getProperty('apiKey')
    mySecret = PropertiesService.getUserProperties().getProperty('apiSecret') 
  }
  
  //myKey = 'testingForFailure'
  //mySecret = 'testingForFailure'
  Logger.log(PropertiesService.getUserProperties().getKeys());

// this works but will use the wrong persisted token based if switching between metadata and discovery apis
 return OAuth2.createService('WorldCat Discovery API')
      .setPropertyStore(PropertiesService.getUserProperties()) // use cache as per advice in https://github.com/gsuitedevs/apps-script-oauth2
      .setCache(CacheService.getUserCache())
      .setTokenUrl('https://oauth.oclc.org/token')      // Set the endpoint URLs.

      // Set the client ID and secret.
      .setClientId(myKey)
      .setClientSecret(mySecret)

      // Sets the custom grant type to use.
      .setGrantType('client_credentials')
      .setScope(scope)
      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties());
      
} // end get APIService

function getDiscoveryApiService() {  //https://github.com/gsuitedevs/apps-script-oauth2/blob/master/samples/TwitterAppOnly.gs
                            //https://github.com/gsuitedevs/apps-script-oauth2

   
    OCLCurl = 'https://americas.discovery.api.oclc.org/worldcat/search/v2/' ;  
    scope = "wcapi" ;  //scope = ['wcapi:view_retained_holdings', 'wcapi:view_summary_holdings', 'wcapi:view_institution_holdings', 'wcapi:view_holdings' ] 
/* wondering if at some point will need finer scopes, here's an example. The Python version did need finer scopes
      .setScope([
        "https://www.googleapis.com/auth/spreadsheets.currentonly",
        "https://www.googleapis.com/auth/script.external_request",
        "https://www.googleapis.com/auth/cloud-platform"
      ]);
*/
    myKey = PropertiesService.getUserProperties().getProperty('apiKey')
    mySecret = PropertiesService.getUserProperties().getProperty('apiSecret') 
 
  //myKey = 'testingForFailure'
  //mySecret = 'testingForFailure'
  Logger.log(PropertiesService.getUserProperties().getKeys());

  return OAuth2.createService('WorldCat Discovery API')
      .setPropertyStore(PropertiesService.getUserProperties()) // use cache as per advice in https://github.com/gsuitedevs/apps-script-oauth2
      .setCache(CacheService.getUserCache())
      .setTokenUrl('https://oauth.oclc.org/token')      // Set the endpoint URLs.

      // Set the client ID and secret.
      .setClientId(myKey)
      .setClientSecret(mySecret)

      // Sets the custom grant type to use.
      .setGrantType('client_credentials')
      .setScope(scope)
      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties());
      
} // end get DiscoveryAPIService

//===================================================================================================
function getMetadataApiService() {  //https://github.com/gsuitedevs/apps-script-oauth2/blob/master/samples/TwitterAppOnly.gs
                            //https://github.com/gsuitedevs/apps-script-oauth2
    OCLCurl = "https://metadata.api.oclc.org/worldcat";
    scope = "WorldCatMetadataAPI"
    myKey = PropertiesService.getUserProperties().getProperty('mapiKey')
    mySecret = PropertiesService.getUserProperties().getProperty('mapiSecret')  
   
  //myKey = 'testingForFailure'
  //mySecret = 'testingForFailure'
  Logger.log(PropertiesService.getUserProperties().getKeys());

 return OAuth2.createService('WorldCat Metadata API')
      .setPropertyStore(PropertiesService.getUserProperties()) // use cache as per advice in https://github.com/gsuitedevs/apps-script-oauth2
      .setCache(CacheService.getUserCache())
      .setTokenUrl('https://oauth.oclc.org/token')      // Set the endpoint URLs.

      // Set the client ID and secret.
      .setClientId(myKey)
      .setClientSecret(mySecret)

      // Sets the custom grant type to use.
      .setGrantType('client_credentials')
      .setScope(scope)
      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties());
      
} // end get MetaAPIService

//===================================================================================================
function checkForData(dataSheetToCheck, columns, startline, endline) { // see if any data exists in ranges that will be overwritten
  let searchRange =[] ;
  for (letter of columns) {
      searchRange.push(letter + startline + ':' + letter + endline) ;
  }
  var testValues = dataSheetToCheck.getRangeList(searchRange).getRanges().map(range => [range.getValues()]);  
  var flattened = testValues.flat(2) // flatten to an array to test
  if (flattened.some(element => element != "")) { // if there is data in any of the cells
    return 1; // true, yes there is data
  } else {
    return 0 ; // false, no data in cells
  } // end if data in flattened - not all cells are blank
} // end checkForData
//===================================================================================================    
function stopLookup() {
  PropertiesService.getUserProperties().setProperty("cancelscript", "yes");
  cancel = "YES" // change global cancel variable - this didn't seem to work
}
//===================================================================================================    
function reset() { // Reset the authorization state, so that it can be re-tested.
  var service = getAPIService(); // https://github.com/googleworkspace/apps-script-oauth2
  service.reset();
}
//===================================================================================================    
function getPercentDone() { // never got this working - if any has ideas would love to hear them
  //ui.alert("PD Property: " + PropertiesService.getUserProperties().getProperty('percentDone')); // this seems to always be 100 when we get here
  return PropertiesService.getUserProperties().getProperty('percentDone');
  //return percentDone ;
}
