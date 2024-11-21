function getSPPHoldings(oclc, SPP) {
    // https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app
    let ui = SpreadsheetApp.getUi();
    let numberSPPHoldings = 0 ;
    let retainedBy = [] ;
    let currentOCLC = "" ;

    if (!apiService.hasAccess()) { // check if expired.  not sure this is the right way to do that
      getApiService();
    }
   // should try/except this
     if (apiService.hasAccess()) { //      
      var url = OCLCurl + 'retained-holdings?oclcNumber=' + oclc + '&spProgram=' + SPP ;
      var response = UrlFetchApp.fetch(url, {
        headers: {
           Authorization: 'Bearer ' + apiService.getAccessToken()
        },
        validateHttpsCertificates: false,
        muteHttpExceptions: true
       });

     ////CHECK RESPONSE HEADER NOT 403 or 404  //Logger.log(response.getHeaders()); //Logger.log(response.getContentText());
     if(response.getResponseCode() != 200) { // not 200
          numberSPPHoldings = "";
          retainedBy = [];
          currentOCLC = "Server Error: " + response.getResponseCode() ;
     } else { //valid response
          var results = JSON.parse(response.getContentText());

          if (results.numberOfHoldings) {// NEED TO ADD A CHECK IN HERE - SOMETIMES "numberOfHoldings": 1 but no detailed holdings (oclc API glitch)
            /*
{detailedHoldings=[{format=zu, lhrLastUpdated=20210215, sharedPrintCommitments=[{actionNote=committed to retain, dateOfAction=20160630, commitmentExpirationDate=20310630, authorization=EAST, institution=MBU}], lhrControlNumber=352397802, lhrDateEntered=20210215, location={sublocationCollection=BOSS, holdingLocation=BOS}, hasSharedPrintCommitment=Y, summary=Local Holdings Available., oclcNumber=123456}], numberOfHoldings=1.0} 
*/
            //numberSPPHoldings = results.numberOfHoldings; // will be wrong if multple lhrs for same symbol
            for (lib in results.detailedHoldings) {

             /*Occasionaly see results with no holdingLocaion and "lhrControlNumber": "UnavailableLHR352140713",*/
              if (!results.detailedHoldings[lib].lhrControlNumber.startsWith("Unavailable")) {  
                retainedBy.push(results.detailedHoldings[lib].location.holdingLocation); //  holdings symbol
              } else if (numberSPPHoldings > 0) { // else you should decrement the number of holdings, as one of them is Unavailable/bad
                --numberSPPHoldings
              }
            } // end foreach results.detailedHoldings (LHRs)
            
            //ui.alert("SPP: " + numberSPPHoldings); 
            //Logger.log("results.detailedHoldings[0].oclcNumber: "+ results.detailedHoldings[0].oclcNumber);
            // NEED TO ADD A CHECK IN HERE - SOMETIMES "numberOfHoldings": 1 but no detailed holdings (oclc API glitch)
            currentOCLC = results.detailedHoldings[0].oclcNumber
          }
     } // end else valid response
    } // end apiService.hasAccess //Logger.log(apiService.getLastError());
    
    ////dedup retainedBy - some ocns have muliple LHRs for the same title, e.g. volume, e.g. 47118284 symbol PBU
    retainedBy = Array.from(new Set(retainedBy)) 
    numberSPPHoldings = retainedBy.length
    return [numberSPPHoldings, retainedBy, currentOCLC] ;
} // end getSPPHoldings

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

     if(response2.getResponseCode() != 200) {
       return  response2.getResponseCode() ;
     } else {
       var result2 = JSON.parse(response2.getContentText());
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
       //Logger.log(apiService.getLastError());
       return   "Authorization failed"        
  } // end else doesn't have access
    return [holdingsCount, otitle, CurrentOCN, Merged];
  
  } // end get worldcat holdings
