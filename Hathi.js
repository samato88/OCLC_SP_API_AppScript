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

    if (typeof response.items[0] == "undefined") {
      if (merged && merged.length > 0) { // there are alt oclc numbers
        for (altNumb in merged) {
          //Logger.log("altNumb: " + altNumb + " Merged: " + merged + " typeof: " + typeof merged + " len: " + merged.length);
          //Logger.log(merged[altNumb]);
            hathiurl = "http://catalog.hathitrust.org/api/volumes/brief/oclc/" +  merged[altNumb] +  ".json"
            response = JSON.parse(UrlFetchApp.fetch(hathiurl,HTTP_OPTIONS).getContentText());
              if (typeof response.items[0] != "undefined") {
                //Logger.log("type of response alt: " + typeof response.items[0]) ;
                break;
              } // if got response break to outer
       } // end for altNumb in merged
     } // end if merged numbers
    }// end if undefined response
  
  
    if   (typeof response.items[0] != "undefined") { // make sure you got something
      r = response.items[0].usRightsString ;
      htid = response.items[0].htid
      //  pubdate   = response.records[Object.keys(response.records)[0]].publishDates;   
      htitle = response.records[Object.keys(response.records)[0]].titles[0];
      recordurl = "https://catalog.hathitrust.org/Record/" + response.items[0].fromRecord  ;
      rights = '=hyperlink("' + recordurl +  '","' + r + '")';    
    }

  } // end try
  catch(err) {
    console.log("Hathi Error: " + err)
    rights = err;
    htid = ""
    htitle = ""
  }
  return [rights, htid, htitle] ;
}// end get HathiHoldings  

