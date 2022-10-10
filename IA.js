function getIAHoldingsMerged(oclc, merged) {
  var iaurl = "https://archive.org/advancedsearch.php?q=oclc-id:" + oclc + "&fl[]=identifier&output=json"
  var iaholdings = ""
  var count = 0

  try {
   var r =  UrlFetchApp.fetch(iaurl);
   var response = JSON.parse(r.getContentText()) ;
   if (response) {
    if  (response.response.numFound == 0 ) {
      if (typeof merged !== 'undefined') {

       if (merged.length > 0) { // there are alt oclc numbers
        for (altNumb in merged) {
          iaurl = "https://archive.org/advancedsearch.php?q=oclc-id%3A" +  merged[altNumb] +  "&fl%5B%5D=licenseurl,identifier&sort%5B%5D=&sort%5B%5D=&sort%5B%5D=&rows=50&page=1&output=json"
           response = JSON.parse(UrlFetchApp.fetch(iaurl,HTTP_OPTIONS).getContentText());
           if (response.response.numFound > 0) { 
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

      // retry address unavailable errors 3 times, then move on
      //Exception: Address unavailable: https://archive.org/advancedsearch.php?q=oclc-id:1606080&fl[]=identifier&output=json
      if(err.message.startsWith("Address unavailable: https://archive.org") && count < 3) {
         Utilities.sleep(300) // brief pause to give api a break
         Logger.log("IA Error on %s", err.message)
         ++count
         getIAHoldingsMerged(oclc, merged)
      } else {
    //iaholdings = err
    iaholdings = "ERROR retrieving data"
      }
  }
 
  return iaholdings ;
}// end getIAHoldingsMerged
