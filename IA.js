function getIAHoldingsMerged(oclc, merged) {
  var iaurl = "https://archive.org/advancedsearch.php?q=oclc-id:" + oclc + "&fl[]=identifier&output=json"
  var iaholdings = ""
  
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
    iaholdings = err
  }
 
  return iaholdings ;
}// end getIAHoldingsMerged
