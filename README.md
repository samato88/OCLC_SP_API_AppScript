# OCLC_SP_API_AppScript

This project is a Google Sheets Add-on for use by Shared Print programs. It uses the OCLC Worldcat Search API or the WorldCat Metadata API to look for Shared Print committments registered with OCLC. It also can be used to search HathiTrust and Internet Archives by OCLC number and return links if titles are digitally availalbe in those services. 

Prerequisits:
Sheet must have OCLC numbers in column A, and the first row of the sheet is reserved for headers.
API Key and Secret for either the Searh or Metadata API

Caveats:
Overwrites columns B-K, depending on what search options selected, you will receive a warning about this
App Script has a 6 minute timeout for free accounts, 30 minute timeout for corporate accounts. This script will currently silently fail upon timeout, so do small batches!!
Hoping to fix this up a bit in the next version. 

v.18

---
Many thanks to [MatchMARC](https://github.com/suranofsky/tech-services-g-sheets-addon) for providing an excellent example. Much code borrowed from them.  

Comments, fixes, additions much appreciated.
