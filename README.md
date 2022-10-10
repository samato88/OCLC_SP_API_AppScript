# OCLC_SP_API_AppScript

This project is a Google Sheets Add-on for use by Shared Print programs. It uses the OCLC Search API to look for Shared Print committments registered with OCLC. It also can be used to search HathiTrust and Internet Archives by OCLC number and return links if titles are digitally availalbe in those services. 

Prerequisits:
Sheet must have OCLC numbers in column A, and the first row of the sheet is reserved for headers.
API Key and Secret

Caveats:
Overwrites columns B-K, depending on what search options selected
App Script has a 6 minute timeout

v.13

---
Many thanks to [MatchMARC](https://github.com/suranofsky/tech-services-g-sheets-addon) for providing an excellent example. Much code borrowed from them.  

Comments, fixes, additions much appreciated.
