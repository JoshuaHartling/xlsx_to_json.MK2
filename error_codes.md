# Error Codes
* Code 1: device filter (device_list) is using improper syntax.  Should be in either string or list of string format.
* Code 2: there is two or more columns with duplicate headers.  The first entry in each column should be unique as this
references a unique metavariable.
    * Exception: empty cells or cell's with placeholder values (i.e., "N/A") can
    be duplicated as they are skipped anyways.
* Code 3: could not open/create worksheet for editing in "main-json2xlsx". 
* Code 7: user did not continue "main-json2xlsx" because ADOMs don't match.
* Code 8: user did not continue "main-json2xlsx" script because backup failed.
* Code 9: user did not continue "main-json2xlsx" script at initial prompt.
* Code 101: hostname column is missing.
* Code 102: VDOM column is missing.  