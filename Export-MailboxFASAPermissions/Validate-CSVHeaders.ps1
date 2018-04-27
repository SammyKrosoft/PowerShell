[array]$CSVFileRequiredHeaders = "PrimarySMTPAddress", "SendAsPErmissions", "FullAccessPermissions", "SendOnBehalfPermissions"

$CheckCSVFile = import-csv .\sample.csv

Foreach ($Item in $CheckCSVFile){$CheckCSVFile.PrimarySMTPAddress}


