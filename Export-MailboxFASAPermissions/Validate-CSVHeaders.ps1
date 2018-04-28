cls

Function ValidateHeadersFromCSV {
    Param(
        [Parameter(Mandatory=$true, Position = 0, ParameterSetName = "NormalRun")][string]$FilePathAndName,
        [Parameter(Mandatory =$true, Position = 1, ParameterSetName = "NormalRun")][string[]]$CSVFilerequiredHeaders
    )
    $DuplicateHeader = $false
    $MissingHeader = $false
    If (!(Test-Path $FilePathAndName)){
        $MsgErrFileNotFound = "The file $InputFile is incorrect or doesn't exist ... please try again with another file or the correct path."
        Write-Host $MsgErrFileNotFound -BackgroundColor Yellow -ForegroundColor Red
        Return $false
    } Else {
        #Get the first line of the CSV file => THIS is what we will validate
        [string[]]$HeadersFromFile = (Get-content -Path $FilePathAndName | Select -first 1).Split(",")
        # $CSVFilerequiredHeaders
        # $CSVFilerequiredHeaders.count
        # $HeadersFromFile;
        # $HeadersFromFile.count

        # exit
        #Putting message in a variable to do localization in the future (French + English)
        $msgValidatingCSVHeader = "Validating the CSV headers..."
        Write-host $msgValidatingCSVHeader -BackgroundColor yellow -ForegroundColor blue
        #Parsing CSV required file headers coming from parameter
        #3 cases : 1 matching header in the file for each required header, 1 of the headers is missing, or we have duplicate headers 
        Foreach ($RequiredHeader in $CSVFilerequiredHeaders) {
            Write-Host "Checking $RequiredHeader" -BackgroundColor Green
            #Header counter to identify whether we have no matches, one match, or more than one
            $HeaderMatch = 0
            #We compare each CSV required header with each header of the file -> 3 cases : 1 match (wanted), 0 matches (CSV file not valid) or more than 1 matches (duplicates in CSV headers, CSV File not valid) 
            Foreach ($FileHeader in $HeadersFromFile){
                if($($FileHeader.Trim()) -eq $RequiredHeader -or $($FileHeader.Trim()) -eq """$RequiredHeader"""){$HeaderMatch++}
            }
            If ($HeaderMatch -eq 1){
                $msgFound1Match = "Ok"
                Write-Host $msgFound1Match -BackgroundColor green -ForegroundColor black
            } ElseIf($headerMatch -eq 0) {
                Write-Host "$RequiredHeader not found in file ! Please correct our CSV or select another CSV file. Exiting..."
                $MissingHeader = $true
                [array]$MissingHeaderDetails += $RequiredHeader
            } Else {
                Write-Host "Cannot have more than 1 header named $RequiredHeader - please correct your CSV or select another CSV. Exiting..."
                $DuplicateHeader = $true
                [array]$DuplicateHeaderDetails += $RequiredHeader
            }
        }
    }
    If ($Missingheader -or $DuplicateHeader){
        write-host "MISSING HEADER   : $MissingHeader"
        Write-host "Duplicate Header : $DuplicateHeader"
        If ($MissingHeader){
            $msgMissingHeader = "Missing Headers in file : $($MissingHeaderDetails -join ", "), please use a CSV file with proper headers"
            Write-Host $msgMissingHeader -BackgroundColor yellow -ForegroundColor red
        }
        If ($DuplicateHeader){
            $msgDuplicateHeader = "Duplicate Headers in file : $($DuplicateHeaderDetails -join ", "), please use a CSV file with proper headers"
            Write-Host $msgDuplicateHeader -BackgroundColor yellow -ForegroundColor red
        }
        return $False
    }Else{
        Return $True
    }
}

ValidateHeadersFromCSV -FilePathAndName ".\sample.csv" -CSVFilerequiredHeaders "PrimarySMTPAddress", "SendAsPErmissions", "FullAccessPermissions", "SendOnBehalfPermissions"
