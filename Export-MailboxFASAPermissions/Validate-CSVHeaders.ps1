cls


Function ValidateHeadersFromCSV {
    Param{
        [Parameter(Mandatory=$true, Position = 0, ParameterSetName = "NormalRun")][string]$FilePathAndName,
        [Parameter(Mandatory =$true, Position = 1, ParameterSetName = "NormalRun")][string[]]$HeadersToValidate = @("Header1", "Header2", "Header3")
    }

    $HeadersFromFile = (Get-content -Path ".\sample.csv" | Select -first 1).Split(",")

    Foreach ($RequiredHeader in $CSVFilerequiredHeaders) {
        $HeaderMatch = 0
        Foreach ($FileHeader in $HeadersFromFile){if($FileHeader -eq $RequiredHeader){$HeaderMatch++}}
        If ($HeaderMatch = 1){
            Write-Host "Find 1 match in CSV Headers for $RequiredHeader => we're good for this one" 
        } ElseIf($headerMatch = 0) {
            Write-Host "$RequiredHeader not found in file ! Please correct our CSV or select another CSV file. Exiting..."
            $CSVHasRequiredHeaders = $false
            Exit
        } Else {
            Write-Host "Cannot have more than 1 header named $RequiredHeader - please correct your CSV or select another CSV. Exiting..."
            $CSVHasRequiredHeaders = $false
            exit
        }
    }
}