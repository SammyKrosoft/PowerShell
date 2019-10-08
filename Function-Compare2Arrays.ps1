Function Compare2Arrays {
    [CmdLetBinding()]
    Param(
        [Parameter(Position = 1)][array]$ReqFields,
        [Parameter(Position = 2)][array]$FieldsToCompareToReqFields
    )
    #comparing each formfield in the doc with the form field names defined to check if no one is missing
    LogYellow "There are $($ReqFields.count) text fields to check, the document contains $($FieldsToCompareToReqFields.count) Text Formfields" -b blue

    If ($($ReqFields.count) -ne $($FieldsToCompareToReqFields.count)){
    LogBlue "There is mismatch in the number of fields" -B yellow
    }

    $NbMatches = 0
    $MissingField = @()
    $found = $false
    $AtLeastOneMissing = $false
    Foreach ($chkitem in $ReqFields) {
        LogBlue "Checking $chkitem"
        Foreach ($docitem in $FieldsToCompareToReqFields){
            If ($chkItem -eq $($DocItem.'FormField Name')){
                LogGreen "This Field is in the Doc !"
                $found = $true
                $NbMatches += 1
            }
        }
        If (-not $found){
            LogMag "$chkitem not found in the document ..."
            $MissingField += $chkitem
            $AtLeastOneMissing = $True
        }
        $found = $false
    }

    If ($AtLeastOneMissing){
        Write-Host "At least one field is missing in the Doc" -ForegroundColor red -BackgroundColor yellow
        Write-Host "There are $($MissingField.count) fields missing in the doc"
        $MissingField | Out-Host
        Return $false
    } Else {
        Write-Host "All fields there !"
        Return $True
    }
}
