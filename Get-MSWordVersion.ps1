Function Get-PowerShellVersion {$PSVersion = $PSVersionTable.PSVersion.Major;Return $PSVersion}

Function Get-MSWordVersion {

    $MSWord = New-Object -ComObject Word.Application

    $MSWordversion = $MSWord.Version

    #Quitting Word gracefully, freeing the COM object and cleaning the variable
    $MSWord.Quit()
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSword)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable MSword

    return $MSWordVersion
}

#$MSWordVersion = Get-MSWordVersion
$MSWordVersion = 12.1
Write-host "MSWord version installed is : $MSWordVersion"

If ($MSWordVersion -ge 15){
    Write-host "MSWord version is $MSWordVErsion and is greater than 2013, we're good to go !"
} Else {
    Write-Host "Alas, MSWord version is $MSWordversion and is older than 2013 ... exiting"
    exit
}
