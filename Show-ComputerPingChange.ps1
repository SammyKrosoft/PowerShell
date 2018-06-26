<#
.SYNOPSIS
This script loops through a computer status, and display reachability status change.

.PARAMETER Computer
The Computer to monitor reachability

.PARAMETER FileName
If specified, logs the ping change into a file

.INPUTS
None. You cannot pipe objects to that script.

.OUTPUTS


.EXAMPLE
C:\PS> .\Full-Name.ps1
Your full name is Merlin the Wizard

.EXAMPLE
C:\PS> .\Full-Name.ps1 -FirstName "Jane" -LastName "Doe"
Your full name is Jane Doe

.EXAMPLE
C:\PS> .\Full-Name.ps1 "Jane" "Doe"
Your full name is Jane Doe

.LINK
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
https://github.com/SammyKrosoft
#>

Function IsComputerReachable {
    [CmdLetBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $False, Position = 0, ParameterSetName = "NormalRun")][string]$ComputerName = $($ENV:ComputerName),
        [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "CheckOnly")][switch]$CheckVersion
    )

    If (Test-Connection $ComputerName -Count 1 -ErrorAction SilentlyContinue) {
        Return $True
    } Else {
        Return $False
    }
}

Function Show-ComputerStatusChange ($ComputerName,$SleepTime=2) {
    $FirstStatus = IsComputerReachable $ComputerName
    $CurrentTest = $FirstStatus
    $CurrentTime = $((Get-Date).ToString())

    Write-Host "Beginning..."
    Write-Host "Status => $CurrentTest at $CurrentTime"

    While ($True) {
        $LastTest = $CurrentTest
        $LastTime = $CurrentTime
        Sleep $SleepTime
        $CurrentTest = IsComputerReachable $ComputerName
        $CurrentTime = $((Get-Date).ToString())
        
        # write-host ("Last Test and time") + $LastTime + $LastTest
        # Write-Host ("Current Test and time") + $CurrentTime + $CurrentTest
        If ($CurrentTest -ne $LastTest){
            Write-Host "Status changed - Last status was = $LastTest, and it changed at $CurrentTime to => $CurrentTest" -BackgroundColor Yellow -ForegroundColor Blue
        }
    }
}

cls
Show-ComputerStatusChange -ComputerName E2010