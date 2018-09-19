Function Get-E2016ReportValues {
    <#
    .SYNOPSIS
        Get Excel table data to be updated in the Word document

    .DESCRIPTION
        Get Excel table data to be updated in the Word document

    .PARAMETER FirstNumber
        Just a parameter sample...

    .PARAMETER CheckVersion
        This parameter will just dump the script current version.

    .INPUTS
        None.

    .OUTPUTS
        CSV file

    .EXAMPLE
    .\Do-Something.ps1
    This will launch the script and do someting

    .EXAMPLE
    .\Do-Something.ps1 -CheckVersion
    This will dump the script name and current version like :
    SCRIPT NAME : Do-Something.ps1
    VERSION : v1.0

    .NOTES
    None

    .LINK
        https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

    .LINK
        https://github.com/SammyKrosoft
    #>
    [CmdLetBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = "NormalRun")][string]$Department = "SSC",
        [Parameter(Mandatory = $false, Position = 3, ParameterSetName = "CheckOnly")][switch]$CheckVersion
    )

    <# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
    #Initializing a $Stopwatch variable to use to measure script execution
    $stopwatch = [system.diagnostics.stopwatch]::StartNew()
    #Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
    # and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
    $DebugPreference = "Continue"
    # Set Error Action to your needs
    $ErrorActionPreference = "SilentlyContinue"
    #Script Version
    $ScriptVersion = "0.1"
    <# Version changes
    v0.1 : first script version
    v0.1 -> v0.5 : 
    #>
    $ScriptName = $MyInvocation.MyCommand.Name
    If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
    # Log or report file definition
    # NOTE: use $PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
    $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
    $OutputReport = "$ScriptPath\$($ScriptName)_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
    # Other Option for Log or report file definition (use one of these)
    $ScriptLog = "$ScriptPath\$($ScriptName)-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
    <# ---------------------------- /SCRIPT_HEADER ---------------------------- #>

    cls

    $ExcelInput = "C:\Users\SammyKrosoft\OneDrive\_Boulot\How-To Procedures\Exchange 2016 docs\E2016Test.xlsx"
    $FileExists = Test-Path $ExcelInput

    If ($FileExists) {
        Write-Host "Excel file exists, continuing..."
    } Else {
        Write-Host "Excel file does not exist, existing..."
        Exit
    }

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    $Workbook = $Excel.Workbooks.Open($ExcelInput)
    $WorkSheet = $Workbook.Worksheets.item("Justice")
    $Worksheet.Activate()
    $WSTable = $Worksheet.ListObjects.Item(1)

    $WSTableRows = $WSTable.ListRows

    $WSTableRows.Count
    #$Row = $WSTableRows[1]
    #$RowVals = $Row.Range

    $WholeInputCollection = @()
    ForEach ($Row in $WSTableRows){
        $ValTrio = @() #Init & Re-init variable as we just want to store the values from each Row
        # there will be 3 columns that is 3 values for each Row
        Foreach ($Val in $($Row.Range)){
            #Write-Host $Val.Text
            $ValTrio += $Val.Text
        }
        #Write-Host "Trio is : $($ValTrio[0]),$($ValTrio[1]),$($ValTrio[2]) "
        $CustomObj = [PSCustomObject]@{
            Description = $($ValTrio[0])
            Value = $($ValTrio[1])
            BookMark = $($ValTrio[2])
        }
        $WholeInputCollection += $CustomObj
    }

    Return $WholeInputCollection

    Write-Host "Closing workbook..." -ForegroundColor Green
    $Workbook.Close()
    Write-Host "Releasing Workbook Com Object..." -ForegroundColor Green
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
    Write-Host "Closing Excel..." -ForegroundColor Green
    $Excel.Quit()
    Write-Host "Releasing Excel Com Object..." -ForegroundColor Green
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    Write-Host "Cleaning Excel variable..." -ForegroundColor Green
    Remove-Variable excel
    Write-Host "Garbage Collection..." -ForegroundColor Green
    [System.GC]::Collect()
    Write-Host "WaitForPendingFinalizers..." -ForegroundColor Green
    [System.GC]::WaitForPendingFinalizers()

    Return $WholeInputCollection

    <# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
    #Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
    $stopwatch.Stop()
    $msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
    Write-Host $msg
    $msg = $null
    $StopWatch = $null
    <# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
}