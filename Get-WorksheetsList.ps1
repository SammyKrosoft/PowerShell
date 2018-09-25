Function LogMag {
    param(
        [Parameter(Mandatory = $false, Position = 1)][string]$message,
        [Parameter(Mandatory = $false)][string]$b = "black"
    )
    Write-Host $message -F Magenta -b $b
}

Function LogGreen {
    param(
        [Parameter(Mandatory = $false, Position = 1)][string]$message,
        [Parameter(Mandatory = $false)][string]$b = "black"
    )
    Write-Host $message -F Green -b $b
}

Function LogYellow {
    param(
        [Parameter(Mandatory = $false, Position = 1)][string]$message,
        [Parameter(Mandatory = $false)][string]$b = "black"
    )

    Write-Host $message -F Yellow -b $b
}

Function LogBlue {
    param(
        [Parameter(Mandatory = $false, Position = 1)][string]$message,
        [Parameter(Mandatory = $false)][string]$b = "black"
    )
    Write-Host $message -F Blue -b $b
}

Function Get-ExcelWorkSheetsNames {
    <#
    .SYNOPSIS
        Get Excel worksheets names

    .DESCRIPTION
        Get Excel worksheets names to populate list, dropdown or just validate...
    #>
    [CmdLetBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "NormalRun")][string]$ExcelInput,
        [Parameter(Mandatory = $false, Position = 2, ParameterSetName = "CheckOnly")][switch]$CheckVersion
    )

    <# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
    #Initializing a $Stopwatch variable to use to measure script execution
    $stopwatch2 = [system.diagnostics.stopwatch]::StartNew()
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
    # NOTE: use $PSScriptRoot in Powershell 3.0 and later or use $LocalScriptExecPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
    <# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
    
    If (-Not $ExcelInput){
        $ExcelInput = ".\E2016Test.xlsx"
        LogMag "No Excel input file specified ... using default:" -b Yellow
        Write-Host $ExcelInput
    } Else {
        "Excel input file specified : $ExcelInput. Continuing ..." | Out-Host
    }

    $FullXLFilePath = $PSScriptRoot + "\" + $ExcelInput
    $FileExists = Test-Path $FullXLFilePath

    LogBlue $FullXLFilePath
    If ($FileExists) {
        LogGreen "Excel file exists, continuing..."
    } Else {
        $msg = "Excel file does not exist, exiting..."
        LogMag $msg
        [System.Windows.MessageBox]::Show($msg,"File not found","Ok","Error")
        Return "FileNotFound"#Trying to return to GUI...
    }

    LogGreen "Opening a new Excel instance..."
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    LogGreen "Opening Excel workbook..."
    $Workbook = $Excel.Workbooks.Open($FullXLFilePath)
    $WorkSheetsObjectsList = $Workbook.Worksheets

    LogBlue "Getting all worksheets (aka ""Tabs"") names..."
    $WorkSheetsList = @()
    Foreach ($Worksheet in $WorkSheetsObjectsList){
        $WorkSheetsList += $($Worksheet.Name)
    }

    Write-Host "Closing workbook..." -ForegroundColor Green
    $Workbook.Close()
    Write-Host "Releasing Workbook Com Object..." -ForegroundColor Green
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    Write-Host "Closing Excel..." -ForegroundColor Green
    $Excel.Quit()
    Write-Host "Releasing Excel Com Object..." -ForegroundColor Green
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Cleaning Excel variable..." -ForegroundColor Green
    Remove-Variable excel
    Write-Host "Garbage Collection..." -ForegroundColor Green
    [System.GC]::Collect()
    Write-Host "WaitForPendingFinalizers..." -ForegroundColor Green
    [System.GC]::WaitForPendingFinalizers()

    <# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
    #Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
    $stopwatch2.Stop()
    $msg = "`n`nThe script took $([math]::round($($StopWatch2.Elapsed.TotalSeconds),2)) seconds to execute..."
    Write-Host $msg
    $msg = $null
    $StopWatch2 = $null
    <# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>

    Return $WorkSheetsList
}

$all = Get-ExcelWorkSheetsNames
$all | Out-Host
