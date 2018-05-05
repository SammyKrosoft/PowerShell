cls


function Global:Convert-HString {
    [CmdletBinding()]
    Param ([Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)] [String]$HString)
    Begin
    {Write-Verbose "Converting Here-String to Array"}
    Process
    {
        $HString -split "`n" | ForEach-Object {
            $ComputerName = $_.trim()
            if ($ComputerName -notmatch "#")
            {$ComputerName}
        }
    }#Process
    End 
    {
        # Nothing to do here.
    }
}#Convert-HString


$Counter = @"
Processor(_total)\% processor time 
Memory\Available MBytes 
Network Interface(*)\Bytes Total/sec
"@ 

$counters = Convert-HString -hstring $counter

$Counters.GetType()

Write-host "`n"

Write-Host "Size of the array : $($Counters.count)"

For ($i=0;$i -le $($Counters.count-1);$i++){
    Write-Host "Counter $i  ->  $($Counters[$i])"
}
