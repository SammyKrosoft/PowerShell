<#
# Script: Get-CounterStats
# Author: Prashanth and Praveen
# Comments: This script will collect the specific counters value from the multiple target machines/servers 
which will be used to analayze the performance of target servers.
#>

#Define Input and output filepath

$servers=get-content ".\servers.txt"
$outfile="C:\perfmon.csv"

################################################################################################################

#Function to have the customized output in CSV format
function Export-CsvFile {
[CmdletBinding(DefaultParameterSetName='Delimiter', SupportsShouldProcess = $true, ConfirmImpact='Medium')]
param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)][System.Management.Automation.PSObject]${InputObject},
    [Parameter(Mandatory=$true, Position=0)][Alias('PSPath')][System.String]${Path},
    [Switch]${Append},
    [Switch]${Force},
    [Switch]${NoClobber},
    [ValidateSet('Unicode','UTF7','UTF8','ASCII','UTF32','BigEndianUnicode','Default','OEM')][System.String]${Encoding},
    [Parameter(ParameterSetName='Delimiter', Position=1)][ValidateNotNull()][System.Char]${Delimiter},
    [Parameter(ParameterSetName='UseCulture')][Switch]${UseCulture},
    [Alias('NTI')][Switch]${NoTypeInformation}
)

begin
{
    # This variable will tell us whether we actually need to append
    # to existing file
    $AppendMode = $false
    try {
    $outBuffer = $null
      if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer))
      {
      $PSBoundParameters['OutBuffer'] = 1
      }
  $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Export-Csv',
    [System.Management.Automation.CommandTypes]::Cmdlet)
        
        
                #String variable to become the target command line
                $scriptCmdPipeline = ''

                # Add new parameter handling
                #region Dmitry: Process and remove the Append parameter if it is present
                if ($Append) {
  
                                $PSBoundParameters.Remove('Append') | Out-Null
    
  if ($Path) {
   if (Test-Path $Path) {        
    # Need to construct new command line
    $AppendMode = $true
    
    if ($Encoding.Length -eq 0) {
     # ASCII is default encoding for Export-CSV
     $Encoding = 'ASCII'
    }
    
    # For Append we use ConvertTo-CSV instead of Export
    $scriptCmdPipeline += 'ConvertTo-Csv -NoTypeInformation '
    
    # Inherit other CSV convertion parameters
    if ( $UseCulture ) {
     $scriptCmdPipeline += ' -UseCulture '
    }
    if ( $Delimiter ) {
     $scriptCmdPipeline += " -Delimiter '$Delimiter' "
    } 
    
    # Skip the first line (the one with the property names) 
    $scriptCmdPipeline += ' | Foreach-Object {$start=$true}'
    $scriptCmdPipeline += '{if ($start) {$start=$false} else {$_}} '
    
    # Add file output
    $scriptCmdPipeline += " | Out-File -FilePath '$Path' -Encoding '$Encoding' -Append "
    
    if ($Force) {
     $scriptCmdPipeline += ' -Force'
    }

    if ($NoClobber) {
     $scriptCmdPipeline += ' -NoClobber'
    }   
   }
  }
} 
  

  
 $scriptCmd = {& $wrappedCmd @PSBoundParameters }

 if ( $AppendMode ) {
  # redefine command line
  $scriptCmd = $ExecutionContext.InvokeCommand.NewScriptBlock(
      $scriptCmdPipeline
    )
} else {
  # execute Export-CSV as we got it because
  # either -Append is missing or file does not exist
  $scriptCmd = $ExecutionContext.InvokeCommand.NewScriptBlock(
      [string]$scriptCmd
    )
}

# standard pipeline initialization
$steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
$steppablePipeline.Begin($PSCmdlet)

 } catch {
   throw
}
    
}

process
{
  try {
      $steppablePipeline.Process($_)
  } catch {
      throw
  }
}

end
{
  try {
      $steppablePipeline.End()
  } catch {
      throw
  }
}

}

################################################################################################################


#Actual script starts here 

function Global:Convert-HString {      
[CmdletBinding()]            
 Param             
   (
    [Parameter(Mandatory=$false,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true)]
    [String]$HString
   )#End Param

Begin 
{
    Write-Verbose "Converting Here-String to Array"
}#Begin
Process 
{
    $HString -split "`n" | ForEach-Object {
    
        $ComputerName = $_.trim()
        if ($ComputerName -notmatch "#")
            {
                $ComputerName
            }    
        
        
        }
}#Process
End 
{
    # Nothing to do here.
}#End

}#Convert-HString


#Performance counters declaration

function Get-CounterStats { 
param 
    ( 
    [String]$ComputerName = $ENV:ComputerName
    
    ) 

$Object =@()


$Counter = @" 
Processor(_total)\% processor time 
\MSExchange RpcClientAccess\RPC Averaged Latency
\MSExchange RpcClientAccess\RPC Requests
Memory\Available MBytes 
PhysicalDisk(*)\Avg. Disk sec/Transfer 
Network Interface(*)\Bytes Total/sec
"@ 

        (Get-Counter -ComputerName $ComputerName -Counter (Convert-HString -HString $Counter)).counterSamples |  
        ForEach-Object { 
        $path = $_.path 
        New-Object PSObject -Property @{
        computerName=$ComputerName;
        Counter        = ($path  -split "\\")[-2,-1] -join "-" ;
        Item        = $_.InstanceName ;
        Value = [Math]::Round($_.CookedValue,2) 
        datetime=(Get-Date -format "yyyy-MM-d hh:mm:ss")
        } 
        
        }
     
   
} 

#Collecting counter information for target servers

foreach($server in $Servers)
{
$d=Get-CounterStats -ComputerName $server |Select-Object computerName,Counter,Item,Value,datetime
$d |Export-CsvFile $outfile  -Append -NoTypeInformation

}

#End of Script