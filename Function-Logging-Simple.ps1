function Write-Log
{
	<#
	.SYNOPSIS
		This function creates or appends a line to a log file.

	.PARAMETER  Message
		The message parameter is the log message you'd like to record to the log file.

	.EXAMPLE
		PS C:\> Write-Log -Message 'Value1'
		This example shows how to call the Write-Log function with named parameters.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)]
		[string]$Message
	)
	
	try
	{
		$DateTime = Get-Date -Format ‘MM-dd-yy HH:mm:ss’
		$Invocation = "$($MyInvocation.MyCommand.Source | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)"
		Add-Content -Value "$DateTime - $Invocation - $Message" -Path "$([environment]::GetEnvironmentVariable('TEMP', 'Machine'))\ScriptLog.log"
		Write-Host $Message
	}
	catch
	{
		Write-Error $_.Exception.Message
	}
}
