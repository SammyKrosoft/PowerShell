<#
    .SYNOPSIS
    Creates a HTML Report showing Sendas, Full Access and Send on Behalf Permission Information for Each Mailbox for your Exchange Organization, selected database or for a single user.
   
   	Serkan Varoglu
	
	http:\\Mshowto.org
	http:\\Get-Mailbox.org
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 1.1, 5 March 2012
	
    .DESCRIPTION
	
    Creates a HTML Report showing Sendas, Full Access and Send on Behalf Permission Information for Each Mailbox for your Exchange Organization, selected database or for a single user.
	By Default Inherited Send As permission and NT Authority\Self account will not be shown in the report unless you run the script with the parameters listed below.
	Also by default all mailboxes will be reported if you want to report a single database, you can use -database parameter to specify your database name or you can get the report for a single user.
	
	.PARAMETER HTMLReport
    Filename to write HTML Report to
	
	.PARAMETER Database
    By default this script will report all mailboxes. If you want to report mailboxes in a single database, you can use this parameter to input your database name.
	
	.PARAMETER Mailbox
    By default this script will report all mailboxes. If you want to report a single mailbox, you can use this parameter to input the mailbox you want to report.
	
	.SWITCH ShowInherited
	If ShowInherited is added as switch the report will show Inherited Sendas permissions for mailboxes as well.
	
	.SWITCH ShowSelf
	If ShowSelf is added as switch the report will show "NT Authority\Self" sendas permission for mailboxes as well.
	
	.EXAMPLE
    Generate the HTML report 
    .\Report-Permissions.ps1 -HTMLReport "C:\Users\SVaroglu\Desktop\MailboxPermissionReport.HTML"
	
#>

param 
( 
[Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Filename to write HTML report to. For Example: c:\DistGroupReport.html')][string]$HTMLReport,
[Parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='This switch will list Inherited Sendas and Full Access permissions as well')][switch]$ShowInherited,
[Parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='This switch will list NT Authority\Self Permission as well')][switch]$ShowSelf,
[Parameter(Position=3,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Choose a specific Database to Report')]$Database,
[Parameter(Position=4,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Choose a Mailbox to Report')]$Mailbox
)
$Watch = [System.Diagnostics.Stopwatch]::StartNew()
$WarningPreference="SilentlyContinue"
$ErrorActionPreference="SilentlyContinue"
$ShowInherited=$ShowInherited.IsPresent
$ShowSelf=$ShowSelf.IsPresent
$u=1
$s=0
$f=0
$b=0
$n=0
$nj=-1
$gj=-1
if (!$database){$dbnull=0}
if (!$mailbox){$mbnull=0}
if (($ShowSelf -like "false") -and ($ShowInherited -like "false"))
{$gentitle="Mailboxes With Custom Permissions"}
else
{$gentitle="Mailboxes"}
$gen="<table border=""1"" bordercolor=""#4384D3"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#4384D3"" align=""center""><th colspan=""5""><font color=""#FFFFFF"">$($gentitle)</font></th></tr><tr>"
$inh="<table border=""1"" bordercolor=""#4384D3"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#4384D3"" align=""center""><th colspan=""5""><font color=""#FFFFFF"">Mailboxes With Only Inherited Permissions</font></th></tr><tr>"
function _Progress
{
    param($PercentComplete,$Status)
    Write-Progress -id 1 -activity "Report for Mailboxes" -status $Status -percentComplete ($PercentComplete)
}
_Progress (($u*100)/100) "Collecting Mailbox Information"
if(!$database -and !$mailbox)
{
$mailboxes=get-mailbox -resultsize unlimited | Sort-Object name
}
elseif ($database -and !$mailbox)
{
$mailboxes=get-mailbox -database $database -resultsize unlimited | Sort-Object name
}
elseif (!$database -and $mailbox)
{
$mailboxes=get-mailbox $mailbox
}
else
{
Write-Host -ForegroundColor Cyan "Please choose database or single mailbox. Both Parameters can not be used at the same time. Ended without compiling a report."
exit}
$mcount=($mailboxes | measure-object).count
if ($mcount -eq 0){
Write-Host -ForegroundColor Cyan "No Mailbox Found. Ended without compiling a report. Please Check Your Input."
exit}
foreach ($mailbox in $mailboxes)
		{
_Progress (($u*95)/$mcount) "Processing $mailbox, $($u) of $($mcount) Mailboxes."
			$SenderBody=""
			$FullBody=""
			$BehalfBody=""
			$sendbehalfs=Get-Mailbox $mailbox | select-object -expand grantsendonbehalfto | select-object -expand rdn | Sort-Object Unescapedname
			if (($ShowSelf -like "true") -and ($ShowInherited -like "true"))
			{
			$senders=Get-ADPermission $mailbox.identity | ?{($_.extendedrights -like "*send-as*")} | Sort-Object name
			$fullsenders=Get-Mailbox $mailbox | Get-MailboxPermission | ?{($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*")}
			}
			elseif (($ShowSelf -like "false") -and ($ShowInherited -like "true"))
			{
			$senders=Get-ADPermission $mailbox.identity | ?{($_.extendedrights -like "*send-as*") -and ($_.User -notlike "NT Authority\self")} | Sort-Object name
			$fullsenders=Get-Mailbox $mailbox | Get-MailboxPermission | ?{($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\self*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*")}
			}
			elseif (($ShowSelf -like "true") -and ($ShowInherited -like "false"))
			{
			$senders=Get-ADPermission $mailbox.identity | ?{($_.extendedrights -like "*send-as*") -and ($_.isinherited -like "false")} | Sort-Object name
			$fullsenders=Get-Mailbox $mailbox | Get-MailboxPermission | ?{($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") -and ($_.IsInherited -like "false")}
			}
			else
			{
			$senders=Get-ADPermission $mailbox.identity | ?{($_.extendedrights -like "*send-as*") -and ($_.isinherited -like "false") -and ($_.User -notlike "NT Authority\self")} | Sort-Object name
			$fullsenders=Get-Mailbox $mailbox | Get-MailboxPermission | ?{($_.AccessRights -like "*fullaccess*") -and ($_.User -notlike "*nt authority\self*") -and ($_.User -notlike "*nt authority\system*") -and ($_.User -notlike "*Exchange Trusted Subsystem*") -and ($_.User -notlike "*Exchange Servers*") -and ($_.IsInherited -like "false")}
			}
			if (!$senders -and !$fullsenders -and !$sendbehalfs)
				{
				$n++
				if ($nj -eq 4)
				{
					$inh+="</tr><tr><td>$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</td>"
					$nj=0
				}
				else{$inh+="<td>$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</td>"
				$nj++}
				}
			else 
				{
						if ($gj -eq 4)
						{
							$gen+="</tr><tr><td><a href=""#$($mailbox.Name)"">$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</a></td>"
							$gj=0
						}
						else
						{
							$gen+="<td><a href=""#$($mailbox.Name)"">$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</a></td>"
							$gj++
						}
					$MailboxTable="<table border=""1"" bordercolor=""#1F497B"" cellpadding=""3"" style=""font-size:8pt;font-family:Arial,sans-serif"" width=""100%""><tr bgcolor=""#1F497B"" align=""center""><th colspan=""3"" ><font color=""#FFFFFF""><a name=""$($mailbox.Name)"">$($mailbox.Name) ( $($mailbox.primarysmtpaddress) )</font></a></th></tr><tr>"
					if (!$senders)
						{
						$SenderBody+="<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#696969"" align=""center"" valign=""top""><font color=""#FFFF99"">No User has Send As Permission On This Mailbox</font></td></table></td>"
						}
					else
						{
						$SenderBody+="<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><tr><td colspan=""3"" align=""center"" valign=""top"" bgcolor=""#696969""><font color=""#FFFFE0"">Send-As Permissions</font></td></tr><tr bgcolor=""#878787"" align=""center"">
						<td nowrap=""nowrap""><font color=""#FFFFFF"">Send as Permission Owner</font></td>
						<td nowrap=""nowrap""><font color=""#FFFFFF"">Deny</font></td>
						<td nowrap=""nowrap""><font color=""#FFFFFF"">Inherited</font></td>
						</tr>"
						foreach ($sender in $senders)
							{
								if (0,2,4,6,8 -contains "$sj"[-1]-48)
								{
								$bgcolor="'#E8E8E8'"
								}
								else
								{
								$bgcolor="'#C8C8C8'"
								}
									$SenderBody+="<tr align=""center"" bgcolor=$($bgcolor)>"
									$SenderBody+="<td><font color=""#003333"">$($sender.user)</font></td>"
									if ($sender.deny -like "true"){$font="red"}else{$font="'#000000'"}
									$SenderBody+="<td><font color=$font>$($sender.deny)</font></td>"
									if ($sender.isinherited -like "false"){$font="red"}else{$font="'#000000'"}
									$SenderBody+="<td><font color=$font>$($sender.isinherited)</font></td>"
									$SenderBody+="</tr>"
								$sj++
							}
						$SenderBody+="</table></td>"
						$s++
						}
						
					if (!$fullsenders)
						{
						$FullBody+="<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#696969"" align=""center"" valign=""top""><font color=""#FFFF99"">No User has Full Access On This Mailbox</font></td></table></td>"
						}
					else
						{
						$FullBody+="<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><tr><td colspan=""3"" align=""center"" valign=""top"" bgcolor=""#696969""><font color=""#FFFFE0"">Full Access Permissions</font></td></tr><tr bgcolor=""#878787"" align=""center"">
						<td nowrap=""nowrap""><font color=""#FFFFFF"">Full Access Permission Owner</font></td>
						<td nowrap=""nowrap""><font color=""#FFFFFF"">Deny</font></td>
						<td nowrap=""nowrap""><font color=""#FFFFFF"">Inherited</font></td>
						</tr>"
						foreach ($fullsender in $fullsenders)
							{
								if (0,2,4,6,8 -contains "$fj"[-1]-48)
								{
								$bgcolor="'#E8E8E8'"
								}
								else
								{
								$bgcolor="'#C8C8C8'"
								}
									$FullBody+="<tr align=""center"" bgcolor=$($bgcolor)>"
									$FullBody+="<td><font color=""#003333"">$($fullsender.user)</font></td>"
									if ($fullsender.deny -like "true"){$font="red"}else{$font="'#000000'"}
									$FullBody+="<td><font color=$font>$($fullsender.deny)</font></td>"
									if ($fullsender.isinherited -like "false"){$font="red"}else{$font="'#000000'"}
									$FullBody+="<td><font color=$font>$($fullsender.isinherited)</font></td>"
									$FullBody+="</tr>"
								$fj++
							}
						$FullBody+="</table></td>"
						$f++
						}
						
					if (!$sendbehalfs)
						{
						$BehalfBody+="<td align=""center"" valign=""top"" width=""33%""><table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#696969"" align=""center"" valign=""top""><font color=""#FFFF99"">No User has Send on Behalf On This Mailbox</font></td></table></td>"
						}
					else
						{
						$BehalfBody+="<td align=""center"" valign=""top"" width=""33%"">
						<table border=""0"" bordercolor=""#1F497B"" cellpadding=""3"" width=""100%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
						<tr><td align=""center valign=""top"" bgcolor=""#696969""><font color=""#FFFFE0"">Send on Behalf</font></td></tr>
						<tr><td bgcolor=""#878787"" nowrap=""nowrap""><font color=""#FFFFFF"">Send On Behalf Permission Owner</font></td></tr>"
						foreach ($sendbehalf in $sendbehalfs)
							{
								if (0,2,4,6,8 -contains "$bj"[-1]-48)
								{
								$bgcolor="'#E8E8E8'"
								}
								else
								{
								$bgcolor="'#C8C8C8'"
								}
									$BehalfBody+="<tr align=""center"" bgcolor=$($bgcolor)>"
									$BehalfBody+="<td><font color=""#003333"">$($sendbehalf.unescapedname)</font></td>"
									$BehalfBody+="</tr>"
								$bj++
							}
						$BehalfBody+="</table></td>"
						$b++
						}	
					$Table+=$MailboxTable+$SenderBody+$FullBody+$BehalfBody+"</tr></table><br><a href=""#top"">&#9650;</a><hr /><br>"
				}
		$u++
		}
_Progress (98) "Completing"
if (($ShowSelf -like "false") -and ($ShowInherited -like "false"))
{
	if(($dbnull -eq 0) -and ($mbnull -eq 0))
	{
	$Summary="<table style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#FFE87C"" >In your Exchange Organization there are $($mcount) mailboxes present."
	$Summary+="Send as Permission explicity configured on $($s) of these mailboxes. Full Access Permission explicity configured on $($f) of these mailboxes. Send on Behalf explicity configured on $($b) of these mailboxes and $($n) mailbox has inherited permissions only.<td></table><br>"
	}
	elseif (($dbnull -ne 0) -and ($mbnull -eq 0))
	{ 
	$Summary="<table style=""font-size:8pt;font-family:Arial,sans-serif""><td bgcolor=""#FFE87C"" >In $($database) mailbox database, there are $($mcount) mailboxes present."
	$Summary+="Send as Permission explicity configured on $($s) of these mailboxes. Full Access Permission explicity configured on $($f) of these mailboxes. Send on Behalf explicity configured on $($b) of these mailboxes and $($n) mailbox has inherited permissions only.<td></table><br>"
	}	
	$Header="
	<body>
	<font size=""1"" face=""Arial,sans-serif"">
	<h3 align=""center"">Mailbox Send As, Full Permission and Send on Behalf Report</h3>
	<h4 align=""center"">Generated $((Get-Date).ToString())</h4>"
	$inh+="</tr></table><br>"
	$gen+="</tr></table><br>"
	$Footer="</table></center><br><br>
	Scripted by <a href=""http://www.get-mailbox.org"">Serkan Varoglu</a>.  
	Elapsed Time To Complete This Report: $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString())</body></html>"
	if (($dbnull -eq 0) -and ($mbnull -eq 0))
	{
	$Output=$Header+$Summary+$gen+$inh+"<br><hr /><br>"+$Table+$Footer
	}
	elseif (($dbnull -ne 0) -and ($mbnull -eq 0))
	{ 
	$Output=$Header+$Summary+$gen+$inh+"<br><hr /><br>"+$Table+$Footer
	}
	else
	{
	if (($s -eq 0) -and ($f -eq 0) -and ($b -eq 0))
		{
		$Note="<center></font><b>Mailbox for $($Mailbox.name) ( $($Mailbox.primarysmtpaddress) ), does not have any explicit permissions set for Send As, Full Access or Send on Behalf</b></center>"
		}
	$Output=$Header+$Note+$Table+$Footer
	}
}
else
{
	$Header="
	<body>
	<font size=""1"" face=""Arial,sans-serif"">
	<h3 align=""center"">Mailbox Send As, Full Permission and Send on Behalf Report</h3>
	<a name=""top""><h4 align=""center"">Generated $((Get-Date).ToString())</h4></a>
	"
	$inh+="</tr></table><br>"
	$gen+="</tr></table><br>"
	$Footer="</table></center><br><br>
	<font size=""1"" face=""Arial,sans-serif"">Scripted by <a href=""http://www.get-mailbox.org"">Serkan Varoglu</a>.  
	Elapsed Time To Complete This Report: $($Watch.Elapsed.Minutes.ToString()):$($Watch.Elapsed.Seconds.ToString())</font></body></html>"
	$Output=$Header+$gen+$Table+$Footer

}
$Output | Out-File $HTMLReport