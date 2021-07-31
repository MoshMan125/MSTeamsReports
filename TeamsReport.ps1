<#	
	.NOTES
	===========================================================================
	 Created on:   	2/24/2021 11:14 AM
	 Created by:   	MoshMan125
	===========================================================================
	.DESCRIPTION
		Report all team channels, members and roles.
#>
#variables paths and such...
$CSV = "$env:USERPROFILE\Documents\Teams_UserAudit.csv"
$date = Get-Date -Format 'MM.dd-HH.mm.ss'
$newName = 'Teams_UserAudit-' + $date + '.csv'
$fromEmail = '<EmailAddress>'
$emailPassword = '<Password>'

# email function
function sendMail
{
	[CmdletBinding()]
	param
	(
		[parameter(Mandatory = $true)]
		$messageBody,
		[Parameter(Mandatory = $true)]
		$messageSubject,
		[Parameter(Mandatory = $true)]
		$sendTo,
		[Parameter(Mandatory = $false)]
		$attachment
	)
	'Sending Email'
	
	#SMTP server name
	$smtpServer = 'smtp.office365.com'
	
	#Creating a Mail object
	$msg = new-object Net.Mail.MailMessage
	
	$emailCredential = New-Object System.Net.NetworkCredential($fromEmail, $emailPassword)
	
	#Creating SMTP server object
	$smtp = new-object Net.Mail.SmtpClient($smtpServer)
	$smtp.Port = 587
	
	$smtp.EnableSSl = $true
	$smtp.Credentials = $emailCredential
	
	#Email structure
	$msg.From = $fromEmail
	$msg.To.add($toEmail)
	$msg.subject = $messageSubject
	$msg.body = $messageBody
	$msg.Attachments = $attachment
	
	#Sending email
	$smtp.Send($msg)
	
	"Email Sent"
}

# connect to MicrosoftTeams function
function MSConnnect
{
	$installCheck = Get-Module -Name MicrosoftTeams
	if ($installCheck -eq $null)
	{
		try
		{
			Install-Module -Name MicrosoftTeams -Force -ErrorAction Stop
		}
		catch
		{
			(New-Object -COM WScript.Shell).PopUp("Failed to install MicrosoftTeams module, please install manually and try again.", 0, "Error", 48)
		}
	}	
}

#check if output file already exists
$reportCheck = Test-Path $CSV

if ($reportCheck -eq 'True')
{
	Try
	{
		Rename-Item -Path $CSV -NewName $newName -ErrorAction Stop
	}
	catch
	{
		(New-Object -COM WScript.Shell).PopUp("Failed to generate report, please close $CSV and try again.", 0, "Error", 48)
	}
}

#connect to teams powershell module
MSConnnect

#loop through each team and channel to output team name, channel name, team guid, users, and roles
foreach ($Team in ($Teams = Get-Team))
{
	$Channels = Get-TeamChannel -GroupId $Team.GroupID;
	foreach ($Channel in $Channels)
	{
		$Users = Get-TeamChannelUser -GroupId $Team.GroupID -DisplayName $Channel.DisplayName;
		foreach ($User in $Users)
		{
			[PSCustomObject]@{
				'ChannelDisplayName' = $Channel.DisplayName
				'TeamDisplayName'    = $Team.DisplayName
				'TeamGUID'		     = $Team.GroupID
				'Users'			     = [string]$User.User
				'Role'			     = [string]$User.Role
			} | Export-Csv -Path $CSV -NoTypeInformation -Append
		}
	}
}

#display pop-up box showing completion and file path
$reportCheck = Test-Path $CSV
if ($reportCheck -eq 'True')
{
	$popup = (New-Object -COM WScript.Shell).PopUp("Report completed, do you want to open the file?`n `nPath:$CSV", 0, "TeamsReport Complete", 0x4)
	if ($popup -eq 6)
	{
		Invoke-Item -Path $CSV
	}
}
else
{
	(New-Object -COM WScript.Shell).PopUp("ERROR: Please try again, if the error persists please contact IT", 0, "Error", 48)
}

#send as email?
$emailPopup = (New-Object -COM WScript.Shell).PopUp("Report completed, do you want to send the file via email?", 0, "TeamsReport Complete", 0x4)
if ($emailPopup -eq 6)
{
	[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
	$toEmail = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the users email address to send the report', 'UserInfo')
	sendMail -messageBody 'Find report attached' -messageSubject 'Teams Audit Report' -sendTo $toEmail -attachment $CSV
}
else
{
	exit
}