
<#PSScriptInfo

.VERSION 1.0

.GUID d45e01bb-a56f-440f-ac68-10a981668487

.AUTHOR June Castillote

.COMPANYNAME lazyexchangeadmin.com

.COPYRIGHT june.castillote@gmail.com

.TAGS Office365 GraphAPI OutlookMail OutlookREST Rest API

.LICENSEURI https://raw.githubusercontent.com/junecastillote/getOffice365ActiveUserLicenseDetail/master/LICENSE

.PROJECTURI https://github.com/junecastillote/getOffice365ActiveUserLicenseDetail

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<#

.DESCRIPTION
 Get Active Users' Assigned Licenses using MS Graph API and Outlook Mail REST API, capable of sending the resulting CSV by email.

#>

<#
.SYNOPSIS
    Get Active Users' Assigned Licenses using MS Graph API and Outlook Mail REST API, capable of sending the resulting CSV by email.
.EXAMPLE
    PS C:\> .\getOffice365ActiveUserLicenseDetail.ps1 -appID <appID> -appKey <appKey> -tenantID <tenantID>
    This example will query the License assignment data from Office 365 and save the resuting CSV file in the ".\Report" folder.
.EXAMPLE
    PS C:\> .\getOffice365ActiveUserLicenseDetail.ps1 -appID <appID> -appKey <appKey> -tenantID <tenantID> -ReportPath C:\Temp
    This example will query the License assignment data from Office 365 and save the resuting CSV file in the "C:\Temp" folder.
.EXAMPLE
    PS C:\> .\getOffice365ActiveUserLicenseDetail.ps1 -appID <appID> -appKey <appKey> -tenantID <tenantID> -ReportPath C:\Temp -sendEmail $true -From sender@domain.com -To recipient@domain.com
    This example will query the License assignment data from Office 365 and save the resuting CSV file in the "C:\Temp" folder. Then send the summary report by email with the CSV file attached.
#>
Param(
    [CmdletBinding()]

    [Parameter(Mandatory=$true,Position=0)]
    [string]
    $appID,

    [Parameter(Mandatory=$true,Position=1)]
    [string]
    $appKey,

    [Parameter(Mandatory=$true,Position=2)]
    [string]
    $tenantID,

    [Parameter()]
    [boolean]
    $sendEmail=$false,

    [Parameter()]
    [string]
    $From,

    [Parameter()]
    [string[]]
    $To,

    [Parameter()]
    [string]
    $ReportPath,

    [Parameter()]
    [string]
    $LogPath
)

#Region FUNCTIONS
#..........................................................
#Function to stop transcribing
Function Stop-TxnLogging
{
	$txnLog=""
	Do {
		try {
			Stop-Transcript | Out-Null
		}
		catch [System.InvalidOperationException]{
			$txnLog="stopped"
		}
    } While ($txnLog -ne "stopped")
}

#Function to Start transcribing
Function Start-TxnLogging
{
    param
    (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$LogFile
    )
	Stop-TxnLogging
    Start-Transcript $LogFile -Append
}
#..........................................................
#EndRegion FUNCTIONS

Stop-TxnLogging

# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$fileSuffix = (Get-Date -Format "dd-MMM-yyyy")
$scriptInfo = Test-ScriptFileInfo -Path $MyInvocation.MyCommand.Definition

#Region Paths
#..........................................................
#If ReportPath if blank, create the default path
if (!$ReportPath)
{
    $ReportPath = "$script_root\Report"
}

#If ReportPath does not exist, create it
if (!(Test-Path $ReportPath))
{
     New-Item -ItemType Directory -Path $ReportPath | Out-Null
}

#Define report CSV file
$ReportFileNameCSV = "ActiveUserLicenseDetail_$fileSuffix.csv"
$ReportFileNameHTML = "ActiveUserLicenseCount_$fileSuffix.HTML"
$ReportCSV = "$ReportPath\$ReportFileNameCSV"
$ReportHTML = "$ReportPath\$ReportFileNameHTML"

#If LogPath if blank, create the default path
if (!$LogPath)
{
    $LogPath = "$script_root\Log"
}

#If ReportPath does not exist, create it
if (!(Test-Path $LogPath))
{
     New-Item -ItemType Directory -Path $LogPath | Out-Null
}

#Define Log file
$LogFile = "$LogPath\Log_$fileSuffix.log"

#Start Transcript
Start-TxnLogging $LogFile
#..........................................................
#EndRegion Paths

#Region MailParamCheck
#..........................................................
$isAllGood = $true

if ($sendEmail -eq $true)
{
    if (!$From)
    {
        Write-Output "A valid sender email address is not specified."
        $isAllGood = $false
    }

    if (!$To)
    {
        Write-Output "No recipient specified."
        $isAllGood = $false
    }
}

if ($isAllGood -eq $false)
{
    Write-Output "Exiting Script."
    EXIT
}
#..........................................................
#EndRegion MailParamCheck

if ($To)
{
    $toAddressJSON = @()
    $To | ForEach-Object {$toAddressJSON += @{EmailAddress = @{Address = $_}}}
}

#Graph API Token
Write-Output "Acquire Graph Token."
try {
$body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$appID;client_secret=$appKey}
$oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token -Body $body
$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
}
catch {
    Write-Output "Error getting Graph Token."
    Write-Output $_.Exception.Message
    EXIT
}

$activeUserDetailURI = "https://graph.microsoft.com/v1.0/organization"

$organizationName = (Invoke-RestMethod -Method Get -Uri $activeUserDetailURI -Headers $headerParams -ErrorAction STOP).Value.DisplayName

$activeUserDetailURI = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D30')"
try {
    $raw = (Invoke-RestMethod -Method Get -Uri $activeUserDetailURI -Headers $headerParams -ErrorAction STOP).Remove(0,3)| ConvertFrom-Csv
}
catch {
    Write-Output "Error retrieving data."
    Write-Output $_.Exception.Message
    EXIT
}
$raw | Export-Csv $ReportCSV -NoTypeInformation
Write-Output "Report saved to $ReportCSV."

#If sendEmail is not equal to TRUE, exit here.
if ($sendEmail -ne $true) {EXIT}

#Else, continue with email report

#convert attachment to base 64 encoded format
Write-Output "Converting Report to Base 64 for use as email attachment."
$fileContentBytes = [System.Text.Encoding]::UTF8.GetBytes((get-content $ReportCSV -Raw))
$base64_csv = [System.Convert]::ToBase64String($fileContentBytes)

#Outlook API
Write-Output "Acquire Outlook Token."

try {
    $outlookMailApiUri = "https://outlook.office.com/api/v2.0/users/$($From)/sendmail"
    $body = @{grant_type="client_credentials";scope="https://outlook.office.com/.default";client_id=$appID;client_secret=$appKey}
    $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token -Body $body -ErrorAction STOP
    $headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
}
catch {
    Write-Output "Error getting Outlook Token."
    Write-Output $_.Exception.Message
    EXIT
}


# Why use Outlook Mail REST API to send email instead of MS Graph API?
# Because MS Graph API cannot handle attachments larger than 4MB

Write-Output "Creating Report."

$licenseProps = @{
    Exchange = ($raw | Where-Object {$_."Has Exchange License" -eq $true}).count
    Sharepoint = ($raw | Where-Object {$_."Has Sharepoint License" -eq $true}).count
    OneDrive = ($raw | Where-Object {$_."Has OneDrive License" -eq $true}).count
    SkypeForBusiness = ($raw | Where-Object {$_."Has Skype For Business License" -eq $true}).count
    Teams = ($raw | Where-Object {$_."Has Teams License" -eq $true}).count
    Yammer = ($raw | Where-Object {$_."Has Yammer License" -eq $true}).count
}

$license = New-Object psobject -Property $licenseProps

#message
$messageSubject = "[$($organizationName)] Assigned Licenses Report : " + (Get-Date -format F)

$cssString = @'
<style type="text/css">
.tftable {table-layout:fixed;width: 40%;font-family:"Segoe UI";font-size:12px;color:#333333;border-width: 1px;border-color: #729ea5;border-collapse: collapse;}
.tftable th {width: 30%;font-size:12px;background-color:#acc8cc;border-width: 1px;padding: 8px;border-style: solid;border-color: #729ea5;text-align:left;}
.tftable tr {background-color:#d4e3e5;}
.tftable td {width: 10%font-size:12px;border-width: 1px;padding: 8px;border-style: solid;border-color: #729ea5;}
.tftable tr:hover {background-color:#ffffff;}
</style>
'@

$messageBody = '<html>'
$messageBody += "<head><title>$($messageSubject)</title>"
$messageBody += '<meta http-equiv="Content-Type content="text/html; charset=ISO-8859-1 />'
$messageBody += $cssString
$messageBody += '</head><body>'
$messageBody += '<p><font face="Segoe UI"><h3>Summary of Assigned Licenses Count</h3></font></p>'
$messageBody += '<table class="tftable">'
$messageBody += '<tr><th>Exchange</th><td>'+("{0:n0}" -f $license.Exchange)+'</td></tr>'
$messageBody += '<tr><th>SharePoint</th><td>'+("{0:n0}" -f $license.Sharepoint)+'</td></tr>'
$messageBody += '<tr><th>OneDrive</th><td>'+("{0:n0}" -f $license.OneDrive)+'</td></tr>'
$messageBody += '<tr><th>Skype for Business</th><td>'+("{0:n0}" -f $license.SkypeForBusiness)+'</td></tr>'
$messageBody += '<tr><th>Teams</th><td>'+("{0:n0}" -f $license.Teams)+'</td></tr>'
$messageBody += '<tr><th>Yammer</th><td>'+("{0:n0}" -f $license.Yammer)+'</td></tr>'
$messageBody += '</table><hr />'
$messageBody += '<p><font face="Segoe UI"><h3>End of Report<h3></font></p>'
$messageBody += '<p><font size="2" face="Segoe UI">'
$messageBody += 'Source: ' + ($env:COMPUTERNAME) + '<br />'
$messageBody += 'Script Path: ' + ($MyInvocation.MyCommand.Definition) + '<br />'
$messageBody += 'Script Version: <a href="' + ($scriptInfo.ProjectURI) + '">'+ ($MyInvocation.MyCommand.Definition.ToString().Split("\")[-1].Split(".")[0]) + ' ' + ($scriptInfo.version) + '</a><br />'

$messageBody += '</body>'
$messageBody += '</html>'
$messageBody | out-file $ReportHTML


#Careful with this, the Outlook REST API body is CASE-sensitive when in comes to the parameters (eg. 'Subject' is not the same as 'subject')
$mailBody = @{
	Message = @{
		Subject = $messageSubject
		Body = @{
			ContentType = "HTML"
            Content = $messageBody
            #Content = "ATTACHED"
		}
		ToRecipients = @(
			$ToAddressJSON
		)
		Attachments = @(
			@{
				"@odata.type" = "#Microsoft.OutlookServices.FileAttachment"
				Name = "$($ReportFileNameCSV)"
				ContentType = "multipart/mixed"
				ContentBytes = $base64_csv
			}
		)
	}
	SaveToSentItems = $false
}
$mailBody = $mailBody | ConvertTo-JSON -Depth 4

#send email
Write-Output "Send Email Report."
try {
    Invoke-RestMethod -Method Post -Uri $outlookMailApiUri -Body $mailbody -Headers $headerParams -ContentType application/json -ErrorAction STOP
}
catch {
    Write-Output "Error sending email."
    Write-Output $_.Exception.Message
    EXIT
}