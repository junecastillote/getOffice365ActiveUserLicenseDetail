
<#PSScriptInfo

.VERSION 1.0

.GUID d45e01bb-a56f-440f-ac68-10a981668487

.AUTHOR June Castillote

.COMPANYNAME lazyexchangeadmin.com

.COPYRIGHT june.castillote@gmail.com

.TAGS Office365 GraphAPI OutlookMail OutlookREST Rest API

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 Get Active Users' Assigned Licenses using MS Graph API and Outlook Mail REST API 

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
Function WriteError
{
    param 
    (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Message
    )
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: $Message" -ForegroundColor RED
}

Function WriteInfo
{
    param 
    (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Message
    )
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": INFO: $Message" -ForegroundColor Yellow
}
#Function to get current system timezone (for PS versions below 5)
Function Get-TimeZoneInfo
{  
	$tzName = ([System.TimeZone]::CurrentTimeZone).StandardName
	$tzInfo = [System.TimeZoneInfo]::FindSystemTimeZoneById($tzName)
	Return $tzInfo	
}
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

#Function to get Script Version and ProjectURI for PS vesions below 5.1
Function Get-ScriptInfo
{
    param
    (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Path
	)
	
	$scriptFile = Get-Content $Path

	$props = @{
		Version = ""
		ProjectURI = ""
	}

	$scriptInfo = New-Object PSObject -Property $props

	# Get Version
	foreach ($line in $scriptFile)
	{	
		if ($line -like ".VERSION*")
		{
			$scriptInfo.Version = $line.Split(" ")[1]
			BREAK
		}	
	}

	# Get ProjectURI
	foreach ($line in $scriptFile)
	{
		if ($line -like ".PROJECTURI*")
		{
			$scriptInfo.ProjectURI = $line.Split(" ")[1]
			BREAK
		}		
	}
	Remove-Variable scriptFile
    Return $scriptInfo
}
#..........................................................
#EndRegion FUNCTIONS
Stop-TxnLogging
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$fileSuffix = (Get-Date -Format "dd-MMM-yyyy")

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
        WriteError "A valid sender email address is not specified."
        $isAllGood = $false
    }

    if (!$To)
    {
        WriteError "No recipient specified."
        $isAllGood = $false
    }
}

if ($isAllGood -eq $false)
{
    WriteError "Exiting Script."
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
WriteInfo "Acquire Graph Token"
try {
$body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$appID;client_secret=$appKey}
$oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token -Body $body
$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
}
catch {
    WriteError "Error getting Graph Token."
    WriteError $_.Exception.Message
    EXIT
}

$graphApiUri = "https://graph.microsoft.com/v1.0/organization"

$organizationName = (Invoke-RestMethod -Method Get -Uri $graphApiUri -Headers $headerParams -ErrorAction STOP).Value.DisplayName

$graphApiUri = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D30')"
try {
    $raw = (Invoke-RestMethod -Method Get -Uri $graphApiUri -Headers $headerParams -ErrorAction STOP).Remove(0,3)| ConvertFrom-Csv
}
catch {
    WriteError "Error retrieving data."
    WriteError $_.Exception.Message
    EXIT
}
$raw | Export-Csv $ReportCSV -NoTypeInformation
WriteInfo "Report saved to $ReportCSV."

#If sendEmail is not equal to TRUE, exit here.
if ($sendEmail -ne $true) {EXIT}

#Else, continue with email report

#convert attachment to base 64 encoded format
WriteInfo "Converting Report to Base 64 for use as email attachment."
$fileContentBytes = [System.Text.Encoding]::UTF8.GetBytes((get-content $ReportCSV -Raw))
$base64_csv = [System.Convert]::ToBase64String($fileContentBytes)

#Outlook API
WriteInfo "Acquire Outlook Token"

try {
    $outlookMailApiUri = "https://outlook.office.com/api/v2.0/users/$($From)/sendmail"
    $body = @{grant_type="client_credentials";scope="https://outlook.office.com/.default";client_id=$appID;client_secret=$appKey}
    $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token -Body $body -ErrorAction STOP
    $headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
}
catch {
    WriteError "Error getting Outlook Token."
    WriteError $_.Exception.Message
    EXIT
}


# Why use Outlook Mail REST API to send email instead of MS Graph API?
# Because MS Graph API cannot handle attachments larger than 4MB

WriteInfo "Creating Report."

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
table.steelBlueCols {
    font-family: "Arial Black", Gadget, sans-serif;
    border: 4px solid #555555;
    background-color: #555555;
    width: 400px;
    text-align: center;
    border-collapse: collapse;
  }
  table.steelBlueCols td, table.steelBlueCols th {
    border: 1px solid #555555;
    padding: 5px 10px;
  }
  table.steelBlueCols tbody td {
    font-size: 12px;
    font-weight: bold;
    color: #FFFFFF;
  }
  table.steelBlueCols tr:nth-child(even) {
    background: #398AA4;
  }
  table.steelBlueCols thead {
    background: #398AA4;
    border-bottom: 10px solid #398AA4;
  }
  table.steelBlueCols thead th {
    font-size: 15px;
    font-weight: bold;
    color: #FFFFFF;
    text-align: left;
    border-left: 2px solid #398AA4;
  }
  table.steelBlueCols thead th:first-child {
    border-left: none;
  }
  
  table.steelBlueCols tfoot td {
    font-size: 13px;
  }
  table.steelBlueCols tfoot .links {
    text-align: right;
  }
  table.steelBlueCols tfoot .links a{
    display: inline-block;
    background: #FFFFFF;
    color: #398AA4;
    padding: 2px 8px;
    border-radius: 5px;
  }
</style>
'@

$messageBody = '<html>'
$messageBody += "<head><title>$($messageSubject)</title>"
$messageBody += '<meta http-equiv="Content-Type content="text/html; charset=ISO-8859-1 />'
$messageBody += $cssString
$messageBody += '</head><body>'
$messageBody += '<table class="steelBlueCols">'
$messageBody += '<tr><th>Exchange</th><td>'+("{0:n0}" -f $license.Exchange)+'</td></tr>'
$messageBody += '<tr><th>SharePoint</th><td>'+("{0:n0}" -f $license.Sharepoint)+'</td></tr>'
$messageBody += '<tr><th>OneDrive</th><td>'+("{0:n0}" -f $license.OneDrive)+'</td></tr>'
$messageBody += '<tr><th>Skype for Business</th><td>'+("{0:n0}" -f $license.SkypeForBusiness)+'</td></tr>'
$messageBody += '<tr><th>Teams</th><td>'+("{0:n0}" -f $license.Teams)+'</td></tr>'
$messageBody += '<tr><th>Yammer</th><td>'+("{0:n0}" -f $license.Yammer)+'</td></tr>'
$messageBody += '</table>'
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
WriteInfo "Send Email Report"
try {
    Invoke-RestMethod -Method Post -Uri $outlookMailApiUri -Body $mailbody -Headers $headerParams -ContentType application/json -ErrorAction STOP
}
catch {
    WriteError "Error sending email."
    WriteError $_.Exception.Message
    EXIT
}