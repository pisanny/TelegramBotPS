<# 
 
.DESCRIPTION 
    Registration completed in the Telegram channel.
 
.NOTES 
    Author: Vladimir Pisanny
    Last Updated: 10/29/2018   
#> 


# The functions Get-CredType and Read-Creds are taken from the script CredMan.ps1, ator Jim Harrison (jim@isatools.org).

function Get-CredType
{
	Param
	(
		[Parameter(Mandatory=$true)][ValidateSet("GENERIC",
												  "DOMAIN_PASSWORD",
												  "DOMAIN_CERTIFICATE",
												  "DOMAIN_VISIBLE_PASSWORD",
												  "GENERIC_CERTIFICATE",
												  "DOMAIN_EXTENDED",
												  "MAXIMUM",
												  "MAXIMUM_EX")][String] $CredType
	)
	
	switch($CredType)
	{
		"GENERIC" {return [PsUtils.CredMan+CRED_TYPE]::GENERIC}
		"DOMAIN_PASSWORD" {return [PsUtils.CredMan+CRED_TYPE]::DOMAIN_PASSWORD}
		"DOMAIN_CERTIFICATE" {return [PsUtils.CredMan+CRED_TYPE]::DOMAIN_CERTIFICATE}
		"DOMAIN_VISIBLE_PASSWORD" {return [PsUtils.CredMan+CRED_TYPE]::DOMAIN_VISIBLE_PASSWORD}
		"GENERIC_CERTIFICATE" {return [PsUtils.CredMan+CRED_TYPE]::GENERIC_CERTIFICATE}
		"DOMAIN_EXTENDED" {return [PsUtils.CredMan+CRED_TYPE]::DOMAIN_EXTENDED}
		"MAXIMUM" {return [PsUtils.CredMan+CRED_TYPE]::MAXIMUM}
		"MAXIMUM_EX" {return [PsUtils.CredMan+CRED_TYPE]::MAXIMUM_EX}
	}
}

function Read-Creds
{
<#
.Synopsis
  Reads specified credentials for operating user

.Description
  Calls Win32 CredReadW via [PsUtils.CredMan]::CredRead

.INPUTS

.OUTPUTS
  [PsUtils.CredMan+Credential] if successful
  [Management.Automation.ErrorRecord] if unsuccessful or error encountered

.PARAMETER Target
  Specifies the URI for which the credentials are associated
  If not provided, the username is used as the target
  
.PARAMETER CredType
  Specifies the desired credentials type; defaults to 
  "CRED_TYPE_GENERIC"
#>

	Param
	(
		[Parameter(Mandatory=$true)][ValidateLength(1,32767)][String] $Target,
		[Parameter(Mandatory=$false)][ValidateSet("GENERIC",
												  "DOMAIN_PASSWORD",
												  "DOMAIN_CERTIFICATE",
												  "DOMAIN_VISIBLE_PASSWORD",
												  "GENERIC_CERTIFICATE",
												  "DOMAIN_EXTENDED",
												  "MAXIMUM",
												  "MAXIMUM_EX")][String] $CredType = "GENERIC"
	)
	
	if("GENERIC" -ne $CredType -and 337 -lt $Target.Length) #CRED_MAX_DOMAIN_TARGET_NAME_LENGTH
	{
		[String] $Msg = "Target field is longer ($($Target.Length)) than allowed (max 337 characters)"
		[Management.ManagementException] $MgmtException = New-Object Management.ManagementException($Msg)
		[Management.Automation.ErrorRecord] $ErrRcd = New-Object Management.Automation.ErrorRecord($MgmtException, 666, 'LimitsExceeded', $null)
		return $ErrRcd
	}
	[PsUtils.CredMan+Credential] $Cred = New-Object PsUtils.CredMan+Credential
    [Int] $Results = 0
	try
	{
		$Results = [PsUtils.CredMan]::CredRead($Target, $(Get-CredType $CredType), [Ref]$Cred)
	}
	catch
	{
		return $_
	}
	
	switch($Results)
	{
        0 {break}
        0x80070490 {return $null} #ERROR_NOT_FOUND
        default
        {
    		[String] $Msg = "Error reading credentials for target '$Target' from '$Env:UserName' credentials store"
    		[Management.ManagementException] $MgmtException = New-Object Management.ManagementException($Msg)
    		[Management.Automation.ErrorRecord] $ErrRcd = New-Object Management.Automation.ErrorRecord($MgmtException, $Results.ToString("X"), $ErrorCategory[$Results], $null)
    		return $ErrRcd
        }
	}
	return $Cred
}


$SqlServer = "SQL server"
$SqlDB = "db"

$domain = "domain"
$sysemail = "AlertBot@domain.by"
$smtp = "mail server"
$credman = Read-Creds -Target 'http://server'
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $credman.UserName, (ConvertTo-SecureString -String $credman.CredentialBlob -AsPlainText -Force)

# Microsoft Exchange Web Services Managed API 2.2

Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
Import-Module ActiveDirectory

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server=$SqlServer; Database=$SqlDB; Integrated Security=True"
$SqlConnection.Open()

$Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($credman.UserName,$credman.CredentialBlob,$domain)
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.Credentials = $Credentials
$exchService.AutodiscoverUrl($sysemail)
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchservice,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$psPropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

$Items = $Inbox.FindItems(100) | ? Subject -Match "Chat ID"

foreach ($Item in $Items)
    {
     $SID = "'" + (Get-ADUser -Filter {mail -eq $Item.Sender.Address}).SID + "'"
     $SqlCmd = $SqlConnection.CreateCommand()
     $SqlCmd.CommandText = "SELECT CAST(CASE WHEN EXISTS(SELECT * FROM chatusers where [SID] = $SID AND [Registered] = 0) THEN 1 ELSE 0 END AS BIT)"
     $Reader = $SqlCmd.ExecuteReader()
     $table = new-object "System.Data.DataTable"
     $table.Load($Reader)
     $Reader.close()
     [bool]$queryResult = $($table.Rows[0])[0]
     if ($queryResult) {
           $SqlStr = "UPDATE chatusers SET [Registered] = '1' WHERE [SID] = $SID"
           $SqlCmd = $SqlConnection.CreateCommand()
           $SqlCmd.CommandText = $SqlStr
           $SqlCmd.ExecuteNonQuery() | Out-Null
           $Body = 'Hi ' + $Item.Sender.Name + '. Registration successfully completed.'
           Send-MailMessage -Credential $Credential -Port 587 -To $Item.Sender.Address -From $sysemail -SmtpServer $smtp -Subject "Chat ID" -Body $Body
      }
      $Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete, $true)
    }

$SqlConnection.close()