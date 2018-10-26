<# 
 
.DESCRIPTION 
    Request registration in the Telegram channel.
 
.NOTES 
    Author: Vladimir Pisanny
    Last Updated: 10/26/2018   
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

$proxy ="http://"
$smtp = "smtp server"
$sysemail = "AlertBot@domain.by"
$credman = Read-Creds -Target 'http://server'
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $credman.UserName, (ConvertTo-SecureString -String $credman.CredentialBlob -AsPlainText -Force)
$SqlServer = "SQL server"
$SqlDB = "DB"

$token = "XXXX"
$URLG = "https://api.telegram.org/bot$token/getUpdates?offset=$UpdateId&timeout=$ChatTimeout"

$Request = Invoke-WebRequest -Uri $URLG -Method Get -Proxy $proxy
$content = ConvertFrom-Json $Request.content

$message = @()
foreach ($str in $content.result) {

             $props = @{
                        ok = $content.ok
                        UpdateId = $str.update_id
                        Message_ID = $str.message.message_id
                        first_name = $str.message.from.first_name
                        last_name = $str.message.from.last_name
                        chat_id = $str.message.chat.id
                        text = $str.message.text
                       }
    $message += $props
}

$users = @()
Import-Module ActiveDirectory
$message | %{If($_.text -match "reg") {
    $re = "[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
    $email = [regex]::MAtch($_.text, $re, "IgnoreCase ").Value
    $aduser = try{Get-ADUser -Filter {mail -eq $email} -Properties mail,mobile} Catch {$null}
        
    If( $aduser) {
                  $user = @{
                            chat_id = $_.chat_id
                            SID = $aduser.SID
                            Enabled = $aduser.Enabled
                            Name = $aduser.Name
                            mail = $aduser.mail
                            mobile = $aduser.mobile
                           }       
                  $users += $user
                 }
   }
}

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server=$SqlServer; Database=$SqlDB; Integrated Security=True"
$SqlConnection.Open()

foreach ($user in $users) { 
            $SID = "'" + $user.SID.Value + "'"
            $SqlCmd = $SqlConnection.CreateCommand()
            $SqlCmd.CommandText = "SELECT CAST(CASE WHEN EXISTS(SELECT * FROM chatusers where [SID] = $SID) THEN 1 ELSE 0 END AS BIT)"
            $Reader = $SqlCmd.ExecuteReader()
            $table = new-object "System.Data.DataTable"
            $table.Load($Reader)
            $Reader.close()
            [bool]$queryResult = $($table.Rows[0])[0]
            if (!($queryResult)) {
                    $SqlStr = "INSERT INTO chatusers (chat_id,SID, Registered, DisplayName) VALUES ('" + $user.chat_id + "','" + $user.SID + "','0','" + $user.Name + "')"
                    $SqlCmd = $SqlConnection.CreateCommand()
                    $SqlCmd.CommandText = $SqlStr
                    $SqlCmd.ExecuteNonQuery() | Out-Null
                    $Body = 'Hi ' + $user.Name + ', yor Chat ID: ' + $user.chat_id + '. Reply to the email to complete the registration please.'
                    Send-MailMessage -Credential $Credential -Port 587 -To $user.mail -From $sysemail -SmtpServer $smtp -Subject "Chat ID" -Body $Body
                }
           }


$SqlConnection.close()
