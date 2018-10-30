<# 
 
.DESCRIPTION 
    Registration completed in the Telegram channel.
 
.NOTES 
    Author: Vladimir Pisanny
    Last Updated: 10/29/2018   
#> 


# The functions Get-CredType and Read-Creds are taken from the script CredMan.ps1, ator Jim Harrison (jim@isatools.org).

$CredManPath = ".\CredMan.ps1"
$Target = 'http://server'
$CredManScript = [Scriptblock]::Create((Get-Content $CredManPath -Raw) + "Read-Creds -Target $Target")
$credman = &$CredManScript
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $credman.UserName, (ConvertTo-SecureString -String $credman.CredentialBlob -AsPlainText -Force)

$SqlServer = "SQL server"
$SqlDB = "db"

$domain = "domain"
$sysemail = "AlertBot@domain.by"
$smtp = "mail server"


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