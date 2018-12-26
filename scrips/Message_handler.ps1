<# 
 
.DESCRIPTION 
    Handles messages in the system mailbox.
 
.NOTES 
    Author: Vladimir Pisanny
    Last Updated: 11/01/2018   
#> 

$CredManPath = "$ENV:USERPROFILE\Documents\ps\CredMan.ps1"
$Target = 'http://server'

$token = "xxxxxx"
$proxy ="http://"

$SqlServer = "SQL server"
$SqlDB = "db"

$domain = "domain"
$sysemail = "AlertBot@domain.by"
$monmail = "MonitoringSystems@domain.by"
$smtp = "mail server"

$CredManScript = [Scriptblock]::Create((Get-Content $CredManPath -Raw) + "Read-Creds -Target $Target")
$credman = &$CredManScript
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $credman.UserName, (ConvertTo-SecureString -String $credman.CredentialBlob -AsPlainText -Force)

Import-Module ActiveDirectory
Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server=$SqlServer; Database=$SqlDB; Integrated Security=True"
$SqlConnection.Open()
$SqlCmd = $SqlConnection.CreateCommand()

$Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($credman.UserName,$credman.CredentialBlob,$domain)
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.Credentials = $Credentials
$exchService.AutodiscoverUrl($sysemail)
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchservice,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$psPropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

$Items = $Inbox.FindItems(100)

foreach ($Item in $Items)
    {
     If ($Item.Subject -Match "Chat ID")
        {
         $SID = "'" + (Get-ADUser -Filter {mail -eq $Item.Sender.Address}).SID + "'"
         $SqlCmd.CommandText = "SELECT CAST(CASE WHEN EXISTS(SELECT * FROM chatusers where [SID] = $SID AND [Registered] = 0) THEN 1 ELSE 0 END AS BIT)"
         $Reader = $SqlCmd.ExecuteReader()
         $table = new-object "System.Data.DataTable"
         $table.Load($Reader)
         $Reader.close()
         [bool]$queryResult = $($table.Rows[0])[0]
         If ($queryResult) {
              $SqlStr = "UPDATE chatusers SET [Registered] = '1' WHERE [SID] = $SID"
              $SqlCmd.CommandText = $SqlStr
              $SqlCmd.ExecuteNonQuery() | Out-Null
              $Body = 'Hi ' + $Item.Sender.Name + '. Registration successfully completed.'
              Send-MailMessage -Credential $Credential -Port 587 -To $Item.Sender.Address -From $sysemail -SmtpServer $smtp -Subject "Chat ID" -Body $Body
            }
        }
     If ($Item.Sender.Address -eq $monmail)
        {
         $Item.load($psPropertySet)
         $payload = @{ "parse_mode" = "Markdown"; "disable_web_page_preview" = "True" }
         $textmessage =  $Item.Body
         Get-ADGroupMember $Item.DisplayTo.Trim() | ? name -ne $credman.UserName | % {
             $SqlCmd.CommandText = "SELECT chat_id FROM chatusers where SID = '" + $_.SID + "' AND Registered = '1'"
             $Reader = $SqlCmd.ExecuteReader()
             $table = new-object "System.Data.DataTable"
             $table.Load($Reader)
             $Reader.close()
             If ($table -ne $null)
                {
                 $recipient = $($table.Rows[0])[0]
                 $URLS = "https://api.telegram.org/bot$token/sendMessage?chat_id=$recipient&text=$textmessage"
                 $Request = Invoke-WebRequest -Uri $URLS -Method Post -ContentType "application/json; charset=utf-8" -Body (ConvertTo-Json -Compress -InputObject $payload) -Proxy $proxy
                }
            }
        }
     $Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete, $true)
    }

$SqlConnection.close()