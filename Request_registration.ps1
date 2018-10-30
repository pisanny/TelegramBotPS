<# 
 
.DESCRIPTION 
    Request registration in the Telegram channel.
 
.NOTES 
    Author: Vladimir Pisanny
    Last Updated: 10/26/2018   
#> 

$CredManPath = ".\CredMan.ps1"
$Target = 'http://server'
$CredManScript = [Scriptblock]::Create((Get-Content $CredManPath -Raw) + "Read-Creds -Target $Target")
$credman = &$CredManScript
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $credman.UserName, (ConvertTo-SecureString -String $credman.CredentialBlob -AsPlainText -Force)

$proxy ="http://"
$smtp = "smtp server"
$sysemail = "AlertBot@domain.by"

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
