<# 
 
.DESCRIPTION 
    Implementation of the "reg" command.
 
.NOTES 
    Author: Vladimir Pisanny
    Last Updated: 11/01/2018   
#> 

Import-Module ActiveDirectory

$re = "[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
$email = [regex]::MAtch($message.text, $re, "IgnoreCase ").Value
$aduser = try{Get-ADUser -Filter {mail -eq $email} -Properties mail,mobile} Catch {$null}

If( $aduser)
    {
     $SqlCmd = $SqlConnection.CreateCommand()
     $SqlCmd.CommandText = "SELECT CAST(CASE WHEN EXISTS(SELECT * FROM chatusers where [SID] = '" + $aduser.SID.Value + "') THEN 1 ELSE 0 END AS BIT)"
     $Reader = $SqlCmd.ExecuteReader()
     $table = new-object "System.Data.DataTable"
     $table.Load($Reader)
     $Reader.close()
     [bool]$queryResult = $($table.Rows[0])[0]
     If (!($queryResult))
        {
         $SqlStr = "INSERT INTO chatusers (chat_id,SID, Registered, DisplayName) VALUES ('" + $message.chat_id + "','" + $aduser.SID.Value + "','0','" + $aduser.Name + "')"
         $SqlCmd = $SqlConnection.CreateCommand()
         $SqlCmd.CommandText = $SqlStr
         $SqlCmd.ExecuteNonQuery() | Out-Null
         $Body = 'Hi ' + $aduser.Name + ', yor Chat ID: ' + $message.chat_id + '. Reply to the email to complete the registration please.'
         Send-MailMessage -Credential $Credential -Port 587 -To $aduser.mail -From $sysemail -SmtpServer $smtp -Subject "Chat ID" -Body $Body
         $script:result = $Body
        }
        Else
            {
             $script:result = @{
                                UpdateId = $message.UpdateId
                                text = 'Registration already requested.'
                                chat_id = $message.chat_id
                                Message_ID = $message.Message_ID
                                first_name = $message.first_name
                                last_name  = $message.last_name
                               }
            }
    }
Else
    {
     $script:result = @{
                        UpdateId = $message.UpdateId
                        text = 'Invalid email.'
                        chat_id = $message.chat_id
                        Message_ID = $message.Message_ID
                        first_name = $message.first_name
                        last_name  = $message.last_name
                       }
    }