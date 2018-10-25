<# 
 
.DESCRIPTION 
    Get chat ID to e-mail
 
.NOTES 
    Author: Vladimir Pisanny
    Last Updated: 10/25/2018   
#> 

$proxy ="http://"
$smtp = "srv"
$sysemail = "email@domen.by"
$Credential = Get-Credential


$token = "xxxxx"
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

Import-Module ActiveDirectory
$message | %{If($_.text -match "reg") {
    $re = "[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
    $email = [regex]::MAtch($_.text, $re, "IgnoreCase ").Value
    $user = try{Get-ADUser -Filter {mail -eq $email} -Properties mail} Catch {$null}

    If( !([string]::IsNullOrWhiteSpace($user))) {
         $Body = 'Hi ' + $_.first_name + ', yor Chat ID: ' + $_.chat_id + '.'
         Send-MailMessage -Credential $Credential -To $user.mail -From $sysemail -SmtpServer $smtp -Subject "Chat ID" -Body $Body
        }
   }
}
