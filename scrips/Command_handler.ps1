<# 
 
.DESCRIPTION 
    Command handler.
 
.NOTES 
    Author: Vladimir Pisanny
    Last Updated: 10/26/2018   
#> 

$CredManPath = ".\CredMan.ps1"
$Target = 'http://server'

$token = "xxxx"
$proxy ="http://"

$SqlServer = "SQL server"
$SqlDB = "db"

$domain = "domain"
$smtp = "mail server"
$sysemail = "AlertBot@domain.by"

$CredManScript = [Scriptblock]::Create((Get-Content $CredManPath -Raw) + "Read-Creds -Target $Target")
$credman = &$CredManScript
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $credman.UserName, (ConvertTo-SecureString -String $credman.CredentialBlob -AsPlainText -Force)

$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server=$SqlServer; Database=$SqlDB; Integrated Security=True"
$SqlConnection.Open()
$SqlCmd = $SqlConnection.CreateCommand()

$SqlCmd.CommandText = "SELECT UpdateId FROM UpdateId"
$Reader = $SqlCmd.ExecuteReader()
$table = new-object "System.Data.DataTable"
$table.Load($Reader)
$Reader.close()
$UpdateId = ($table.Rows.UpdateId + 1).ToString()

$URLG = "https://api.telegram.org/bot$token/getUpdates?offset=$UpdateId&timeout=$ChatTimeout"
$Request = Invoke-WebRequest -Uri $URLG -Method Get -Proxy $proxy
$content = ConvertFrom-Json $Request.content

$messages = @()
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
    $messages += $props
}

$SqlCmd.CommandText = "SELECT command FROM scripts"
$Reader = $SqlCmd.ExecuteReader()
$table = new-object "System.Data.DataTable"
$table.Load($Reader)
$Reader.close()
$commands = $table.Rows.command

$results = @()
foreach ($message in $messages)
    {
     foreach ($command in $commands)
        {
         If($message.text -match $command)
            {
             $SqlCmd.CommandText = "SELECT scriptblock FROM scripts WHERE command = '" + $command + "'"
             $Reader = $SqlCmd.ExecuteReader()
             $table = new-object "System.Data.DataTable"
             $table.Load($Reader)
             $Reader.close()
             $scriptBlock = [Scriptblock]::Create($($table.Rows[0])[0])
             &$scriptBlock
             $results += $result
            }
        }
    }

If ($messages -ne $null)
    {
     $SqlCmd.CommandText = "DELETE FROM UpdateId"
     $SqlCmd.ExecuteNonQuery() | Out-Null
     $SqlCmd.CommandText = "INSERT INTO UpdateId (UpdateId) VALUES ('" + $messages.Item($messages.Count -1).UpdateId + "')"
     $SqlCmd.ExecuteNonQuery() | Out-Null
    }

$SqlConnection.close()

$payload = @{ "parse_mode" = "Markdown"; "disable_web_page_preview" = "True" }
foreach ($result in $results)
    {
     $recipient = $result.chat_id.ToString()
     $textmessage = $result.text
     $URLS = "https://api.telegram.org/bot$token/sendMessage?chat_id=$recipient&text=$textmessage"
     $Request = Invoke-WebRequest -Uri $URLS -Method Post -ContentType "application/json; charset=utf-8" -Body (ConvertTo-Json -Compress -InputObject $payload) -Proxy $proxy
    }
