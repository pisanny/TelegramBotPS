# PowerShell Scripts

- Get_ChatId_To_Email.ps1 - Получить ChatId Telegram по корпоративной почте с проверкой.

- Request_registration.ps1 - Запрос регистации на канал Telegram.

- Registration _Completed.ps1 - Окончание регистрациии.

- Command_handler.ps1 - Обработчик команд для чата.

Для Request_registration.ps1, Registration _Completed.ps1 и Command_handler.ps1 используется таблицы:

### chatusers

| chat_id | SID               | Registered | DisplayName      |
|---------|-------------------|------------|------------------|
|xxxxxxxxx|S-x-x-xx-xxxxxxxxxx| 0          | Vladimir Pisanny |

### scripts

| command  | scriptblock                              | 
|----------|------------------------------------------|
| gethello |$script:result = @{                       |
|          |        UpdateId = $message.UpdateId      |
|          |        text = 'Hello!'                   |
|          |        chat_id = $message.chat_id        |
|          |        Message_ID = $message.Message_ID  |
|          |        first_name = $message.first_name  |
|          |        last_name  = $message.last_name   |
|          |       }                                  |

### UpdateId

| UpdateId |
|----------|
| xxxxxxxx |

Для хранения паролей используется встроенная в ОС служба Windows CredMan. Используется скрипт [CredMan.ps1][] - автор Jim Harrison (jim@isatools.org). 

Для подключения к почтовому ящику Exchange используется [Microsoft Exchange Web Services Managed API 2.2][].

[CredMan.ps1]: https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Credentials-d44c3cde
[Microsoft Exchange Web Services Managed API 2.2]: http://techgenix.com/microsoft-exchange-web-services-managed-api-22-released/

- reg.ps1 - Реализация команды get, содержиться в таблице scripts.