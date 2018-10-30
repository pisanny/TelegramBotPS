# ps

Get_ChatId_To_Email.ps1 - Получить ChatId Telegram по корпоративной почте с проверкой.

Request_registration.ps1 - Запрос регистации на канал Telegram.

Registration _Completed.ps1 - Окончание регистрациии.

Для Request_registration.ps1 и Registration _Completed.ps1 используется таблица chatusers:

| chat_id | SID               | Registered | DisplayName      |
|---------|-------------------|------------|------------------|
|xxxxxxxxx|S-x-x-xx-xxxxxxxxxx| 0          | Vladimir Pisanny |

Для хранения паролей используется встроенная в ОС служба Windows CredMan. Используется скрипт [CredMan.ps1][] - автор Jim Harrison (jim@isatools.org). 

Для подключения к почтовому ящику Exchange используется [Microsoft Exchange Web Services Managed API 2.2][].

[CredMan.ps1]: https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Credentials-d44c3cde
[Microsoft Exchange Web Services Managed API 2.2]: http://techgenix.com/microsoft-exchange-web-services-managed-api-22-released/
