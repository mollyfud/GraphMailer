```mermaid
classDiagram
    direction LR

    class AuthenticationData {
        +string TenantId
        +string ClientId
        +string ClientSecret
    }

    class GraphAuth {
        -AuthenticationData _authData
        -ConcurrentDictionary _clientCache
        +static EnableLogging(string logFilePath)
        +static DisableLogging()
        +GetAuthenticatedGraphClient() GraphServiceClient
    }

    class GraphEmailSender {
        -GraphServiceClient _graphClient
        -static HttpClient _httpClient
        +SendEmailAsync(...) Task
        +SendEmail(...) void
        -SendEmailWithLargeAttachmentsAsync(...) Task
        -UploadSmallAttachmentAsync(...) Task
        -UploadLargeAttachmentAsync(...) Task
    }

    class O365GraphMailer {
        -GraphEmailSender _emailSender
        +static EnableLogging(string logFilePath)
        +static DisableLogging()
        +SendEmailAsync(...) Task
        +SendEmail(...) void
    }

    O365GraphMailer ..> AuthenticationData : Uses
    O365GraphMailer ..> GraphAuth : Creates (Internal)
    O365GraphMailer --> GraphEmailSender : Composes
    GraphAuth --> AuthenticationData : Holds
    GraphAuth ..> GraphServiceClient : Creates
    GraphEmailSender ..> GraphServiceClient : Uses
    ```