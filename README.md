# GraphMailer Library

A .NET library for sending emails via the Microsoft Graph API, designed to be easily consumed by both .NET and OpenEdge ABL applications.

## Prerequisites

Before using this library, you need to register an application in Azure Active Directory (Azure AD) and grant it the necessary permissions to send emails.

1.  **Register an application** in your Azure AD tenant.
2.  **Create a client secret** for the application.
3.  **Grant API Permissions**: Add the `Mail.Send` permission from the `Microsoft Graph` API (Application permission type).
    *   **Note:** If you intend to send attachments larger than 3MB, you must also grant the `Mail.ReadWrite` permission. This is required to create draft messages for the large attachment upload process.
4.  **Grant Admin Consent**: An administrator must grant admin consent for the selected permissions in your tenant.

You will need the following three pieces of information from your Azure AD application:
-   Tenant ID
-   Client ID
-   Client Secret

## .NET Usage

Here are examples of how to use the `GraphMailer` library in a .NET application.

### 1. Add a reference to the `GraphMailer.dll` in your project.

### 2. Sending a Simple Email

```csharp
using GraphMailer;
using System;
using System.Threading.Tasks;

public class EmailExample
{
    public async Task SendGraphEmail()
    {
        // 1. Populate your authentication data from your Azure AD app.
        var authData = new AuthenticationData
        {
            TenantId = "YOUR_TENANT_ID",
            ClientId = "YOUR_CLIENT_ID",
            ClientSecret = "YOUR_CLIENT_SECRET"
        };

        // 2. Create an instance of the mailer.
        var mailer = new O365GraphMailer(authData);

        // 3. Send the email.
        try
        {
            await mailer.SendEmailAsync(
                fromAddress: "sender@yourdomain.com",
                toAddress: "recipient@anotherdomain.com",
                subject: "Hello from GraphMailer",
                body: "<h1>Hello World!</h1><p>This is an email sent via the Microsoft Graph API.</p>",
                saveToSentItems: true
            );
            Console.WriteLine("Email sent successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending email: {ex.Message}");
        }
    }
}
```

### 3. Sending an Email with Attachments

You can also send an email by creating a `System.Net.Mail.MailMessage` object, which allows you to include attachments.

```csharp
using GraphMailer;
using System;
using System.Net.Mail;
using System.IO;
using System.Threading.Tasks;

public class EmailWithAttachmentExample
{
    public async Task SendGraphEmailWithAttachment()
    {
        // 1. Populate your authentication data.
        var authData = new AuthenticationData
        {
            TenantId = "YOUR_TENANT_ID",
            ClientId = "YOUR_CLIENT_ID",
            ClientSecret = "YOUR_CLIENT_SECRET"
        };

        // 2. Create an instance of the mailer.
        var mailer = new O365GraphMailer(authData);

        // 3. Create a MailMessage object.
        var mailMessage = new MailMessage
        {
            From = new MailAddress("sender@yourdomain.com"),
            Subject = "Email with Attachment",
            Body = "<h1>Hello!</h1><p>This email contains an attachment.</p>",
            IsBodyHtml = true
        };
        mailMessage.To.Add("recipient@anotherdomain.com");

        // 4. Add an attachment.
        // Create a dummy file for demonstration purposes.
        var filePath = Path.Combine(Path.GetTempPath(), "attachment.txt");
        File.WriteAllText(filePath, "This is the content of the attachment.");
        mailMessage.Attachments.Add(new Attachment(filePath));

        // 5. Send the email.
        try
        {
            await mailer.SendEmailAsync(mailMessage, saveToSentItems: true);
            Console.WriteLine("Email with attachment sent successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending email: {ex.Message}");
        }
        finally
        {
            // Clean up the dummy file.
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
    }
}
```

## Logging

You can enable verbose logging to diagnose issues with the Graph API or authentication. This writes detailed logs from the Azure SDK to a specified file.

### .NET

```csharp
// Enable logging
O365GraphMailer.EnableLogging(@"C:\Temp\GraphMailer.log");

// ... perform operations ...

// Disable logging
O365GraphMailer.DisableLogging();
```

### OpenEdge ABL

```abl
// Enable logging
GraphMailer.O365GraphMailer:EnableLogging("C:\Temp\GraphMailer.log").

// ... perform operations ...

// Disable logging
GraphMailer.O365GraphMailer:DisableLogging().
```

## Token Diagnostics

If you receive a `401 Unauthorized` error from the Graph API, you can dump the current access token to inspect exactly what permissions it contains. This is the most reliable way to confirm whether `Mail.Send` (and other required permissions) have been properly granted and consented in Azure AD.

The dump writes the raw JWT string — which you can paste directly into **[https://jwt.ms](https://jwt.ms)** — along with a decoded summary of the key claims, including a clear warning if `Mail.Send` is missing from the `roles` claim.

> **Note:** This method always acquires a **fresh** token directly from Azure AD, bypassing the internal MSAL cache. This means the dump always reflects the current state of your app registration, even if a cached token is still being used for sending.

### .NET

```csharp
var mailer = new O365GraphMailer(authData);

// Write the dump to a file and get the string back
string dump = await mailer.GetTokenDumpAsync(@"C:\Temp\token_dump.txt");
Console.WriteLine(dump);

// Or just get the string without writing to a file
string dump = await mailer.GetTokenDumpAsync();

// Synchronous version
string dump = mailer.GetTokenDump(@"C:\Temp\token_dump.txt");
```

### OpenEdge ABL

```abl
DEFINE VARIABLE cDump AS CHARACTER NO-UNDO.

// Write to file and return the dump string
cDump = oMailer:GetTokenDump("C:\Temp\token_dump.txt").
MESSAGE cDump VIEW-AS ALERT-BOX.

// Or without a file
cDump = oMailer:GetTokenDump().
MESSAGE cDump VIEW-AS ALERT-BOX.
```

### Sample Output

```
=== RAW TOKEN - paste into https://jwt.ms ===

eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIs...

=== DECODED PAYLOAD ===
{
  "aud": "https://graph.microsoft.com",
  "appid": "your-client-id",
  "tid": "your-tenant-id",
  "roles": [
    "Mail.Send"
  ],
  "exp": 1741737262,
  ...
}
=== KEY CLAIMS ===
roles : ["Mail.Send"]
  OK: Mail.Send is present.
exp   : 2026-03-11 22:34:22 UTC
appid : your-client-id
tid   : your-tenant-id
```

If `Mail.Send` is missing from `roles`, you will see a warning like:

```
roles : []
  *** WARNING: 'roles' is EMPTY. Mail.Send has not been granted/consented in Azure AD. ***
```

In that case, go to **Azure AD → App registrations → [Your App] → API permissions**, confirm `Mail.Send` is listed as an **Application** permission, and click **Grant admin consent**.

## OpenEdge ABL Usage

You can consume the .NET library directly from OpenEdge ABL.

### 1. Deployment & Dependencies

Place the `GraphMailer.dll` and all its dependencies in a location accessible by your OpenEdge application (e.g., your assemblies directory).

**Note on Assembly Resolution:** This library includes a built-in assembly resolver. This helps avoid version conflicts and eliminates the need for complex `app.config` binding redirects for standard dependencies. Ensure all the following DLLs (and their dependencies) are present in the same directory:

*   `Azure.Core.dll`
*   `Azure.Identity.dll`
*   `Microsoft.Bcl.AsyncInterfaces.dll`
*   `Microsoft.Graph.dll`
*   `Microsoft.Graph.Core.dll`
*   `Microsoft.Identity.Client.dll`
*   `System.Text.Json.dll`
*   `System.Memory.dll`
*   ...and other dependencies from the build output.

### 2. Sending a Simple Email

Make sure your `prowin.exe.config` (or equivalent) is configured to allow .NET integration.

```abl
USING System.*.
USING System.Threading.Tasks.*.
USING GraphMailer.*.

BLOCK-LEVEL ON ERROR UNDO, THROW.

// 1. Populate your authentication data from your Azure AD app.
DEFINE VARIABLE oAuthData      AS CLASS AuthenticationData NO-UNDO.
DEFINE VARIABLE oMailer        AS CLASS O365GraphMailer    NO-UNDO.
DEFINE VARIABLE oException     AS CLASS Exception          NO-UNDO.

// Optional: Enable logging to debug issues
// O365GraphMailer:EnableLogging("C:\Temp\GraphMailer.log").

oAuthData = NEW AuthenticationData().
oAuthData:TenantId = "YOUR_TENANT_ID".
oAuthData:ClientId = "YOUR_CLIENT_ID".
oAuthData:ClientSecret = "YOUR_CLIENT_SECRET".

// 2. Create an instance of the mailer.
oMailer = NEW O365GraphMailer(oAuthData).

// 3. Define email parameters.
DEFINE VARIABLE cFromAddress AS CHARACTER NO-UNDO INITIAL "sender@yourdomain.com".
DEFINE VARIABLE cToAddress   AS CHARACTER NO-UNDO INITIAL "recipient@anotherdomain.com".
DEFINE VARIABLE cSubject     AS CHARACTER NO-UNDO INITIAL "Hello from GraphMailer via OpenEdge".
DEFINE VARIABLE cBody        AS CHARACTER NO-UNDO INITIAL "<h1>Hello World!</h1><p>This is an email sent from ABL via the Microsoft Graph API.</p>".
DEFINE VARIABLE lSave        AS LOGICAL   NO-UNDO INITIAL TRUE.

// 4. Send the email synchronously.
oMailer:SendEmail(cFromAddress, cToAddress, cSubject, cBody, lSave).

MESSAGE "Email sent successfully!" VIEW-AS ALERT-BOX.

CATCH oException AS Progress.Lang.Error:
    // Catching Progress.Lang.Error is a good practice for .NET exceptions
    MESSAGE "Error sending email:" oException:GetMessage()
        VIEW-AS ALERT-BOX.
END CATCH.

FINALLY:
    // Clean up .NET objects
    IF VALID-OBJECT(oAuthData) THEN DELETE OBJECT oAuthData.
    IF VALID-OBJECT(oMailer) THEN DELETE OBJECT oMailer.
END FINALLY.
```

### 3. Sending an Email with Attachments

```abl
USING System.*.
USING System.IO.*.
USING System.Net.Mail.*.
USING GraphMailer.*.

BLOCK-LEVEL ON ERROR UNDO, THROW.

// 1. Populate authentication data.
DEFINE VARIABLE oAuthData      AS CLASS AuthenticationData NO-UNDO.
DEFINE VARIABLE oMailer        AS CLASS O365GraphMailer    NO-UNDO.
DEFINE VARIABLE oMailMessage   AS CLASS MailMessage        NO-UNDO.
DEFINE VARIABLE oAttachment    AS CLASS Attachment         NO-UNDO.
DEFINE VARIABLE cAttachmentPath AS CHARACTER              NO-UNDO.
DEFINE VARIABLE oException     AS CLASS Exception          NO-UNDO.

oAuthData = NEW AuthenticationData().
oAuthData:TenantId = "YOUR_TENANT_ID".
oAuthData:ClientId = "YOUR_CLIENT_ID".
oAuthData:ClientSecret = "YOUR_CLIENT_SECRET".

// 2. Create mailer instance.
oMailer = NEW O365GraphMailer(oAuthData).

// 3. Create MailMessage object.
oMailMessage = NEW MailMessage(
    "sender@yourdomain.com",
    "recipient@anotherdomain.com",
    "Email with Attachment from ABL",
    "<h1>Hello!</h1><p>This email contains an attachment sent from OpenEdge ABL.</p>"
).
oMailMessage:IsBodyHtml = TRUE.

// 4. Add an attachment.
cAttachmentPath = System.IO.Path:Combine(System.IO.Path:GetTempPath(), "attachment.txt").
System.IO.File:WriteAllText(cAttachmentPath, "This is the content of the attachment from ABL.").
oAttachment = NEW Attachment(cAttachmentPath).
oMailMessage:Attachments:Add(oAttachment).

// 5. Send the email.
oMailer:SendEmail(oMailMessage, TRUE).

MESSAGE "Email with attachment sent successfully!" VIEW-AS ALERT-BOX.

CATCH oException AS Progress.Lang.Error:
    MESSAGE "Error sending email:" oException:GetMessage()
        VIEW-AS ALERT-BOX.
END CATCH.

FINALLY:
    // Clean up .NET objects and the dummy file.
    IF VALID-OBJECT(oAuthData) THEN DELETE OBJECT oAuthData.
    IF VALID-OBJECT(oMailer) THEN DELETE OBJECT oMailer.
    IF VALID-OBJECT(oMailMessage) THEN DELETE OBJECT oMailMessage.
    IF System.IO.File:Exists(cAttachmentPath) THEN
        System.IO.File:Delete(cAttachmentPath).
END FINALLY.
```

## Troubleshooting

### Authentication Failed
If you receive an "Authentication failed" error:
1.  Verify your **Tenant ID**, **Client ID**, and **Client Secret**.
2.  Ensure the application in Azure AD has the **`Mail.Send`** permission (Application type).
3.  **Important:** Ensure **Admin Consent** has been granted for these permissions.

### Access Denied (Large Attachments)
If you receive an "Access is denied" error when sending attachments larger than 3MB:
1.  The library uses the Microsoft Graph "large attachment" API (upload session) for files > 3MB.
2.  This process requires the **`Mail.ReadWrite`** permission (Application type) to create and modify a draft message.
3.  Grant this permission in Azure AD and ensure Admin Consent is granted.

## Extensibility

This library is designed to be simple and straightforward. You can extend it in several ways:

*   **Error Handling**: Implement more robust error handling and logging.
*   **Configuration**: Instead of hardcoding authentication data, load it from a configuration file (`app.config`, `web.config`, etc.).
*   **Dependency Injection**: In applications that support it, you can register `O365GraphMailer` and its dependencies with a dependency injection container.

## Signing the DLL

To sign the DLL with a strong name key:
1. Generate a key file: `sn -k GraphMailer.snk`
2. The project is already configured to use this file.

