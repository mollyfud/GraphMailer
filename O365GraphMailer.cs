using Azure.Identity;
using System;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace GraphMailer
{
    /// <summary>
    /// A facade class that simplifies sending emails through the Microsoft Graph API.
    /// This is the main entry point for the GraphMailer library.
    /// </summary>
    public class O365GraphMailer
    {
        private readonly GraphAuth _graphAuth;
        private readonly GraphEmailSender _emailSender;

        /// <summary>
        /// Initializes a new instance of the <see cref="O365GraphMailer"/> class.
        /// </summary>
        /// <param name="authData">The authentication data required for Microsoft Graph API.</param>
        public O365GraphMailer(AuthenticationData authData)
        {
            _graphAuth = new GraphAuth(authData);
            var graphClient = _graphAuth.GetAuthenticatedGraphClient();
            _emailSender = new GraphEmailSender(graphClient, message => GraphAuth.WriteLogEntry("Info", message));
        }

        /// <summary>
        /// Enables verbose logging for the Azure SDK to a specified file.
        /// </summary>
        /// <param name="logFilePath">The full path to the log file.</param>
        public static void EnableLogging(string logFilePath)
        {
            GraphAuth.EnableLogging(logFilePath);
        }

        /// <summary>
        /// Disables logging.
        /// </summary>
        public static void DisableLogging()
        {
            GraphAuth.DisableLogging();
        }

        /// <summary>
        /// Sends an email asynchronously.
        /// </summary>
        /// <param name="fromAddress">The sender's email address.</param>
        /// <param name="toAddress">The recipient's email address.</param>
        /// <param name="subject">The subject of the email.</param>
        /// <param name="body">The body of the email, which can be HTML.</param>
        /// <param name="saveToSentItems">A boolean indicating whether to save the email to the Sent Items folder. Default is true.</param>
        /// <returns>A task that represents the asynchronous operation.</returns>
        public async Task SendEmailAsync(string fromAddress, string toAddress, string subject, string body, bool saveToSentItems = true)
        {
            try
            {
                await _emailSender.SendEmailAsync(fromAddress, toAddress, subject, body, saveToSentItems);
            }
            catch (AuthenticationFailedException ex)
            {
                string detailedMessage = "Authentication failed. Please check your credentials and ensure admin consent has been granted for the 'Mail.Send' permission in Azure AD.";
                if (ex.InnerException != null)
                {
                    detailedMessage += "\n\nInner Exception Details: " + ex.InnerException.Message;
                }
                throw new Exception(detailedMessage, ex);
            }
        }

        /// <summary>
        /// Sends an email based on a MailMessage object asynchronously.
        /// </summary>
        /// <param name="message">The MailMessage object to send.</param>
        /// <param name="saveToSentItems">A boolean indicating whether to save the email to the Sent Items folder. Default is true.</param>
        /// <returns>A task that represents the asynchronous operation.</returns>
        public async Task SendEmailAsync(MailMessage message, bool saveToSentItems = true)
        {
            try
            {
                await _emailSender.SendEmailAsync(message, saveToSentItems);
            }
            catch (AuthenticationFailedException ex)
            {
                string detailedMessage = "Authentication failed. Please check your credentials and ensure admin consent has been granted for the 'Mail.Send' permission in Azure AD.";
                if (ex.InnerException != null)
                {
                    detailedMessage += "\n\nInner Exception Details: " + ex.InnerException.Message;
                }
                throw new Exception(detailedMessage, ex);
            }
            catch (Exception ex)
            {
                // Check for Access Denied which might indicate missing Mail.ReadWrite permission for large attachments
                if (ex.Message.Contains("Access is denied") || (ex.InnerException != null && ex.InnerException.Message.Contains("Access is denied")))
                {
                    throw new Exception("Access denied. Sending large attachments (>3MB) requires the 'Mail.ReadWrite' permission to create a draft message. Please ensure this permission is granted in Azure AD.", ex);
                }
                throw;
            }
        }

        /// <summary>
        /// Sends an email synchronously.
        /// </summary>
        /// <param name="fromAddress">The sender's email address.</param>
        /// <param name="toAddress">The recipient's email address.</param>
        /// <param name="subject">The subject of the email.</param>
        /// <param name="body">The body of the email, which can be HTML.</param>
        /// <param name="saveToSentItems">A boolean indicating whether to save the email to the Sent Items folder. Default is true.</param>
        public void SendEmail(string fromAddress, string toAddress, string subject, string body, bool saveToSentItems = true)
        {
            try
            {
                _emailSender.SendEmail(fromAddress, toAddress, subject, body, saveToSentItems);
            }
            catch (AuthenticationFailedException ex)
            {
                string detailedMessage = "Authentication failed. Please check your credentials and ensure admin consent has been granted for the 'Mail.Send' permission in Azure AD.";
                if (ex.InnerException != null)
                {
                    detailedMessage += "\n\nInner Exception Details: " + ex.InnerException.Message;
                }
                throw new Exception(detailedMessage, ex);
            }
        }

        /// <summary>
        /// Sends an email based on a MailMessage object synchronously.
        /// </summary>
        /// <param name="message">The MailMessage object to send.</param>
        /// <param name="saveToSentItems">A boolean indicating whether to save the email to the Sent Items folder. Default is true.</param>
        public void SendEmail(MailMessage message, bool saveToSentItems = true)
        {
            try
            {
                _emailSender.SendEmail(message, saveToSentItems);
            }
            catch (AuthenticationFailedException ex)
            {
                string detailedMessage = "Authentication failed. Please check your credentials and ensure admin consent has been granted for the 'Mail.Send' permission in Azure AD.";
                if (ex.InnerException != null)
                {
                    detailedMessage += "\n\nInner Exception Details: " + ex.InnerException.Message;
                }
                throw new Exception(detailedMessage, ex);
            }
            catch (Exception ex)
            {
                // Check for Access Denied which might indicate missing Mail.ReadWrite permission for large attachments
                if (ex.Message.Contains("Access is denied") || (ex.InnerException != null && ex.InnerException.Message.Contains("Access is denied")))
                {
                    throw new Exception("Access denied. Sending large attachments (>3MB) requires the 'Mail.ReadWrite' permission to create a draft message. Please ensure this permission is granted in Azure AD.", ex);
                }
                throw;
            }
        }

        /// <summary>
        /// Acquires a fresh access token and returns a diagnostic dump suitable for pasting into https://jwt.ms.
        /// The dump includes the raw JWT and a decoded summary that highlights the 'roles' claim,
        /// which must contain 'Mail.Send' for email sending to work.
        /// Optionally writes the dump to a file.
        /// </summary>
        /// <param name="filePath">Optional file path to write the dump to.</param>
        /// <returns>The diagnostic dump string.</returns>
        public async Task<string> GetTokenDumpAsync(string filePath = null)
        {
            var token = await _graphAuth.GetTokenAsync().ConfigureAwait(false);
            var dump = BuildTokenDump(token);

            if (!string.IsNullOrEmpty(filePath))
            {
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    Directory.CreateDirectory(directory);
                File.WriteAllText(filePath, dump);
            }

            return dump;
        }

        /// <summary>
        /// Acquires a fresh access token and returns a diagnostic dump suitable for pasting into https://jwt.ms.
        /// The dump includes the raw JWT and a decoded summary that highlights the 'roles' claim,
        /// which must contain 'Mail.Send' for email sending to work.
        /// Optionally writes the dump to a file.
        /// </summary>
        /// <param name="filePath">Optional file path to write the dump to.</param>
        /// <returns>The diagnostic dump string.</returns>
        public string GetTokenDump(string filePath = null)
        {
            return GetTokenDumpAsync(filePath).GetAwaiter().GetResult();
        }

        private static string BuildTokenDump(string token)
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== RAW TOKEN - paste into https://jwt.ms ===");
            sb.AppendLine();
            sb.AppendLine(token);
            sb.AppendLine();
            sb.AppendLine("=== DECODED PAYLOAD ===");

            try
            {
                var parts = token.Split('.');
                if (parts.Length < 2)
                {
                    sb.AppendLine("(Token does not appear to be a valid JWT)");
                    return sb.ToString();
                }

                // Base64url → base64
                var payload = parts[1].Replace('-', '+').Replace('_', '/');
                switch (payload.Length % 4)
                {
                    case 2: payload += "=="; break;
                    case 3: payload += "="; break;
                }

                var json = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(payload));

                using (var doc = System.Text.Json.JsonDocument.Parse(json))
                {
                    sb.AppendLine(System.Text.Json.JsonSerializer.Serialize(
                        doc.RootElement,
                        new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));

                    sb.AppendLine("=== KEY CLAIMS ===");

                    if (doc.RootElement.TryGetProperty("roles", out var roles))
                    {
                        sb.AppendLine($"roles : {roles}");
                        if (roles.GetArrayLength() == 0)
                            sb.AppendLine("  *** WARNING: 'roles' is EMPTY. Mail.Send has not been granted/consented in Azure AD. ***");
                        else if (!roles.ToString().Contains("Mail.Send"))
                            sb.AppendLine("  *** WARNING: 'Mail.Send' is not in 'roles'. Grant the Mail.Send Application permission and re-consent. ***");
                        else
                            sb.AppendLine("  OK: Mail.Send is present.");
                    }
                    else
                    {
                        sb.AppendLine("roles : (claim not present)");
                        sb.AppendLine("  *** WARNING: No 'roles' claim. Mail.Send Application permission has not been granted/consented in Azure AD. ***");
                    }

                    if (doc.RootElement.TryGetProperty("exp", out var exp))
                        sb.AppendLine($"exp   : {DateTimeOffset.FromUnixTimeSeconds(exp.GetInt64()):yyyy-MM-dd HH:mm:ss} UTC");

                    if (doc.RootElement.TryGetProperty("appid", out var appId))
                        sb.AppendLine($"appid : {appId.GetString()}");

                    if (doc.RootElement.TryGetProperty("tid", out var tid))
                        sb.AppendLine($"tid   : {tid.GetString()}");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"(Could not decode token payload: {ex.Message})");
            }

            return sb.ToString();
        }
    }
}