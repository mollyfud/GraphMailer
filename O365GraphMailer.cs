using Azure.Identity;
using System;
using System.Net.Mail;
using System.Threading.Tasks;

namespace GraphMailer
{
    /// <summary>
    /// A facade class that simplifies sending emails through the Microsoft Graph API.
    /// This is the main entry point for the GraphMailer library.
    /// </summary>
    public class O365GraphMailer
    {
        private readonly GraphEmailSender _emailSender;

        /// <summary>
        /// Initializes a new instance of the <see cref="O365GraphMailer"/> class.
        /// </summary>
        /// <param name="authData">The authentication data required for Microsoft Graph API.</param>
        public O365GraphMailer(AuthenticationData authData)
        {
            var graphAuth = new GraphAuth(authData);
            var graphClient = graphAuth.GetAuthenticatedGraphClient();
            _emailSender = new GraphEmailSender(graphClient);
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
    }
}