using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using Microsoft.Graph.Users.Item.Messages.Item.Attachments.CreateUploadSession;
using System.Text;
using Microsoft.Kiota.Abstractions;

namespace GraphMailer
{
    /// <summary>
    /// Handles sending emails using the Microsoft Graph API.
    /// </summary>
    public class GraphEmailSender
    {
        private readonly GraphServiceClient _graphClient;
        private static readonly HttpClient _httpClient = new HttpClient();
        private readonly Action<string> _log;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphEmailSender"/> class.
        /// </summary>
        /// <param name="graphClient">An authenticated <see cref="GraphServiceClient"/>.</param>
        /// <param name="log">An optional log writer for diagnostics.</param>
        public GraphEmailSender(GraphServiceClient graphClient, Action<string> log = null)
        {
            if (graphClient == null)
            {
                throw new ArgumentNullException(nameof(graphClient));
            }
            _graphClient = graphClient;
            _log = log;
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
            var message = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = body
                },
                ToRecipients = new List<Recipient>
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = toAddress
                        }
                    }
                }
            };

            var sendMailBody = new SendMailPostRequestBody
            {
                Message = message,
                SaveToSentItems = saveToSentItems
            };

            try
            {
                await _graphClient.Users[fromAddress]
                    .SendMail
                    .PostAsync(sendMailBody)
                    .ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogGraphFailure("SendMail", fromAddress, message, ex, null, null, 0, saveToSentItems);
                throw;
            }
        }

        /// <summary>
        /// Sends an email asynchronously.
        /// </summary>
        /// <param name="message">The MailMessage object to send.</param>
        /// <param name="saveToSentItems">A boolean indicating whether to save the email to the Sent Items folder. Default is true.</param>
        /// <returns>A task that represents the asynchronous operation.</returns>
        public async Task SendEmailAsync(MailMessage message, bool saveToSentItems = true)
        {
            if (message == null)
            {
                throw new ArgumentNullException(nameof(message));
            }

            if (message.From == null)
            {
                throw new ArgumentException("The 'From' address cannot be null in the MailMessage.", nameof(message));
            }

            long totalSize = GetTotalAttachmentSize(message);
            const long MaxSimpleSize = 3 * 1024 * 1024;
            bool usesLargeAttachmentWorkflow = totalSize > MaxSimpleSize || message.Attachments.Any(a => a.ContentStream.Length > MaxSimpleSize);

            if (usesLargeAttachmentWorkflow)
            {
                await SendEmailWithLargeAttachmentsAsync(message, saveToSentItems, totalSize).ConfigureAwait(false);
                return;
            }

            var graphMessage = new Message
            {
                Subject = message.Subject,
                Body = new ItemBody
                {
                    ContentType = message.IsBodyHtml ? BodyType.Html : BodyType.Text,
                    Content = message.Body
                },
                ToRecipients = message.To.Select(r => new Recipient { EmailAddress = new EmailAddress { Address = r.Address } }).ToList(),
                CcRecipients = message.CC.Select(r => new Recipient { EmailAddress = new EmailAddress { Address = r.Address } }).ToList(),
                BccRecipients = message.Bcc.Select(r => new Recipient { EmailAddress = new EmailAddress { Address = r.Address } }).ToList()
            };

            if (message.Attachments.Any())
            {
                var attachments = new List<Microsoft.Graph.Models.Attachment>();
                foreach (var attachment in message.Attachments)
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        if (attachment.ContentStream.CanSeek)
                        {
                            attachment.ContentStream.Position = 0;
                        }
                        attachment.ContentStream.CopyTo(memoryStream);
                        var fileAttachment = new FileAttachment
                        {
                            Name = attachment.Name,
                            ContentType = attachment.ContentType.MediaType,
                            ContentBytes = memoryStream.ToArray()
                        };
                        attachments.Add(fileAttachment);
                    }
                }
                graphMessage.Attachments = attachments.ToList<Microsoft.Graph.Models.Attachment>();
            }

            var sendMailBody = new SendMailPostRequestBody
            {
                Message = graphMessage,
                SaveToSentItems = saveToSentItems
            };

            try
            {
                await _graphClient.Users[message.From.Address]
                    .SendMail
                    .PostAsync(sendMailBody)
                    .ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogGraphFailure("SendMail", message.From.Address, graphMessage, ex, message, null, totalSize, saveToSentItems);
                throw;
            }
        }

        private async Task SendEmailWithLargeAttachmentsAsync(MailMessage message, bool saveToSentItems, long totalAttachmentSize)
        {
            var draftMessage = new Message
            {
                Subject = message.Subject,
                Body = new ItemBody
                {
                    ContentType = message.IsBodyHtml ? BodyType.Html : BodyType.Text,
                    Content = message.Body
                },
                ToRecipients = message.To.Select(r => new Recipient { EmailAddress = new EmailAddress { Address = r.Address } }).ToList(),
                CcRecipients = message.CC.Select(r => new Recipient { EmailAddress = new EmailAddress { Address = r.Address } }).ToList(),
                BccRecipients = message.Bcc.Select(r => new Recipient { EmailAddress = new EmailAddress { Address = r.Address } }).ToList()
            };

            string messageId = null;

            try
            {
                var createdMessage = await _graphClient.Users[message.From.Address].Messages.PostAsync(draftMessage).ConfigureAwait(false);
                messageId = createdMessage.Id;
            }
            catch (Exception ex)
            {
                LogGraphFailure("CreateDraft", message.From.Address, draftMessage, ex, message, null, totalAttachmentSize, saveToSentItems);
                throw;
            }

            try
            {
                foreach (var attachment in message.Attachments)
                {
                    if (attachment.ContentStream.CanSeek)
                    {
                        attachment.ContentStream.Position = 0;
                    }

                    if (attachment.ContentStream.Length < 3 * 1024 * 1024)
                    {
                        await UploadSmallAttachmentAsync(message.From.Address, messageId, attachment, message, draftMessage, totalAttachmentSize, saveToSentItems).ConfigureAwait(false);
                    }
                    else
                    {
                        await UploadLargeAttachmentAsync(message.From.Address, messageId, attachment, message, draftMessage, totalAttachmentSize, saveToSentItems).ConfigureAwait(false);
                    }
                }

                try
                {
                    await _graphClient.Users[message.From.Address].Messages[messageId].Send.PostAsync().ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    LogGraphFailure("SendDraft", message.From.Address, draftMessage, ex, message, messageId, totalAttachmentSize, saveToSentItems);
                    throw;
                }
            }
            catch
            {
                try
                {
                    await _graphClient.Users[message.From.Address].Messages[messageId].DeleteAsync().ConfigureAwait(false);
                }
                catch (Exception cleanupEx)
                {
                    WriteDiagnostic("Warning", $"Failed to delete draft message after send failure. User: {message.From.Address}, DraftId: {messageId ?? "(null)"}. {cleanupEx}");
                }
                throw;
            }
        }

        private async Task UploadSmallAttachmentAsync(string userId, string messageId, System.Net.Mail.Attachment attachment, MailMessage originalMessage, Message draftMessage, long totalAttachmentSize, bool saveToSentItems)
        {
            using (var memoryStream = new MemoryStream())
            {
                try
                {
                    await attachment.ContentStream.CopyToAsync(memoryStream).ConfigureAwait(false);
                    var fileAttachment = new FileAttachment
                    {
                        Name = attachment.Name,
                        ContentType = attachment.ContentType.MediaType,
                        ContentBytes = memoryStream.ToArray()
                    };

                    await _graphClient.Users[userId].Messages[messageId].Attachments.PostAsync(fileAttachment).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    LogGraphFailure("UploadSmallAttachment", userId, draftMessage, ex, originalMessage, messageId, totalAttachmentSize, saveToSentItems, attachment.Name, attachment.ContentStream.Length);
                    throw;
                }
            }
        }

        private async Task UploadLargeAttachmentAsync(string userId, string messageId, System.Net.Mail.Attachment attachment, MailMessage originalMessage, Message draftMessage, long totalAttachmentSize, bool saveToSentItems)
        {
            var attachmentItem = new AttachmentItem
            {
                AttachmentType = AttachmentType.File,
                Name = attachment.Name,
                Size = attachment.ContentStream.Length,
                ContentType = attachment.ContentType.MediaType
            };

            var uploadSessionRequestBody = new CreateUploadSessionPostRequestBody
            {
                AttachmentItem = attachmentItem
            };

            try
            {
                var uploadSession = await _graphClient.Users[userId].Messages[messageId].Attachments.CreateUploadSession.PostAsync(uploadSessionRequestBody).ConfigureAwait(false);
                int chunkSize = 320 * 1024 * 12;
                var stream = attachment.ContentStream;
                if (stream.CanSeek)
                {
                    stream.Position = 0;
                }

                long totalLength = stream.Length;
                byte[] buffer = new byte[chunkSize];
                long uploadedBytes = 0;

                while (uploadedBytes < totalLength)
                {
                    int bytesRead = await stream.ReadAsync(buffer, 0, chunkSize).ConfigureAwait(false);
                    if (bytesRead <= 0)
                    {
                        break;
                    }

                    using (var request = new HttpRequestMessage(HttpMethod.Put, uploadSession.UploadUrl))
                    {
                        request.Content = new ByteArrayContent(buffer, 0, bytesRead);
                        request.Content.Headers.ContentRange = new ContentRangeHeaderValue(uploadedBytes, uploadedBytes + bytesRead - 1, totalLength);
                        request.Content.Headers.ContentLength = bytesRead;

                        using (var response = await _httpClient.SendAsync(request).ConfigureAwait(false))
                        {
                            if (!response.IsSuccessStatusCode)
                            {
                                string responseBody = response.Content == null
                                    ? string.Empty
                                    : await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                                throw new InvalidOperationException($"Failed to upload chunk for attachment '{attachment.Name}'. Status: {(int)response.StatusCode} {response.ReasonPhrase}. Range: {uploadedBytes}-{uploadedBytes + bytesRead - 1}/{totalLength}. Response: {responseBody}");
                            }
                        }
                    }

                    uploadedBytes += bytesRead;
                }
            }
            catch (Exception ex)
            {
                LogGraphFailure("UploadLargeAttachment", userId, draftMessage, ex, originalMessage, messageId, totalAttachmentSize, saveToSentItems, attachment.Name, attachment.ContentStream.Length);
                throw;
            }
        }

        private static long GetTotalAttachmentSize(MailMessage message)
        {
            long totalSize = 0;
            if (message.Attachments.Any())
            {
                foreach (var attachment in message.Attachments)
                {
                    totalSize += attachment.ContentStream.Length;
                }
            }

            return totalSize;
        }

        private void LogGraphFailure(string stage, string sender, Message graphMessage, Exception exception, MailMessage originalMessage, string draftMessageId, long totalAttachmentSize, bool saveToSentItems, string attachmentName = null, long? attachmentSize = null)
        {
            var details = BuildFailureDetails(stage, sender, graphMessage, exception, originalMessage, draftMessageId, totalAttachmentSize, saveToSentItems, attachmentName, attachmentSize);
            WriteDiagnostic("Error", details);
        }

        private string BuildFailureDetails(string stage, string sender, Message graphMessage, Exception exception, MailMessage originalMessage, string draftMessageId, long totalAttachmentSize, bool saveToSentItems, string attachmentName, long? attachmentSize)
        {
            var builder = new StringBuilder();
            builder.AppendLine($"Mail send failed at stage '{stage}'.")
                .AppendLine($"Sender: {sender ?? "(null)"}")
                .AppendLine($"Subject: {graphMessage?.Subject ?? "(null)"}")
                .AppendLine($"SaveToSentItems: {saveToSentItems}")
                .AppendLine($"DraftMessageId: {draftMessageId ?? "(null)"}")
                .AppendLine($"To: {FormatRecipients(graphMessage?.ToRecipients)}")
                .AppendLine($"Cc: {FormatRecipients(graphMessage?.CcRecipients)}")
                .AppendLine($"Bcc: {FormatRecipients(graphMessage?.BccRecipients)}")
                .AppendLine($"AttachmentCount: {originalMessage?.Attachments.Count ?? 0}")
                .AppendLine($"TotalAttachmentSize: {totalAttachmentSize}");

            if (!string.IsNullOrWhiteSpace(attachmentName))
            {
                builder.AppendLine($"AttachmentName: {attachmentName}");
            }

            if (attachmentSize.HasValue)
            {
                builder.AppendLine($"AttachmentSize: {attachmentSize.Value}");
            }

            AppendExceptionDetails(builder, exception);
            return builder.ToString().TrimEnd();
        }

        private static string FormatRecipients(IEnumerable<Recipient> recipients)
        {
            if (recipients == null)
            {
                return "(none)";
            }

            var addresses = recipients
                .Select(r => r?.EmailAddress?.Address)
                .Where(address => !string.IsNullOrWhiteSpace(address))
                .ToList();

            return addresses.Any() ? string.Join(";", addresses) : "(none)";
        }

        private static void AppendExceptionDetails(StringBuilder builder, Exception exception)
        {
            int depth = 0;
            var current = exception;
            while (current != null)
            {
                builder.AppendLine($"Exception[{depth}] Type: {current.GetType().FullName}")
                    .AppendLine($"Exception[{depth}] Message: {current.Message}");

                var apiException = current as ApiException;
                if (apiException != null)
                {
                    builder.AppendLine($"Exception[{depth}] ResponseStatusCode: {apiException.ResponseStatusCode}");
                    if (apiException.ResponseHeaders != null)
                    {
                        AppendHeaderValue(builder, apiException.ResponseHeaders, depth, "request-id");
                        AppendHeaderValue(builder, apiException.ResponseHeaders, depth, "client-request-id");
                        AppendHeaderValue(builder, apiException.ResponseHeaders, depth, "x-ms-ags-diagnostic");
                        AppendHeaderValue(builder, apiException.ResponseHeaders, depth, "Retry-After");
                    }
                }

                current = current.InnerException;
                depth++;
            }

            builder.AppendLine("ExceptionDetail:")
                .AppendLine(exception.ToString());
        }

        private static void AppendHeaderValue(StringBuilder builder, IDictionary<string, IEnumerable<string>> headers, int depth, string headerName)
        {
            IEnumerable<string> values;
            if (headers.TryGetValue(headerName, out values) && values != null)
            {
                builder.AppendLine($"Exception[{depth}] {headerName}: {string.Join(",", values)}");
            }
        }

        private void WriteDiagnostic(string level, string message)
        {
            try
            {
                if (_log != null)
                {
                    _log(message);
                    return;
                }

                GraphAuth.WriteLogEntry(level, message);
            }
            catch
            {
                // Never let diagnostics break mail sending.
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
            RunSync(() => SendEmailAsync(fromAddress, toAddress, subject, body, saveToSentItems));
        }

        /// <summary>
        /// Sends an email synchronously.
        /// </summary>
        /// <param name="message">The MailMessage object to send.</param>
        /// <param name="saveToSentItems">A boolean indicating whether to save the email to the Sent Items folder. Default is true.</param>
        public void SendEmail(MailMessage message, bool saveToSentItems = true)
        {
            RunSync(() => SendEmailAsync(message, saveToSentItems));
        }

        /// <summary>
        /// Runs an async Task synchronously without causing deadlocks.
        /// </summary>
        /// <param name="task">A function that returns the Task to execute.</param>
        private static void RunSync(Func<Task> task)
        {
            var oldContext = System.Threading.SynchronizationContext.Current;
            var synch = new System.Threading.SynchronizationContext();
            System.Threading.SynchronizationContext.SetSynchronizationContext(synch);
            try
            {
                task().GetAwaiter().GetResult();
            }
            finally
            {
                System.Threading.SynchronizationContext.SetSynchronizationContext(oldContext);
            }
        }
    }
}
