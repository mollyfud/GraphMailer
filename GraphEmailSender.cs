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

namespace GraphMailer
{
    /// <summary>
    /// Handles sending emails using the Microsoft Graph API.
    /// </summary>
    public class GraphEmailSender
    {
        private readonly GraphServiceClient _graphClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphEmailSender"/> class.
        /// </summary>
        /// <param name="graphClient">An authenticated <see cref="GraphServiceClient"/>.</param>
        public GraphEmailSender(GraphServiceClient graphClient)
        {
            if (graphClient == null)
            {
                throw new ArgumentNullException(nameof(graphClient));
            }
            _graphClient = graphClient;
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

            await _graphClient.Users[fromAddress]
                .SendMail
                .PostAsync(sendMailBody)
                .ConfigureAwait(false);
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

            // Check if we need to use the large attachment workflow
            long totalSize = 0;
            if (message.Attachments.Any())
            {
                foreach (var attachment in message.Attachments)
                {
                    totalSize += attachment.ContentStream.Length;
                }
            }

            const long MaxSimpleSize = 3 * 1024 * 1024; // 3 MB

            if (totalSize > MaxSimpleSize || message.Attachments.Any(a => a.ContentStream.Length > MaxSimpleSize))
            {
                await SendEmailWithLargeAttachmentsAsync(message, saveToSentItems).ConfigureAwait(false);
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

            await _graphClient.Users[message.From.Address]
                .SendMail
                .PostAsync(sendMailBody)
                .ConfigureAwait(false);
        }

        private async Task SendEmailWithLargeAttachmentsAsync(MailMessage message, bool saveToSentItems)
        {
            // 1. Create Draft
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

            var createdMessage = await _graphClient.Users[message.From.Address].Messages.PostAsync(draftMessage).ConfigureAwait(false);
            string messageId = createdMessage.Id;

            try
            {
                // 2. Upload Attachments
                foreach (var attachment in message.Attachments)
                {
                    if (attachment.ContentStream.CanSeek)
                    {
                        attachment.ContentStream.Position = 0;
                    }

                    if (attachment.ContentStream.Length < 3 * 1024 * 1024)
                    {
                        await UploadSmallAttachmentAsync(message.From.Address, messageId, attachment).ConfigureAwait(false);
                    }
                    else
                    {
                        await UploadLargeAttachmentAsync(message.From.Address, messageId, attachment).ConfigureAwait(false);
                    }
                }

                // 3. Send Draft
                await _graphClient.Users[message.From.Address].Messages[messageId].Send.PostAsync().ConfigureAwait(false);
            }
            catch
            {
                // Cleanup: Delete the draft if something went wrong
                try
                {
                    await _graphClient.Users[message.From.Address].Messages[messageId].DeleteAsync().ConfigureAwait(false);
                }
                catch { /* Ignore cleanup errors */ }
                throw;
            }
        }

        private async Task UploadSmallAttachmentAsync(string userId, string messageId, System.Net.Mail.Attachment attachment)
        {
            using (var memoryStream = new MemoryStream())
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
        }

        private async Task UploadLargeAttachmentAsync(string userId, string messageId, System.Net.Mail.Attachment attachment)
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

            var uploadSession = await _graphClient.Users[userId].Messages[messageId].Attachments.CreateUploadSession.PostAsync(uploadSessionRequestBody).ConfigureAwait(false);

            // Chunked Upload Logic
            // 320 KB * 12 = 3840 KB = 3.75 MB chunks
            int chunkSize = 320 * 1024 * 12; 

            using (var httpClient = new HttpClient())
            {
                var stream = attachment.ContentStream;
                if (stream.CanSeek) stream.Position = 0;
                
                long totalLength = stream.Length;
                byte[] buffer = new byte[chunkSize];
                long uploadedBytes = 0;

                while (uploadedBytes < totalLength)
                {
                    int bytesRead = await stream.ReadAsync(buffer, 0, chunkSize).ConfigureAwait(false);
                    
                    using (var request = new HttpRequestMessage(HttpMethod.Put, uploadSession.UploadUrl))
                    {
                        request.Content = new ByteArrayContent(buffer, 0, bytesRead);
                        request.Content.Headers.ContentRange = new ContentRangeHeaderValue(uploadedBytes, uploadedBytes + bytesRead - 1, totalLength);
                        request.Content.Headers.ContentLength = bytesRead;

                        var response = await httpClient.SendAsync(request).ConfigureAwait(false);
                        if (!response.IsSuccessStatusCode)
                        {
                            throw new Exception($"Failed to upload chunk. Status: {response.StatusCode}, Reason: {response.ReasonPhrase}");
                        }
                    }

                    uploadedBytes += bytesRead;
                }
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
