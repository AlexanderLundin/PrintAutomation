using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Net;
using System.Printing;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace PrintAutomation
{
    class Program
    {
        // Define the scope for Gmail API
        static string[] Scopes = { GmailService.Scope.GmailModify };
        static string ApplicationName = "PrintAutomationClient";

        static async Task Main(string[] args)
        {
            UserCredential credential;

            // Create or load the user credentials file
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true));
            }

            // Create the Gmail API service
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Set up the query parameters to search for PDF attachments
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List("me");
            //build query
            var query = new List<string>();
            //query.Add("is:unread");
            query.Add("subject:\"print\"");
            query.Add("has:attachment filename:pdf");
            string q = string.Join(" ", query);
            request.Q = q;
            request.MaxResults = 10;

            while (true)
            {
                //sleep 10 seconds
                System.Threading.Thread.Sleep(10000);
                // Retrieve the emails matching the query
                ListMessagesResponse response = request.Execute();
                if (response != null && response.Messages != null)
                {
                    ProcessEmails(service, response);
                }
            }

        }

        private static void ProcessEmails(GmailService service, ListMessagesResponse response)
        {
            foreach (var email in response.Messages)
            {
                // Get the email message
                Message message = service.Users.Messages.Get("me", email.Id).Execute();
                // Get the PDF attachment
                var attachment = message.Payload?.Parts?.FirstOrDefault(x => x.Filename.EndsWith(".pdf"));
                if (attachment != null)
                {
                    string filename = Path.Combine(@"E:\PDFs", attachment.Filename);
                    ProcessMessageParts(service, message, filename);

                }
            }
        }

        private static void ProcessMessageParts(GmailService service, Message message, string filename)
        {


            foreach (var part in message.Payload.Parts)
            {
                if (part.Filename != null && part.Body.AttachmentId != null)
                {
                    // This is a regular attachment
                    var attachment = service.Users.Messages.Attachments.Get("me", message.Id, part.Body.AttachmentId).Execute();
                    var attachmentData = attachment.Data;
                    // process attachment data...
                    PrintDataAsPdf(service, message, filename, attachmentData);

                }
                else if (part.Filename != null && part.Body.Data != null)
                {
                    throw new Exception("uncomplete code");
                    // This is a multi-part message or inline image
                    var data1 = part.Body.Data;
                    if (part.Body.AttachmentId != null)
                    {
                        // This is an inline image
                        var attachment1 = service.Users.Messages.Attachments.Get("me", message.Id, part.Body.AttachmentId).Execute();
                        data1 = attachment1.Data;
                    }
                    // process attachment data...
                }
            }
        }

        private static void PrintDataAsPdf(GmailService service, Message message, string filename, string attachmentData)
        {
            // Convert the attachment data from base64url to byte array
            byte[] convertedByteArray = Convert.FromBase64String(attachmentData.Replace('-', '+').Replace('_', '/'));


            // Save the PDF attachment to disk
            using (FileStream stream = new FileStream(filename, FileMode.Create, FileAccess.Write))
            {
                stream.Write(convertedByteArray, 0, convertedByteArray.Length);
            }

            // Set up the print document
            PrintDocument printDoc = new PrintDocument();
            printDoc.PrinterSettings.PrinterName = "HPAB4538 (HP DeskJet 3700 series)";
            printDoc.PrinterSettings.Copies = 1;
            printDoc.DefaultPageSettings.Color = false;
            printDoc.DefaultPageSettings.PaperSize = new PaperSize("Letter", 850, 1100);

            // Print the PDF file
            printDoc.PrinterSettings.PrintFileName = filename;
            printDoc.Print();


            // Modify the message label to mark it as read
            var mods = new ModifyMessageRequest();
            mods.RemoveLabelIds = new List<string>() { "UNREAD" };
            service.Users.Messages.Modify(mods, "me", message.Id).Execute();
        }
    }
}
