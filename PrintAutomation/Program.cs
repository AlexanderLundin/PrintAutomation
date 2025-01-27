using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Net;
using System.Printing;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using PdfiumViewer;

namespace PrintAutomation
{
    public class Program
    {
        // Define the scope for Gmail API
        static string[] Scopes = { GmailService.Scope.GmailModify };
        static string ApplicationName = "PrintAutomationClient";
        public static TextFileLog log = new TextFileLog();

        static async Task Main(string[] args)
        {
            try
            {
                //MessageBox.Show("main started");
                var exeDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                var logPath = exeDir + "\\" + "PrintLog.txt";
                log = new TextFileLog(logPath);
                log.Write("Program started.");

                GmailService service = await GetNewTokenWithUserInput();

                // Set up the query parameters to search for PDF attachments
                UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List("me");
                ProcessEmailsToPrintWithPdfAttachment(service, request);
                ProcessEmailsToPrintWithZInSubject(service, request);
                log.Write("Program completed.");
            }
            catch (Exception ex)
            {
                log.Exception(ex);
            }

        }

        private static void ProcessEmailsToPrintWithPdfAttachment(GmailService service, UsersResource.MessagesResource.ListRequest request)
        {
            //build query
            string q = GetEmailsWithPdfsToPrintQuery();
            request.Q = q;
            request.MaxResults = 10;

            // Retrieve the emails matching the query
            ListMessagesResponse response = request.Execute();
            if (response != null && response.Messages != null)
            {
                log.Write("found emails to print.");
                PrintPdfShippingLabelWorkflow.ProcessEmails(service, response);
            }
        }

        private static string GetEmailsWithPdfsToPrintQuery()
        {
            var query = new List<string>();
            query.Add("is:unread");
            query.Add("subject:\"print\"");
            query.Add("has:attachment filename:pdf");
            string q = string.Join(" ", query);
            return q;
        }


        private static void ProcessEmailsToPrintWithZInSubject(GmailService service, UsersResource.MessagesResource.ListRequest request)
        {
            string q = GetEmailsEmailsToPrintWithZInSubject();
            request.Q = q;
            request.MaxResults = 10;

            ListMessagesResponse response = request.Execute();
            if (response != null && response.Messages != null)
            {
                log.Write("found emails with z in subject to print.");
                PrintPdfShippingLabelWorkflow.ProcessEmails(service, response);
            }
        }

        private static string GetEmailsEmailsToPrintWithZInSubject()
        {
            var query = new List<string>();
            query.Add("is:unread");
            query.Add("subject:\"z\"");
            string q = string.Join(" ", query);
            return q;
        }

        private static async Task<GmailService> GetNewTokenWithUserInput()
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

            if (credential.Token.IsExpired(credential.Flow.Clock))
            {
                log.Write("credential expired, delete token.json and run again to get new token (requires user input).");
                //https://stackoverflow.com/questions/71777420/i-want-to-use-google-api-refresh-tokens-forever
                //If your app is in testing set it to production and your refresh token will stop expiring.
                //always refresh the auth

            }

            if (true)
            {
                await credential.RefreshTokenAsync(CancellationToken.None);
            }


            // Create the Gmail API service
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }

        public static void MarkMessageAsRead(GmailService service, Message message)
        {
            // Modify the message label to mark it as read
            var mods = new ModifyMessageRequest();
            mods.RemoveLabelIds = new List<string>() { "UNREAD" };
            service.Users.Messages.Modify(mods, "me", message.Id).Execute();
        }
    }
}
