using Google.Apis.Gmail.v1.Data;
using Google.Apis.Gmail.v1;
using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static PrintAutomation.Program;

namespace PrintAutomation
{
    public class PrintPdfShippingLabelWorkflow
    {
        public static void ProcessEmails(GmailService service, ListMessagesResponse response)
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

        public static void ProcessMessageParts(GmailService service, Message message, string filename)
        {


            foreach (var part in message.Payload.Parts)
            {
                if (part.Filename != null && part.Body.AttachmentId != null)
                {
                    log.Write("Pdf: " + filename);
                    // This is a regular attachment
                    var attachment = service.Users.Messages.Attachments.Get("me", message.Id, part.Body.AttachmentId).Execute();
                    var attachmentData = attachment.Data;
                    // process attachment data...
                    log.Write("saving pdf to disk...");
                    filename = SavePdfToDisk(filename, attachmentData);
                    log.Write("Pdf saved as: " + filename);
                    log.Write("printing to pdf...");
                    PrintPDF("HPAB4538 (HP DeskJet 3700 series)", "Letter", filename);
                    log.Write("marking email as read...");
                    MarkMessageAsRead(service, message);
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

        public static string SavePdfToDisk(string filename, string attachmentData)
        {
            var dir = Path.GetDirectoryName(filename);
            var ext = Path.GetExtension(filename);
            var baseName = Path.GetFileNameWithoutExtension(filename);

            var newName = filename;
            var exists = File.Exists(filename);
            var i = 1;
            while (exists)
            {
                newName = dir + "\\" + baseName + "(" + i + ")" + ext;
                exists = File.Exists(newName);
                i++;
            }

            // save newlyCalculated save
            if (!filename.Equals(newName))
            {
                filename = newName;
            }
            // Convert the attachment data from base64url to byte array
            byte[] convertedByteArray = Convert.FromBase64String(attachmentData.Replace('-', '+').Replace('_', '/'));

            // Save the PDF attachment to disk
            using (FileStream stream = new FileStream(filename, FileMode.Create, FileAccess.Write))
            {
                stream.Write(convertedByteArray, 0, convertedByteArray.Length);
            }
            return filename;
        }

        public static bool PrintPDF(string printer, string paperName, string filename)
        {
            try
            {
                // Create the printer settings for our printer
                var printerSettings = new PrinterSettings
                {
                    PrinterName = printer,
                    Copies = 1,
                };
                // Create our page settings for the paper size selected
                var pageSettings = new PageSettings(printerSettings)


                {
                    Margins = new Margins(0, 0, 0, 0),
                };
                foreach (PaperSize papersize in printerSettings.PaperSizes)
                {
                    if (papersize.PaperName == paperName)
                    {
                        pageSettings.PaperSize = papersize;
                        break;
                    }
                }


                // Now print the PDF document
                using (var document = PdfDocument.Load(filename))
                {
                    using (var printDocument = document.CreatePrintDocument())
                    {
                        printDocument.PrinterSettings = printerSettings;
                        printDocument.DefaultPageSettings = pageSettings;
                        printDocument.OriginAtMargins = true;

                        printDocument.Print();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                log.Exception(ex);
                return false;
            }
        }


    }
}
