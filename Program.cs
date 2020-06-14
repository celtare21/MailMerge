using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using GemBox.Spreadsheet;
using System;
using System.IO;
using System.Windows.Forms;
using System.Data;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Collections.Generic;
using System.Net.Mail;
using System.Net;
using System.Net.Mime;

namespace MailMerge
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            SmtpClient smtpClient;
            DataTable table;
            DocToPDFConverter converter;
            PdfDocument pdfDocument;
            Stream docStream = null;
            WordDocument template;
            List<string> emails;
            string folder, email, pass;
            int i = 0;

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            Console.WriteLine("Email:");
            do
            {
                email = Console.ReadLine();
            } while (email == "");

            Console.WriteLine("Password:");
            do
            {
                pass = Console.ReadLine();
            } while (pass == "");

            table = new DataTable();
            emails = new List<string>();
            loadData(ref table, ref emails);

            MessageBox.Show("Load Template Doc:");
            try
            {
                docStream = File.OpenRead(loadFile(true));
            }
            catch (ArgumentNullException)
            {
                MessageBox.Show("No file loaded! Exiting!");
                Environment.Exit(0);
            }
            template = new WordDocument(docStream, FormatType.Docx);
            docStream.Dispose();

            converter = new DocToPDFConverter();
            converter.Settings.EnableFastRendering = true;
            converter.Settings.EmbedFonts = true;
            converter.Settings.OptimizeIdenticalImages = true;

            MessageBox.Show("Choose output folder:");
            folder = selectFolder();
            if (folder == null)
            {
                MessageBox.Show("No folder selected! Exiting!");
                Environment.Exit(0);
            }

            smtpClient = new SmtpClient("smtp-mail.outlook.com")
            {
                Port = 587,
                Credentials = new NetworkCredential(email, pass),
                EnableSsl = true,
            };

            foreach (DataRow dataRow in table.Rows)
            {
                MailMessage mailMessage;
                Attachment attachment;
                WordDocument document;
                string path;

                document = template.Clone();
                document.MailMerge.Execute(dataRow);

                path = folder + "/Diploma_Logiscool_" + dataRow.ItemArray[0].ToString() + ".pdf";

                pdfDocument = converter.ConvertToPDF(document);
                pdfDocument.Save(path);
                pdfDocument.Close(true);

                mailMessage = new MailMessage
                {
                    From = new MailAddress(email),
                    Subject = "Test",
                    Body = "Acesta este un test",
                };

                attachment = new Attachment(path, MediaTypeNames.Application.Pdf);
                mailMessage.Attachments.Add(attachment);
                mailMessage.To.Add(emails[i++]);

                smtpClient.Send(mailMessage);

                document.Dispose();
            }

            MessageBox.Show("Done!");
        }

        private static void loadData(ref DataTable elements, ref List<string> emails)
        {
            ExcelWorksheet worksheet;
            ExcelFile loadedFile = null;
            bool wrow = false;

            MessageBox.Show("Choose the database:");
            try
            {
                loadedFile = ExcelFile.Load(loadFile(false));
            } 
            catch (ArgumentNullException)
            {
                MessageBox.Show("No file loaded! Exiting!");
                Environment.Exit(0);
            }

            elements.Columns.Add("NUME", typeof(string));
            worksheet = loadedFile.Worksheets[0];

            foreach (ExcelRow row in worksheet.Rows)
            {
                foreach (ExcelCell cell in row.AllocatedCells)
                {
                    if (String.Equals(cell.Value.ToString().ToUpper(), "NUME".ToString()) || String.Equals(cell.Value.ToString().ToUpper(), "EMAIL".ToString()))
                        continue;

                    if (!wrow)
                        elements.Rows.Add(cell.Value.ToString().ToUpper());
                    else
                        emails.Add(cell.Value.ToString());

                    wrow = !wrow;
                }
            }
        }

        private static string selectFolder()
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog
            {
                InitialDirectory = "C:\\Users",
                IsFolderPicker = true
            };

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                return dialog.FileName;

            return null;
        }

        static private string loadFile(bool doc)
        {
            OpenFileDialog openFileDialog;

            if (doc)
            {
                openFileDialog = new OpenFileDialog
                {
                    Title = "Browse Doc Files",

                    CheckFileExists = true,
                    CheckPathExists = true,
                    RestoreDirectory = true,

                    DefaultExt = "doc",
                    Filter = "Word File (.docx ,.doc)|*.docx;*.doc",
                    FilterIndex = 2,

                    ShowReadOnly = true
                };
            }
            else
            {
                openFileDialog = new OpenFileDialog
                {
                    Title = "Browse Excel Files",

                    CheckFileExists = true,
                    CheckPathExists = true,
                    RestoreDirectory = true,

                    DefaultExt = "xlsx",
                    Filter = "xlsx files (*.xlsx)|*.xlsx",
                    FilterIndex = 2,

                    ShowReadOnly = true
                };
            }

            if (openFileDialog.ShowDialog() == DialogResult.OK)
                return openFileDialog.FileName.ToString();

            return null;
        }
    }
}