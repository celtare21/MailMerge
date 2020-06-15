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
using System.ComponentModel;

namespace MailMerge
{
    class Program
    {
        private static bool mailSending = false;

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
            string path = null, oldpath;
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
            smtpClient.SendCompleted += new SendCompletedEventHandler(SendCompletedCallback);

            Console.Clear();

            foreach (DataRow dataRow in table.Rows)
            {
                MailMessage mailMessage = null;
                Attachment attachment;
                WordDocument document;

                oldpath = path;
                path = folder + "/Diploma_Logiscool_" + dataRow.ItemArray[0] + ".pdf";

                if (path == oldpath)
                    continue;

                document = template.Clone();
                document.MailMerge.Execute(dataRow);

                pdfDocument = converter.ConvertToPDF(document);
                pdfDocument.Save(path);
                pdfDocument.Close(true);

                try
                {
                    mailMessage = new MailMessage
                    {
                        From = new MailAddress(email),
                        Subject = "Test",
                        Body = "Acesta este un test",
                    };
                }
                catch (FormatException)
                {
                    MessageBox.Show("Invalid email address!");
                    Environment.Exit(0);
                }

                Console.WriteLine(dataRow.ItemArray[0] + " " + emails[i]);
                attachment = new Attachment(path, MediaTypeNames.Application.Pdf);
                mailMessage.Attachments.Add(attachment);
                mailMessage.To.Add(emails[i++]);

                while (mailSending) ;

                mailSending = true;

                try
                {
                    smtpClient.SendAsync(mailMessage, null);
                }
                catch (SmtpException)
                {
                    MessageBox.Show("Couldn't connect to the email address!");
                    Environment.Exit(0);
                }

                document.Dispose();
            }

            while (mailSending) ;

            MessageBox.Show("Done!");
        }


        private static void loadData(ref DataTable elements, ref List<string> emails)
        {
            ExcelWorksheet worksheet;
            ExcelFile loadedFile = null;
            string scell;
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
                    scell = cell.Value.ToString();

                    if (String.Equals(scell.ToUpper(), "NUME") || String.Equals(scell.ToUpper(), "EMAIL"))
                        continue;

                    if (!wrow)
                        elements.Rows.Add(scell.ToUpper());
                    else
                        emails.Add(scell);

                    wrow = !wrow;
                }
            }
        }

        private static string selectFolder()
        {
            CommonOpenFileDialog openFolderDialog = new CommonOpenFileDialog
            {
                InitialDirectory = "C:\\Users",
                IsFolderPicker = true
            };

            if (openFolderDialog.ShowDialog() == CommonFileDialogResult.Ok)
                return openFolderDialog.FileName;

            return null;
        }

        static private string loadFile(bool doc)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.CheckFileExists = true;
            openFileDialog.CheckPathExists = true;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 2;
            openFileDialog.ShowReadOnly = true;

            if (doc)
            {
                openFileDialog.Title = "Browse Doc Files";
                openFileDialog.DefaultExt = "doc";
                openFileDialog.Filter = "Word File (.docx ,.doc)|*.docx;*.doc";

            }
            else
            {
                openFileDialog.Title = "Browse Excel Files";
                openFileDialog.DefaultExt = "xlsx";
                openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx";
            }

            if (openFileDialog.ShowDialog() == DialogResult.OK)
                return openFileDialog.FileName.ToString();

            return null;
        }

        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            mailSending = false;
        }
    }
}