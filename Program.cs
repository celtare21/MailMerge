﻿using Syncfusion.DocIO;
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
using System.Security;

namespace MailMerge
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            SmtpClient smtpClient;
            DataTable names;
            DocToPDFConverter converter;
            PdfDocument pdfDocument;
            Stream docStream = null;
            StreamReader reader;
            WordDocument template, body;
            List<string> emails;
            string folder, email, mailBody;
            SecureString pass;
            string path = null, oldpath;
            int i = 0;

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("MjcyNDY4QDMxMzgyZTMxMmUzMGZXNjBmNHF4TU51RndrbmJ4MjcyMnJMV3ZlYlk1ekZVc1lWcGlqaDFhQm89;MjcyNDY5QDMxMzgyZTMxMmUzMEFGWXNRajhzQld0UERPeVpObzhCTTRyUm15YmVHUTZ3Ny9MVFJQT05WWHM9;MjcyNDcwQDMxMzgyZTMxMmUzMExOeVBJVmFUTlhXaG1FM1lkVjZpRUhJSkV0bExGTkxtc3NyZC9WV2dHMUU9;MjcyNDcxQDMxMzgyZTMxMmUzMFBjZ0dNVHBZYnhpSUFrQzZSd3BNVm82WHRUc1c1VjhlSEV2ak0xeUpDakU9;MjcyNDcyQDMxMzgyZTMxMmUzMEZTYWNPR1ZPejdvQ1JzemRvSFBvZThjZ1lzNUtZYlVCelQ3TjFCOU9CNFk9;MjcyNDczQDMxMzgyZTMxMmUzMEV3cXlVVmsxcGN3WHc3K2YwZDdqUjc1MnFMbDFsRUF0V1pBWDhUaENRenM9;MjcyNDc0QDMxMzgyZTMxMmUzMFE3YjFIazl3WTM5WGx2T2QrVVUvL1B0aGxkT3k1aDMvV2lzZXRXQ3NMeG89;MjcyNDc1QDMxMzgyZTMxMmUzMGFwLytWOWNxckpoYW1mK1pPQnl4N1RBbmM5UFJoU2dzM1dNQVFlb3ZlaTQ9;MjcyNDc2QDMxMzgyZTMxMmUzMFd2WWprcm0xOElBTlg0VGMyT0ViZVBlMWk1Uzl4M2tDNkVkTGJ1T1AxMTQ9;NT8mJyc2IWhia31ifWN9ZmVoYmF8YGJ8ampqanNiYmlmamlmanMDHmgwNj8nMiE2YWITND4yOj99MDw+;MjcyNDc3QDMxMzgyZTMxMmUzMENUVHFZTHU3NENEenRCUkRaSW1KNVNMd3ZZOE9qMlpRZWs5WVZmazdkYjQ9");

            Console.WriteLine("Email:");
            do
            {
                email = Console.ReadLine();
            } while (email == "");

            Console.WriteLine("Password:");
            do
            {
                pass = GetPassword();
            } while (pass.Length == 0);

            names = new DataTable();
            emails = new List<string>();
            loadData(ref names, ref emails);

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

            MessageBox.Show("Load Email Body:");
            try
            {
                docStream = File.OpenRead(loadFile(true));
            }
            catch (ArgumentNullException)
            {
                MessageBox.Show("No file loaded! Exiting!");
                Environment.Exit(0);
            }
            body = new WordDocument(docStream, FormatType.Doc);
            docStream.Dispose();

            docStream = new MemoryStream();

            body.SaveOptions.HtmlExportOmitXmlDeclaration = true;
            body.Save(docStream, FormatType.Html);
            body.Dispose();

            docStream.Position = 0;
            reader = new StreamReader(docStream);
            mailBody = reader.ReadToEnd();
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

            Console.Clear();

            foreach (DataRow dataRow in names.Rows)
            {
                MailMessage mailMessage = null;
                Attachment attachment;
                WordDocument document;

                oldpath = path;
                path = folder + "\\Diploma_Logiscool_" + dataRow.ItemArray[0] + ".pdf";

                if (path != oldpath)
                {
                    document = template.Clone();
                    document.MailMerge.Execute(dataRow);

                    pdfDocument = converter.ConvertToPDF(document);
                    pdfDocument.Save(path);
                    pdfDocument.Close(true);

                    document.Dispose();
                }

                try
                {
                    mailMessage = new MailMessage
                    {
                        From = new MailAddress(email),
                        Subject = "Diploma Logiscool",
                        Body = mailBody,
                        IsBodyHtml = true,
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

                try
                {
                    smtpClient.Send(mailMessage);
                }
                catch (SmtpException)
                {
                    MessageBox.Show("Couldn't connect to the email address!");
                    Environment.Exit(0);
                }

                Console.WriteLine("...Email sent!");
            }

            MessageBox.Show("Done!");
        }

        private static SecureString GetPassword()
        {
            SecureString pwd = new SecureString();

            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else if (i.KeyChar != '\u0000')
                {
                    pwd.AppendChar(i.KeyChar);
                    Console.Write("*");
                }
            }

            return pwd;
        }

        private static void loadData(ref DataTable names, ref List<string> emails)
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
            catch (Exception ex)
            {
                if (ex is ArgumentNullException)
                    MessageBox.Show("No file loaded! Exiting!");
                else if (ex is FreeLimitReachedException)
                    MessageBox.Show("More than 150 rows loaded! Exiting!");
                Environment.Exit(0);
            }

            names.Columns.Add("NUME", typeof(string));
            worksheet = loadedFile.Worksheets[0];

            foreach (ExcelRow row in worksheet.Rows)
            {
                foreach (ExcelCell cell in row.AllocatedCells)
                {
                    if (cell.ValueType == CellValueType.Null)
                        continue;

                    scell = cell.Value.ToString();

                    if (String.Equals(scell.ToUpper(), "NUME") || String.Equals(scell.ToUpper(), "EMAIL"))
                        continue;

                    if (!wrow)
                        names.Rows.Add(scell.ToUpper());
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
    }
}