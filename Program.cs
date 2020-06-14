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

namespace MailMerge
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            DataTable table;
            DocToPDFConverter converter;
            PdfDocument pdfDocument;
            Stream docStream = null;
            WordDocument template;
            string folder;

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            table = new DataTable();
            loadData(ref table);

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

            foreach (DataRow dataRow in table.Rows)
            {
                WordDocument document = template.Clone();

                document.MailMerge.Execute(dataRow);

                pdfDocument = converter.ConvertToPDF(document);
                pdfDocument.Save(folder + "/Diploma_Logiscool_" + dataRow.ItemArray[0].ToString() + ".pdf");
                pdfDocument.Close(true);

                document.Dispose();
            }

            MessageBox.Show("Done!");
        }

        private static void loadData(ref DataTable elements)
        {
            ExcelWorksheet worksheet;
            ExcelFile loadedFile = null;

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
                    if (String.Equals(cell.Value.ToString(), "NUME".ToString()))
                        continue;
                    elements.Rows.Add(cell.Value.ToString());
                }
            }
        }

        private static string selectFolder()
        {
            CommonOpenFileDialog dialog;
            
            dialog = new CommonOpenFileDialog();

            dialog.InitialDirectory = "C:\\Users";
            dialog.IsFolderPicker = true;

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