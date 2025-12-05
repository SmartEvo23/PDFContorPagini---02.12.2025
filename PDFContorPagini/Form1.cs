using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace PDFContorPagini
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            // Ensure button state matches initial checkbox state
            UpdateSelectButtonState();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // You can add any initialization code here if needed.
        }

        private void FileTypeCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            UpdateSelectButtonState();
        }

        private void UpdateSelectButtonState()
        {
            // Button enabled only when at least one file-type checkbox is checked
            btnSelectFolder.Enabled =
                chkPdfFiles.Checked
                || chkWordFiles.Checked
                || checkBox1.Checked
                || checkBox2.Checked
                || checkBox3.Checked
                || checkBox4.Checked;
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            // Guard in case button clicked while no checkbox is selected
            if (!(chkPdfFiles.Checked || chkWordFiles.Checked || checkBox1.Checked || checkBox2.Checked || checkBox3.Checked || checkBox4.Checked))
            {
                MessageBox.Show("Selectați cel puțin un tip de fișier înainte de a continua.", "Alege tip fișiere", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Folosește dialogul standard Windows pentru selectarea unui folder
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Selectați directorul rădăcină care conține fișierele selectate.";

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                string rootFolderPath = fbd.SelectedPath;
                txtFolderPath.Text = rootFolderPath; // Afișează calea selectată

                try
                {
                    long totalPageCount = 0;

                    // PDF count if selected
                    if (chkPdfFiles.Checked)
                    {
                        totalPageCount += CountTotalPdfPages(rootFolderPath);
                    }

                    // Word-related types: map checkboxes to extensions
                    var wordExtensions = new List<string>();
                    if (chkWordFiles.Checked) wordExtensions.Add(".docx");
                    if (checkBox1.Checked) wordExtensions.Add(".doc");
                    if (checkBox2.Checked) wordExtensions.Add(".docm");
                    if (checkBox3.Checked) wordExtensions.Add(".dotx");
                    if (checkBox4.Checked) wordExtensions.Add(".dot");

                    if (wordExtensions.Count > 0)
                    {
                        totalPageCount += CountTotalWordPages(rootFolderPath, wordExtensions);
                    }

                    lblTotalPages.Text = $"Total pagini: {totalPageCount}";
                    MessageBox.Show($"Calcul finalizat. Total pagini: {totalPageCount}", "Succes");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"A apărut o eroare: {ex.Message}", "Eroare Majoră");
                }
            }
        }

        private long CountTotalPdfPages(string rootDirectoryPath)
        {
            long totalPages = 0;

            // Caută toate fișierele .pdf recursiv (în subfoldere)
            string[] pdfFiles = Directory.GetFiles(rootDirectoryPath, "*.pdf", SearchOption.AllDirectories);

            foreach (string file in pdfFiles)
            {
                try
                {
                    using (PdfDocument document = PdfReader.Open(file, PdfDocumentOpenMode.ReadOnly))
                    {
                        totalPages += document.PageCount;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Eroare la citirea fișierului {Path.GetFileName(file)}: {ex.Message}");
                }
            }

            return totalPages;
        }

        private long CountTotalWordPages(string rootDirectoryPath, IEnumerable<string> extensions)
        {
            long totalPages = 0;

            // Collect files for the requested extensions
            var files = new List<string>();
            foreach (var ext in extensions.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                string pattern = "*" + ext;
                files.AddRange(Directory.GetFiles(rootDirectoryPath, pattern, SearchOption.AllDirectories));
            }

            if (files.Count == 0)
                return 0;

            Word.Application wordApp = null;
            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;

                object missing = Type.Missing;
                foreach (var file in files)
                {
                    Word.Document doc = null;
                    object fileName = file;
                    object readOnly = true;
                    object isVisible = false;

                    try
                    {
                        // Open document in read-only, invisible mode
                        doc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing,
                                                     ref missing, ref missing, ref missing, ref missing, ref missing,
                                                     ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                        // ComputeStatistics returns page count
                        int pages = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, false);
                        totalPages += pages;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Eroare la citirea fișierului Word {Path.GetFileName(file)}: {ex.Message}");
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            try
                            {
                                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                                doc.Close(ref saveChanges, ref missing, ref missing);
                            }
                            catch { }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                        }
                    }
                }
            }
            finally
            {
                if (wordApp != null)
                {
                    try
                    {
                        object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                        wordApp.Quit(ref saveChanges, Type.Missing, Type.Missing);
                    }
                    catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
            }

            return totalPages;
        }
    }
}
