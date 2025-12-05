using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace FisiereContorPagini
{
    public partial class Form1 : Form
    {
        // Holds the scanned files and their page counts
        private class FilePageInfo
        {
            public string FilePath { get; set; }
            public long PageCount { get; set; }
        }

        private readonly List<FilePageInfo> scannedFiles = new List<FilePageInfo>();

        public Form1()
        {
            InitializeComponent();
            // Ensure button state matches initial checkbox state
            UpdateSelectButtonState();

            // Files list button disabled until a scan with results is completed
            btnFilesList.Enabled = false;

            // Set window icon (use the application's associated icon as a fallback)
            try
            {
                this.Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch
            {
                // ignore if icon extraction fails
            }
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
                    // Clear any previous scan results
                    scannedFiles.Clear();

                    // Build list of files to process so we can drive the progress bar
                    var allFiles = new List<string>();

                    if (chkPdfFiles.Checked)
                    {
                        allFiles.AddRange(Directory.GetFiles(rootFolderPath, "*.pdf", SearchOption.AllDirectories));
                    }

                    var wordExtensions = new List<string>();
                    if (chkWordFiles.Checked) wordExtensions.Add(".docx");
                    if (checkBox1.Checked) wordExtensions.Add(".doc");
                    if (checkBox2.Checked) wordExtensions.Add(".docm");
                    if (checkBox3.Checked) wordExtensions.Add(".dotx");
                    if (checkBox4.Checked) wordExtensions.Add(".dot");

                    if (wordExtensions.Count > 0)
                    {
                        foreach (var ext in wordExtensions.Distinct(StringComparer.OrdinalIgnoreCase))
                        {
                            string pattern = "*" + ext;
                            allFiles.AddRange(Directory.GetFiles(rootFolderPath, pattern, SearchOption.AllDirectories));
                        }
                    }

                    // Initialize progress bar
                    progressBar1.Minimum = 0;
                    progressBar1.Value = 0;
                    progressBar1.Step = 1;
                    progressBar1.Maximum = Math.Max(1, allFiles.Count); // avoid Maximum 0

                    long totalPageCount = 0;

                    // Process PDF files (synchronously on UI thread; update progress and call DoEvents so UI updates)
                    if (chkPdfFiles.Checked)
                    {
                        string[] pdfFiles = Directory.GetFiles(rootFolderPath, "*.pdf", SearchOption.AllDirectories);
                        foreach (string file in pdfFiles)
                        {
                            try
                            {
                                using (PdfDocument document = PdfReader.Open(file, PdfDocumentOpenMode.ReadOnly))
                                {
                                    totalPageCount += document.PageCount;
                                    scannedFiles.Add(new FilePageInfo { FilePath = file, PageCount = document.PageCount });
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Eroare la citirea fișierului {Path.GetFileName(file)}: {ex.Message}");
                            }
                            finally
                            {
                                if (progressBar1.Value < progressBar1.Maximum)
                                {
                                    progressBar1.PerformStep();
                                    progressBar1.Invalidate();   // force redraw so custom text updates
                                    Application.DoEvents();
                                }
                            }
                        }
                    }

                    // Process Word files. We open Word once and iterate through files.
                    if (wordExtensions.Count > 0)
                    {
                        // Collect word files again (to preserve previous behavior)
                        var files = new List<string>();
                        foreach (var ext in wordExtensions.Distinct(StringComparer.OrdinalIgnoreCase))
                        {
                            string pattern = "*" + ext;
                            files.AddRange(Directory.GetFiles(rootFolderPath, pattern, SearchOption.AllDirectories));
                        }

                        if (files.Count > 0)
                        {
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
                                        totalPageCount += pages;
                                        scannedFiles.Add(new FilePageInfo { FilePath = file, PageCount = pages });
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

                                        if (progressBar1.Value < progressBar1.Maximum)
                                        {
                                            progressBar1.PerformStep();
                                            progressBar1.Invalidate();   // force redraw so custom text updates
                                            Application.DoEvents();
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
                        }
                    }

                    lblTotalPages.Text = $"Total pagini: {totalPageCount}";

                    // Enable the files-list button only if we have scanned files
                    btnFilesList.Enabled = scannedFiles.Count > 0;

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

        private void btnFilesList_Click(object sender, EventArgs e)
        {
            if (scannedFiles == null || scannedFiles.Count == 0)
            {
                MessageBox.Show("Nu există fișiere scanate pentru afișare.", "Lista fișiere", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Create a simple modal window showing the file list with page counts
            using (Form listForm = new Form())
            {
                listForm.Text = "Fișiere scanate";
                listForm.StartPosition = FormStartPosition.CenterParent;
                listForm.Size = new System.Drawing.Size(800, 600);

                // show the same icon as the main window (if present)
                listForm.ShowIcon = true;
                if (this.Icon != null)
                    listForm.Icon = this.Icon;

                TextBox tb = new TextBox
                {
                    Multiline = true,
                    ReadOnly = true,
                    Dock = DockStyle.Fill,
                    ScrollBars = ScrollBars.Both,
                    Font = this.Font,
                    WordWrap = false
                };

                var sb = new System.Text.StringBuilder();
                sb.AppendLine("Pages\tFile");
                sb.AppendLine("-----\t----");
                foreach (var f in scannedFiles)
                {
                    sb.AppendLine($"{f.PageCount}\t{f.FilePath}");
                }

                tb.Text = sb.ToString();
                listForm.Controls.Add(tb);

                listForm.ShowDialog(this);
            }
        }
    }
}
