using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
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

        // Holds a file that failed to be read, with the reason
        private class FileErrorInfo
        {
            public string FilePath { get; set; }
            public string Message { get; set; }
        }

        // Result of a background scan
        private class ScanResult
        {
            public List<FilePageInfo> Files { get; } = new List<FilePageInfo>();
            public List<FileErrorInfo> Errors { get; } = new List<FileErrorInfo>();
            public long TotalPages { get; set; }
        }

        // Progress update reported from the background scan thread
        private class ScanProgress
        {
            public int Processed { get; set; }
            public long Pages { get; set; }
        }

        private readonly List<FilePageInfo> scannedFiles = new List<FilePageInfo>();
        private readonly List<FileErrorInfo> failedFiles = new List<FileErrorInfo>();

        // Total pages from the last completed scan; used by the cost calculator
        private long lastTotalPages = 0;

        // Non-null while a scan is running; used to cancel it from btnCancel_Click
        private CancellationTokenSource scanCts;

        private const string HelpText =
@"1. Bifează tipurile de fișiere pe care vrei să le numeri (.pdf, .docx, .doc, .docm, .dotx, .dot).

2. Apasă ""Selectează locația"" și alege folderul care conține fișierele (se caută automat și în subfoldere).

3. Așteaptă finalizarea scanării. În timpul scanării vezi progresul live (fișiere procesate + total pagini de până acum). Poți apăsa ""Anulează scanarea"" oricând.

4. La final vezi totalul de pagini. Apăsând ""Lista fișierelor scanate"" vezi fiecare fișier cu numărul lui de pagini — poți sorta lista dând click pe antetul unei coloane (Pagini, Fișier sau Status).

5. Din aceeași fereastră poți apăsa ""Export CSV"" ca să salvezi rezultatul într-un fișier pe care îl poți deschide ulterior în Excel.

6. Completează ""Preț per pagină"" (lângă totalul de pagini) ca aplicația să calculeze automat totalul de plată.

Observații:
- Numărarea paginilor din fișiere Word necesită Microsoft Word instalat pe acest calculator.
- Fișierele care nu pot fi citite (corupte, protejate cu parolă etc.) apar în lista de fișiere cu mesajul de eroare exact, nu sunt ignorate silențios.";

        public Form1()
        {
            InitializeComponent();
            // Ensure button state matches initial checkbox state
            UpdateSelectButtonState();

            // Files list button disabled until a scan with results is completed
            btnFilesList.Enabled = false;

            UpdateTotalCost();

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
            // (and no scan is currently in progress)
            btnSelectFolder.Enabled =
                scanCts == null &&
                (chkPdfFiles.Checked
                || chkWordFiles.Checked
                || checkBox1.Checked
                || checkBox2.Checked
                || checkBox3.Checked
                || checkBox4.Checked);
        }

        // Enables/disables the controls that shouldn't be touched while a scan is running
        private void SetScanningState(bool isScanning)
        {
            chkPdfFiles.Enabled = !isScanning;
            chkWordFiles.Enabled = !isScanning;
            checkBox1.Enabled = !isScanning;
            checkBox2.Enabled = !isScanning;
            checkBox3.Enabled = !isScanning;
            checkBox4.Enabled = !isScanning;

            btnCancel.Visible = isScanning;
            btnCancel.Enabled = isScanning;

            if (isScanning)
            {
                btnFilesList.Enabled = false;
            }

            UpdateSelectButtonState();
        }

        private async void btnSelectFolder_Click(object sender, EventArgs e)
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

            if (fbd.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            string rootFolderPath = fbd.SelectedPath;
            txtFolderPath.Text = rootFolderPath; // Afișează calea selectată

            // Build the extension list for Word-family files
            var wordExtensions = new List<string>();
            if (chkWordFiles.Checked) wordExtensions.Add(".docx");
            if (checkBox1.Checked) wordExtensions.Add(".doc");
            if (checkBox2.Checked) wordExtensions.Add(".docm");
            if (checkBox3.Checked) wordExtensions.Add(".dotx");
            if (checkBox4.Checked) wordExtensions.Add(".dot");
            bool includePdf = chkPdfFiles.Checked;

            string[] pdfFiles;
            List<string> wordFiles;

            try
            {
                // Single enumeration of the folder tree (no re-scanning later for progress vs. processing)
                pdfFiles = includePdf
                    ? Directory.GetFiles(rootFolderPath, "*.pdf", SearchOption.AllDirectories)
                    : Array.Empty<string>();

                wordFiles = new List<string>();
                foreach (var ext in wordExtensions.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    wordFiles.AddRange(Directory.GetFiles(rootFolderPath, "*" + ext, SearchOption.AllDirectories));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"A apărut o eroare la citirea folderului: {ex.Message}", "Eroare Majoră");
                return;
            }

            int totalCount = pdfFiles.Length + wordFiles.Count;

            // Initialize progress bar
            progressBar1.Minimum = 0;
            progressBar1.Value = 0;
            progressBar1.Maximum = Math.Max(1, totalCount); // avoid Maximum 0
            progressBar1.Invalidate();

            scannedFiles.Clear();
            failedFiles.Clear();
            lastTotalPages = 0;
            lblTotalPages.Text = "Total pagini: 0";
            UpdateTotalCost();

            scanCts = new CancellationTokenSource();
            SetScanningState(true);

            var progress = new Progress<ScanProgress>(p =>
            {
                if (progressBar1.Value < progressBar1.Maximum)
                {
                    progressBar1.Value = Math.Min(p.Processed, progressBar1.Maximum);
                    progressBar1.Invalidate(); // force redraw so custom text updates
                }

                lblTotalPages.Text = $"Total pagini: {p.Pages:N0}  ({p.Processed}/{totalCount} fișiere)";
            });

            try
            {
                ScanResult result = await RunOnStaThreadAsync(
                    () => ScanFiles(pdfFiles, wordFiles, progress, scanCts.Token),
                    scanCts.Token);

                scannedFiles.AddRange(result.Files);
                failedFiles.AddRange(result.Errors);
                lastTotalPages = result.TotalPages;

                lblTotalPages.Text = $"Total pagini: {result.TotalPages:N0}";
                UpdateTotalCost();
                btnFilesList.Enabled = scannedFiles.Count > 0 || failedFiles.Count > 0;

                progressBar1.Value = progressBar1.Maximum;
                progressBar1.Invalidate();

                string message = $"Calcul finalizat. Total pagini: {result.TotalPages:N0}.";
                if (failedFiles.Count > 0)
                {
                    message += $"\n{failedFiles.Count} fișier(e) nu au putut fi citite — vezi \"{btnFilesList.Text}\" pentru detalii.";
                }
                MessageBox.Show(message, "Succes");
            }
            catch (OperationCanceledException)
            {
                lastTotalPages = 0;
                lblTotalPages.Text = "Total pagini: 0";
                UpdateTotalCost();
                MessageBox.Show("Scanare anulată.", "Anulat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"A apărut o eroare: {ex.Message}", "Eroare Majoră");
            }
            finally
            {
                scanCts?.Dispose();
                scanCts = null;
                SetScanningState(false);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            btnCancel.Enabled = false; // avoid multiple clicks while the scan winds down
            scanCts?.Cancel();
        }

        private void numPricePerPage_ValueChanged(object sender, EventArgs e)
        {
            UpdateTotalCost();
        }

        private void UpdateTotalCost()
        {
            decimal cost = lastTotalPages * numPricePerPage.Value;
            lblTotalCost.Text = $"Total de plată: {cost:N2} lei";
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            ShowTextDialog("Ajutor - Cum se folosește aplicația", HelpText);
        }

        // Runs the given function on a dedicated STA thread and returns its result as a Task.
        // Word Automation (COM interop) is happiest when driven from an STA thread rather than
        // the default (MTA) thread pool, so we spin up our own thread instead of using Task.Run.
        private static Task<T> RunOnStaThreadAsync<T>(Func<T> work, CancellationToken cancellationToken)
        {
            var tcs = new TaskCompletionSource<T>();

            var thread = new Thread(() =>
            {
                try
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    T result = work();
                    tcs.TrySetResult(result);
                }
                catch (OperationCanceledException)
                {
                    tcs.TrySetCanceled(cancellationToken);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            });
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            return tcs.Task;
        }

        // Runs entirely on the background STA thread — must not touch any UI controls directly.
        private ScanResult ScanFiles(string[] pdfFiles, List<string> wordFiles, IProgress<ScanProgress> progress, CancellationToken token)
        {
            var files = new ConcurrentBag<FilePageInfo>();
            var errors = new ConcurrentBag<FileErrorInfo>();
            long totalPages = 0;
            int processed = 0;

            // PDFs have no shared state between files, so they can be read in parallel across cores.
            if (pdfFiles.Length > 0)
            {
                var options = new ParallelOptions
                {
                    CancellationToken = token,
                    MaxDegreeOfParallelism = Math.Max(1, Environment.ProcessorCount)
                };

                Parallel.ForEach(pdfFiles, options, file =>
                {
                    try
                    {
                        using (PdfDocument document = PdfReader.Open(file, PdfDocumentOpenMode.ReadOnly))
                        {
                            Interlocked.Add(ref totalPages, document.PageCount);
                            files.Add(new FilePageInfo { FilePath = file, PageCount = document.PageCount });
                        }
                    }
                    catch (Exception ex)
                    {
                        errors.Add(new FileErrorInfo { FilePath = file, Message = ex.Message });
                    }
                    finally
                    {
                        int done = Interlocked.Increment(ref processed);
                        progress.Report(new ScanProgress { Processed = done, Pages = Interlocked.Read(ref totalPages) });
                    }
                });
            }

            // Word Automation only supports one COM instance driving it sequentially.
            if (wordFiles.Count > 0)
            {
                token.ThrowIfCancellationRequested();

                Word.Application wordApp = null;
                try
                {
                    wordApp = new Word.Application();
                    wordApp.Visible = false;

                    object missing = Type.Missing;
                    foreach (var file in wordFiles)
                    {
                        token.ThrowIfCancellationRequested();

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
                            Interlocked.Add(ref totalPages, pages);
                            files.Add(new FilePageInfo { FilePath = file, PageCount = pages });
                        }
                        catch (Exception ex)
                        {
                            errors.Add(new FileErrorInfo { FilePath = file, Message = ex.Message });
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

                            int done = Interlocked.Increment(ref processed);
                            progress.Report(new ScanProgress { Processed = done, Pages = Interlocked.Read(ref totalPages) });
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

            var result = new ScanResult { TotalPages = Interlocked.Read(ref totalPages) };
            result.Files.AddRange(files);
            result.Errors.AddRange(errors);
            return result;
        }

        private void btnFilesList_Click(object sender, EventArgs e)
        {
            if (scannedFiles.Count == 0 && failedFiles.Count == 0)
            {
                MessageBox.Show("Nu există fișiere scanate pentru afișare.", "Lista fișiere", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DataTable table = BuildResultsTable();

            using (Form listForm = new Form())
            {
                listForm.Text = "Fișiere scanate";
                listForm.StartPosition = FormStartPosition.CenterParent;
                listForm.Size = new System.Drawing.Size(900, 600);

                listForm.ShowIcon = true;
                if (this.Icon != null)
                    listForm.Icon = this.Icon;

                var grid = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    AllowUserToResizeRows = false,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    Font = this.Font,
                    DataSource = table
                };
                grid.DataBindingComplete += (s, ev) =>
                {
                    if (grid.Columns["Pagini"] != null)
                    {
                        grid.Columns["Pagini"].FillWeight = 15;
                        grid.Columns["Pagini"].DefaultCellStyle.Format = "N0";
                    }
                    if (grid.Columns["Fișier"] != null) grid.Columns["Fișier"].FillWeight = 65;
                    if (grid.Columns["Status"] != null) grid.Columns["Status"].FillWeight = 20;
                };

                var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 50 };
                var btnExport = new Button
                {
                    Text = "Export CSV",
                    Font = this.Font,
                    Size = new System.Drawing.Size(150, 36),
                    Location = new System.Drawing.Point(10, 7)
                };
                btnExport.Click += (s, ev) => ExportResultsToCsv(table);
                bottomPanel.Controls.Add(btnExport);

                listForm.Controls.Add(grid);
                listForm.Controls.Add(bottomPanel);

                listForm.ShowDialog(this);
            }
        }

        private DataTable BuildResultsTable()
        {
            var table = new DataTable();
            table.Columns.Add("Pagini", typeof(long));
            table.Columns.Add("Fișier", typeof(string));
            table.Columns.Add("Status", typeof(string));

            foreach (var f in scannedFiles)
            {
                table.Rows.Add(f.PageCount, f.FilePath, "OK");
            }
            foreach (var f in failedFiles)
            {
                table.Rows.Add(DBNull.Value, f.FilePath, "Eroare: " + f.Message);
            }

            return table;
        }

        private void ExportResultsToCsv(DataTable table)
        {
            using (var sfd = new SaveFileDialog
            {
                Filter = "Fișiere CSV (*.csv)|*.csv",
                FileName = "fisiere_scanate.csv"
            })
            {
                if (sfd.ShowDialog(this) != DialogResult.OK)
                    return;

                try
                {
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine("Pagini,Fisier,Status");
                    foreach (DataRow row in table.Rows)
                    {
                        string pages = row["Pagini"] == DBNull.Value ? "" : row["Pagini"].ToString();
                        sb.AppendLine(string.Join(",",
                            CsvEscape(pages),
                            CsvEscape(row["Fișier"].ToString()),
                            CsvEscape(row["Status"].ToString())));
                    }

                    File.WriteAllText(sfd.FileName, sb.ToString(), System.Text.Encoding.UTF8);
                    MessageBox.Show("Export finalizat cu succes.", "Export CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Exportul a eșuat: {ex.Message}", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private static string CsvEscape(string value)
        {
            if (value == null) return "";
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r"))
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }
            return value;
        }

        private void ShowTextDialog(string title, string text)
        {
            using (Form dlg = new Form())
            {
                dlg.Text = title;
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.Size = new System.Drawing.Size(650, 450);
                dlg.ShowIcon = true;
                if (this.Icon != null) dlg.Icon = this.Icon;

                TextBox tb = new TextBox
                {
                    Multiline = true,
                    ReadOnly = true,
                    Dock = DockStyle.Fill,
                    ScrollBars = ScrollBars.Vertical,
                    Font = this.Font,
                    Text = text
                };
                dlg.Controls.Add(tb);
                dlg.ShowDialog(this);
            }
        }
    }
}
