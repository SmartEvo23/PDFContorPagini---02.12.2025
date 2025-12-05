using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SixLabors.ImageSharp;

namespace PDFContorPagini
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // You can add any initialization code here if needed.
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            // Folosește dialogul standard Windows pentru selectarea unui folder
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Selectați directorul rădăcină care conține fișierele PDF.";

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                string rootFolderPath = fbd.SelectedPath;
                txtFolderPath.Text = rootFolderPath; // Afișează calea selectată

                try
                {
                    long totalPageCount = CountTotalPdfPages(rootFolderPath);
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
            // Folosim SearchOption.AllDirectories
            string[] pdfFiles = Directory.GetFiles(rootDirectoryPath, "*.pdf", SearchOption.AllDirectories);

            foreach (string file in pdfFiles)
            {
                try
                {
                    // Deschide fișierul PDF folosind PDFsharp
                    // Modul PdfDocumentOpenMode.ReadOnly este eficient
                    using (PdfDocument document = PdfReader.Open(file, PdfDocumentOpenMode.ReadOnly))
                    {
                        totalPages += document.PageCount;
                    }
                }
                catch (Exception ex)
                {
                    // Ignoră fișierele corupte sau protejate cu parolă și raportează eroarea
                    Console.WriteLine($"Eroare la citirea fișierului {Path.GetFileName(file)}: {ex.Message}");
                    // Puteți adăuga o logare mai detaliată aici dacă doriți
                }
            }

            return totalPages;
        }
    }
}
