using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

using Path = System.IO.Path;
using libraryiTextSharp = iTextSharp.text.pdf;
using libraryiTextSharpParser = iTextSharp.text.pdf.parser;
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;
using System.IO;



// This is the code for your desktop app.
// Press Ctrl+F5 (or go to Debug > Start Without Debugging) to run your app.

namespace PDF_AppTool
{
    public partial class Form1 : Form
    {
        public static string archivoPDFPrincipal = "";
        public static string rutaArchivos = Path.Combine(Directory.GetCurrentDirectory(), "Archivos");
        public string DateFolderStructure = "";
        public string rutaArchivosFecha = "";
        public string rutaArchivosFechaSalida = "";
        public string rutaArchivosFechaTemp = "";
        public string rutaArchivoConsolidado = "";

        public Form1()
        {
            InitializeComponent();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }


        static void creaPDF()
        {
            //MessageBox.Show("Thanks!");
            //New document
            PdfDocument document = new PdfDocument();

            //new page
            PdfPage page = document.AddPage();

            XGraphics gfx = XGraphics.FromPdfPage(page);

            XFont font = new XFont("Arial", 20);

            gfx.DrawString("First line testing", font, XBrushes.Black,
                new XRect(0, 0, page.Width, page.Height),
                XStringFormats.Center);

            gfx.DrawString("second line ", font, XBrushes.Violet,
                new XRect(0, 0, page.Width, page.Height),
                XStringFormats.BottomLeft);
            gfx.DrawString("third", font, XBrushes.Red, new XPoint(100, 300));

            document.Save("C:\\Users\\carlos.aucancela\\Documents\\TEStCArlos.pdf");
        }

        public async Task DividirPDFenArchivos(string filename)
        {
            await Task.Factory.StartNew(() =>
                {
                    PdfDocument inputDocument = PdfReader.Open(filename, PdfDocumentOpenMode.Import);

                    string name = Path.GetFileNameWithoutExtension(filename);
                    for (int idx = 0; idx < inputDocument.PageCount; idx += 2)
                    {
                        // Create new document
                        PdfDocument outputDocument = new PdfDocument();
                        outputDocument.Version = inputDocument.Version;
                        outputDocument.Info.Title =
                          String.Format("Page {0} of {1}", idx + 1, inputDocument.Info.Title);
                        outputDocument.Info.Creator = inputDocument.Info.Creator;

                        // Add the page and save it

                        if (idx + 1 == inputDocument.PageCount && (inputDocument.PageCount % 2) == 0)
                        {
                            outputDocument.AddPage(inputDocument.Pages[idx]);
                            outputDocument.Save(Path.Combine(rutaArchivosFechaTemp, String.Format("{0} - Page {1}.pdf", name, idx + 1)));
                        }
                        else
                        {
                            outputDocument.AddPage(inputDocument.Pages[idx]);
                            outputDocument.AddPage(inputDocument.Pages[idx + 1]);
                            outputDocument.Save(Path.Combine(rutaArchivosFechaTemp, String.Format("{0} - Page {1} and {2}.pdf", name, idx + 1, idx + 2)));
                        }
                    }
                    Thread.Sleep(10);
                }
           );
        }

        public string ExtractOrden(string pdfFileName)
        {
            string textFind = "00000";
            StringBuilder sb = new StringBuilder();
            try
            {
                using (libraryiTextSharp.PdfReader reader = new libraryiTextSharp.PdfReader(pdfFileName))
                {
                    for (int pageNo = 1; pageNo <= 1; pageNo++)
                    //reader.NumberOfPages
                    {
                        libraryiTextSharpParser.ITextExtractionStrategy strategy = new libraryiTextSharpParser.SimpleTextExtractionStrategy();
                        string text = libraryiTextSharpParser.PdfTextExtractor.GetTextFromPage(reader, pageNo, strategy);
                        text = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(text)));
                        sb.Append(text);
                        textFind = getBetween(text, "Orden:", "S.O.#");
                        if (textFind.Length > 0 && textFind.Length <= 12)
                        {
                            return textFind.TrimEnd().TrimStart() + ".pdf";
                        }
                    }
                }
            }
            catch (Exception)
            {
                return "0000-error" + ".pdf";
            }

            return "00000000" + ".pdf";

        }

        public static string getBetween(string strSource, string strStart, string strEnd)
        {
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                int Start, End;
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }

            return "";
        }
        public  IEnumerable<string> GetPdfFiles()
        {
            return Directory.GetFiles(rutaArchivosFechaTemp, "*.pdf", SearchOption.AllDirectories);
            //return Directory.GetFiles(Path.Combine("..", "..", "..", "PdfTextract.Tests", "Data", "PDFs"), "*.pdf", SearchOption.AllDirectories);

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btnLoadFile_Click(object sender, EventArgs e)
        {
            DateFolderStructure = DateTime.Now.ToString("yyyyMMdd_hhmmss");
            rutaArchivosFecha = Path.Combine(Directory.GetCurrentDirectory(), "Archivos", DateFolderStructure);
            rutaArchivosFechaSalida = Path.Combine(Directory.GetCurrentDirectory(), "Archivos", DateFolderStructure, "Salida");
            rutaArchivosFechaTemp = Path.Combine(Directory.GetCurrentDirectory(), "Archivos", DateFolderStructure, "Temp");
            rutaArchivoConsolidado = Path.Combine(Directory.GetCurrentDirectory(), "Archivos", DateFolderStructure, "Consolidado");

            ValidarDirectorio(rutaArchivos);
            ValidarDirectorio(rutaArchivosFechaTemp);
            ValidarDirectorio(rutaArchivosFecha);
            ValidarDirectorio(rutaArchivoConsolidado);
            ValidarDirectorio(rutaArchivosFechaSalida);

            LoadFile();
            btnDividir.Enabled = true;
            btnAbrirCarpeta.Enabled = false;
            btnRenombrar.Enabled = false;
            picOkPaso1.Visible = true;
            picOkPaso2.Visible = false;
            picOkPaso3.Visible = false;
        }

        private void LoadFile()
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.InitialDirectory = rutaArchivosFechaTemp;
            openFile.Title = "Escoge el archivo PDF";
            openFile.Filter = "Solo se permite archivos (*.pdf) | *.pdf";
            openFile.DefaultExt = "pdf";

            if (openFile.ShowDialog() == DialogResult.OK)
            {
                if (System.IO.File.Exists(Path.Combine(rutaArchivoConsolidado, openFile.SafeFileName)))
                {
                    lblNotificaciones.Text = "Arhivo ya procesado";
                    return;
                }
                else
                {
                    File.Copy(openFile.FileName, Path.Combine(rutaArchivoConsolidado, openFile.SafeFileName), true);
                    archivoPDFPrincipal = Path.Combine(rutaArchivoConsolidado, openFile.SafeFileName);
                    lblNotificaciones.Text = "Copia exitosa - Continue con el paso 2";
                }
            }
        }

        private async void btnDividir_Click(object sender, EventArgs e)
        {
            btnDividir.Enabled = false;
            lblNotificaciones.Text = "Procesando - Por favor Espere";

            await DividirPDFenArchivos(archivoPDFPrincipal);

            lblNotificaciones.Text = "Proceso terminado - Continue con el paso 3";
            btnRenombrar.Enabled = true;
            picOkPaso1.Visible = false;
            picOkPaso2.Visible = true;
            picOkPaso3.Visible = false;
        }

        private bool ValidarDirectorio(string ruta)
        {
            try
            {
                bool exists = System.IO.Directory.Exists(ruta);

                if (!exists)
                {
                    CrearDirectorio(ruta);
                }
                else
                {
                    return true;
                }
            }
            catch (Exception e)
            {
                lblNotificaciones.Text = "Eror validando directorio + " + e.Message;
            }
            return false;
        }

        private void CrearDirectorio(string ruta)
        {
            try
            {
                System.IO.Directory.CreateDirectory(ruta);
            }
            catch (Exception e)
            {
                lblNotificaciones.Text = "Eror validando directorio + " + e.Message;
            }

        }
        private async void btnRenombrar_Click(object sender, EventArgs e)
        {

            lblNotificaciones.Text = "Procesando - Espere por favor";


            string mensajeAnotificar = "";


            bool existsAfterCreate = System.IO.Directory.Exists(rutaArchivosFecha);

            await Task.Factory.StartNew(() =>
            {
                if (existsAfterCreate)
                {
                    foreach (var pdfFileName in GetPdfFiles())
                    {
                        File.Copy(pdfFileName, Path.Combine(rutaArchivosFechaSalida, ExtractOrden(Regex.Replace(pdfFileName, @"\r\n?|\n", ""))), true);
                    }
                    mensajeAnotificar = "Proceso terminado exitosamente!    ";
                }
                else
                {
                    mensajeAnotificar = "No se pudo crear la carpeta" + rutaArchivosFechaSalida;
                }
                Thread.Sleep(10);
            });
            btnAbrirCarpeta.Visible = true;
            btnAbrirCarpeta.Enabled = true;
            lblNotificaciones.Text = mensajeAnotificar;
            btnDividir.Enabled = false;
            picOkPaso1.Visible = false;
            picOkPaso2.Visible = false;
            picOkPaso3.Visible = true;
        }

        private void OpenFolder(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    Arguments = folderPath,
                    FileName = "explorer.exe"
                };
                Process.Start(startInfo);
            }
            else
            {
                MessageBox.Show(string.Format("{0} Directory does not exist!", folderPath));
            }
        }


        private void btnAbrirCarpeta_Click(object sender, EventArgs e)
        {
            OpenFolder(rutaArchivosFechaSalida);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("         ***** INFORMACION ***** \n\n Desarrollado por: \n" +
                " Carlos Javier Aucancela" +
                " \n \n Contacto: \n aucancela.carlos@gmail.com ");
        }
    }
}
