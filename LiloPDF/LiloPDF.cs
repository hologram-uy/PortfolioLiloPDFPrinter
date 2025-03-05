using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace LiloPDF
{
    public partial class LiloPDF: Form
    {
        private Dictionary<string, string> fileDictionary = new Dictionary<string, string>();
        public LiloPDF()
        {
            InitializeComponent();
        }
        private void LiloPDF_Load(object sender, EventArgs e)
        {
            try
            {
                // Load printers on list-box
                foreach (string sPrinters in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                {
                    lboxPrinters.Items.Add(sPrinters);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema: " + "\n" + ex.Message);
            }
        }
        private void btnAddPDF_Click(object sender, EventArgs e)
        {
            try
            {
                ofdialog1.Multiselect = true;
                ofdialog1.Filter = "Archivos PDF|*.pdf";
                ofdialog1.Title = "Seleccione los archivos PDF";

                if (ofdialog1.ShowDialog() == DialogResult.OK)
                {
                    string[] filePaths = ofdialog1.FileNames;    // Rutas completas
                    string[] fileNames = ofdialog1.SafeFileNames; // Solo nombres

                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        if (!fileDictionary.ContainsKey(fileNames[i])) // Evita duplicados
                        {
                            fileDictionary[fileNames[i]] = filePaths[i];
                            lboxFiles.Items.Add(fileNames[i]); // Solo muestra el nombre
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema: " + ex.Message);
            }
        }
        private void btnDelPDF_Click(object sender, EventArgs e)
        {
            try
            {
                var selectedItems = lboxFiles.SelectedItems;
                if (lboxFiles.SelectedIndex != -1)
                {
                    for (int i = selectedItems.Count - 1; i >= 0; i--)
                        lboxFiles.Items.Remove(selectedItems[i]);
                }
                else
                {
                    MessageBox.Show("Tiene que seleccionar un elemento de la lista para eliminarlo.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema: " + "\n" + ex.Message);
            }
        }
        private void btnDelAll_Click(object sender, EventArgs e)
        {
            try
            {
                if (lboxFiles == null)
                {
                    MessageBox.Show("Error: ListBox no está inicializado.");
                    return;
                }

                if (lboxFiles.DataSource != null)
                {
                    lboxFiles.DataSource = null;
                }
                else
                {
                    lboxFiles.Items.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema:\n" + ex.Message);
            }
        }
        private void btnPrtPDF_Click(object sender, EventArgs e)
        {
            try
            {
                if (IsEmpty())
                {
                    MessageBox.Show("No hay ningún PDF para imprimir en la lista.");
                    return;
                }

                if (lboxPrinters.SelectedIndex == -1)
                {
                    MessageBox.Show("Debe seleccionar una Impresora de la lista.");
                    return;
                }

                if (MessageBox.Show("¿Confirma la impresión?", "Se están por enviar documentos a imprimir.", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                {
                    var selectedPrinter = lboxPrinters.SelectedItem.ToString();

                    foreach (string pdfName in lboxFiles.Items)
                    {                        
                        if (fileDictionary.TryGetValue(pdfName, out string pdfPath))
                        {
                            print(selectedPrinter, pdfPath);
                        }
                        else
                        {
                            MessageBox.Show($"No se encontró la ruta del archivo: {pdfName}");
                        }
                    }
                }

                MessageBox.Show("Se imprimieron todos los documentos.");

                lboxFiles.Items.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema: " + "\n" + ex.Message);
            }
        }
        private void print(string printerName, string fileName)
        {
            try
            {
                string sumatraPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "SumatraPDF.exe");

                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = sumatraPath;//"C:\\SumatraPDF\\SumatraPDF.exe";
                proc.StartInfo.Arguments = "-print-to " + '"' + printerName + '"' + " " + '"' + fileName + '"';
                proc.StartInfo.RedirectStandardError = false;
                proc.StartInfo.RedirectStandardOutput = false;
                proc.StartInfo.UseShellExecute = true;
                proc.Start();
                proc.WaitForExit();
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("InboundServicioImpresion", ex.Message + " " + ex.StackTrace);
            }
        }
        private bool IsEmpty()
        {
            try
            {
                return lboxFiles.Items.Count == 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema:\n" + ex.Message);
                return false;
            }
        }
        private void pboxLogo_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://www.hologram.com.uy",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema: " + "\n" + ex.Message);
            }
        }
        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show($"" +
                $"LiloPDF Multi Printer \n\n" +
                $"Para empezar, haga click en 'Agregar' para buscar los PDF que desea imprimir." +
                $"Puede seleccionar múltiples archivos a la vez.\n\n" +
                $"Para eliminar, simplemente utilice los botones 'Eliminar' o 'Borrar Todo'.\n\n" +
                $"Seleccione la impresora a la que desee enviar los PDF desde el listado inferior y presione el botón 'Imprimir'.");
        }
    }
}
