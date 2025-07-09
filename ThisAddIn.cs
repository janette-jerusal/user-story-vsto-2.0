// ThisAddIn class implementation
using System;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace UserStorySimilarityAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // This will trigger when the add-in starts
            System.Diagnostics.Debug.WriteLine("UserStorySimilarityAddIn loaded.");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Optional: Add cleanup code here if needed
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
