/*
 * Developer: Abdulla Albreiki
 * Github: https://github.com/0dteam
 * licensed under the GNU General Public License v3.0
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace PhishingReporter
{
    public partial class ThisAddIn
    {
        private static readonly NLog.Logger Logger = AppLogger.Instance.GetCurrentClassLogger();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var sw = Stopwatch.StartNew();
            Logger.Info("PhishingReporter add-in startup begin");

            // STRT-01: Register deferred init handler that runs AFTER Outlook
            // finishes loading all add-ins, outside the resiliency measurement window.
            this.Application.Startup += Application_Startup;

            sw.Stop();
            Logger.Info("PhishingReporter add-in startup complete ({0:F1} ms)", sw.Elapsed.TotalMilliseconds);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
            Logger.Info("PhishingReporter add-in shutdown");
            AppLogger.Instance.Shutdown();
        }

        private void Application_Startup()
        {
            Logger.Info("PhishingReporter deferred initialization begin");
            // STRT-01: This runs AFTER Outlook finishes loading all add-ins,
            // outside the resiliency measurement window.
            // Currently no heavy initialization needed here, but this is the
            // correct place for any future startup work (e.g., warming up
            // GoPhishIntegration, pre-validating configuration).
            Logger.Info("PhishingReporter deferred initialization complete");
        }

        /// <summary>
        /// STRT-02: Direct return bypasses VSTO Ribbon Designer reflection scan.
        /// The VSTO runtime would otherwise scan all assemblies for IRibbonExtension
        /// implementations, adding 100-500 ms to startup. This override eliminates
        /// that scan by returning the Ribbon instance directly.
        /// </summary>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
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
