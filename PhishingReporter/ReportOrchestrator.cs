using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PhishingReporter
{
    /// <summary>
    /// Orchestrates the full phishing report workflow using pre-extracted data.
    ///
    /// THREADING CONTRACT: ExecuteAsync MUST be called from the Outlook UI (STA) thread.
    /// All Outlook OOM access occurs before any await boundary in every execution path.
    /// After an await, only thread-safe operations (Settings, MessageBox, logging) are used.
    ///
    /// WARNING: Do NOT add Outlook OOM access after any await call in this class.
    /// Doing so will cause COMException 0x8001010E on the thread pool thread.
    /// </summary>
    internal static class ReportOrchestrator
    {
        private static readonly NLog.Logger Logger =
            AppLogger.Instance.GetCurrentClassLogger();

        /// <summary>
        /// Executes the full report workflow using pre-extracted email data.
        /// </summary>
        /// <param name="report">Immutable email data extracted on the UI thread.</param>
        /// <param name="application">Outlook Application for creating report email (OOM, UI thread only).</param>
        /// <param name="originalItem">Original selected item for attaching to report email (OOM, UI thread only).</param>
        /// <param name="mailItem">Original MailItem for deletion (OOM, UI thread only). Null for non-MailItem types.</param>
        public static async Task ExecuteAsync(
            EmailReport report,
            Outlook.Application application,
            object originalItem,
            Outlook.MailItem mailItem)
        {
            if (report.GoPhishReportUrl != null)
            {
                await ExecuteGoPhishBranchAsync(report, mailItem).ConfigureAwait(false);
            }
            else
            {
                ExecuteStandardReportBranch(report, application, originalItem, mailItem);
            }
        }

        /// <summary>
        /// GoPhish branch: delete email, send async notification, increment counter.
        /// OOM access (mailItem.Delete) occurs BEFORE the await boundary.
        /// </summary>
        private static async Task ExecuteGoPhishBranchAsync(
            EmailReport report,
            Outlook.MailItem mailItem)
        {
            // OOM access — BEFORE await, on UI thread
            if (mailItem != null)
            {
                mailItem.Delete();
                Logger.Info("Reported email deleted from mailbox");
            }

            // === AWAIT BOUNDARY ===
            // WARNING: After this line, code runs on a thread pool thread.
            // No Outlook OOM access is permitted below.
            GoPhishResult result = await GoPhishIntegration
                .SendReportNotificationAsync(report.GoPhishReportUrl)
                .ConfigureAwait(false);
            Logger.Info("GoPhish notification result: {0}", result);

            // Safe from any thread — Settings.Default is not a COM object
            Properties.Settings.Default.gophish_reports_counter++;
            Properties.Settings.Default.Save();

            // MessageBox auto-marshals to UI thread
            MessageBox.Show(
                "Good job! You have reported a simulated phishing campaign sent by the Information Security Team.",
                "We have a winner!");
        }

        /// <summary>
        /// Standard report branch: compose report body, create and send report email, delete original.
        /// NO await in this method — executes entirely on the UI thread.
        /// </summary>
        private static void ExecuteStandardReportBranch(
            EmailReport report,
            Outlook.Application application,
            object originalItem,
            Outlook.MailItem mailItem)
        {
            // Safe from any thread, but we are on UI thread here
            Properties.Settings.Default.suspecious_reports_counter++;
            Properties.Settings.Default.Save();

            // Pure string building — no OOM access
            string body = ComposeReportBody(report);

            // OOM access — safe, on UI thread (no await in this method)
            Outlook.MailItem reportEmail = null;
            try
            {
                reportEmail = (Outlook.MailItem)application
                    .CreateItem(Outlook.OlItemType.olMailItem);
                reportEmail.To = Properties.Settings.Default.infosec_email;
                reportEmail.Subject = report.IsMailItem
                    ? "[POTENTIAL PHISH] " + report.Subject
                    : "[POTENTIAL PHISH] " + report.ItemType;
                reportEmail.Attachments.Add(originalItem);
                reportEmail.Body = body;
                reportEmail.Save();
                reportEmail.Send();
                Logger.Info("Report email sent to: {0}",
                    Properties.Settings.Default.infosec_email);

                if (mailItem != null)
                {
                    mailItem.Delete();
                    Logger.Info("Reported email deleted from mailbox");
                }
            }
            finally
            {
                if (reportEmail != null)
                {
                    try { Marshal.ReleaseComObject(reportEmail); }
                    catch { /* Prevent cleanup exception from propagating */ }
                    reportEmail = null;
                }
            }
        }

        /// <summary>
        /// Composes the full report email body from pre-extracted data.
        /// Pure string building — no OOM access, safe from any thread.
        /// </summary>
        private static string ComposeReportBody(EmailReport report)
        {
            string body = report.UserInfoSection;
            body += "\n";

            if (report.BasicInfoSection != null)
            {
                body += report.BasicInfoSection;
                body += "\n";
            }

            body += "---------- URLs and Attachments ----------";
            body += "\n # of unique Domains: " + report.UrlAnalysis.UniqueDomains.Count;
            foreach (string domain in report.UrlAnalysis.UniqueDomains)
            {
                body += "\n --> Domain: " + domain.Replace(":", "[:]");
            }

            body += "\n\n # of URLs: " + report.UrlAnalysis.Urls.Count;
            foreach (string url in report.UrlAnalysis.Urls)
            {
                body += "\n --> URL: " + url.Replace(":", "[:]");
            }

            body += "\n\n # of Attachments: " + report.AttachmentHashes.Count;
            foreach (var hash in report.AttachmentHashes)
            {
                body += "\n --> Attachment: " + hash.FileName
                    + " (" + hash.SizeBytes + " bytes)"
                    + "\n\t\tMD5: " + hash.Md5
                    + "\n\t\tSha256: " + hash.Sha256 + "\n";
            }

            body += "\n---------- Headers ----------";
            body += "\n" + report.Headers;
            body += "\n";
            body += report.PluginDetailsSection + "\n\n";

            return body;
        }
    }
}
