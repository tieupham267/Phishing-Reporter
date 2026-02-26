/*
 * Developer: Abdulla Albreiki
 * Github: https://github.com/0dteam
 * licensed under the GNU General Public License v3.0
 */
 
using Microsoft.Office.Core;
using PhishingReporter.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PhishingReporter
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private static readonly NLog.Logger Logger = AppLogger.Instance.GetCurrentClassLogger();
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        public Bitmap getGroup1Image(IRibbonControl control)
        {
            try
            {
                return Resources.phishing;
            }
            catch (System.Exception ex)
            {
                Logger.Error(ex, "Unhandled exception in getGroup1Image callback");
                return null;
            }
        }

        // Functions
        public async void reportPhishing(Office.IRibbonControl control)
        {
            try
            {
                Logger.Info("Report phishing button clicked");
                var areYouSure = MessageBox.Show("Do you want to report this email to the Information Security Team as a potential phishing attempt?", "Are you sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if(areYouSure == DialogResult.Yes)
                {
                    Logger.Info("User confirmed report submission");
                    await reportPhishingEmailToSecurityTeamAsync(control).ConfigureAwait(false);
                }
                else
                {
                    Logger.Info("User cancelled report submission");
                }
            }
            catch (System.Exception ex)
            {
                Logger.Error(ex, "Unhandled exception in reportPhishing callback");
                try
                {
                    MessageBox.Show("An unexpected error occurred. Please try again or contact support.", "Error");
                }
                catch
                {
                    // Swallow silently — even MessageBox can fail in degraded COM states
                }
            }
        }

        /*
         *  Helper functions 
         */

        private async Task reportPhishingEmailToSecurityTeamAsync(IRibbonControl control)
        {
            Logger.Info("Processing selected email for phishing report");

            Selection selection = null;
            MailItem mailItem = null;

            try
            {
                selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;

                if (selection.Count < 1)
                {
                    MessageBox.Show("Select an email before reporting.", "Error");
                    return;
                }
                if (selection.Count > 1)
                {
                    MessageBox.Show("You can report 1 email at a time.", "Error");
                    return;
                }

                if (!(selection[1] is Outlook.MeetingItem
                    || selection[1] is Outlook.ContactItem
                    || selection[1] is Outlook.AppointmentItem
                    || selection[1] is Outlook.TaskItem
                    || selection[1] is Outlook.MailItem))
                {
                    MessageBox.Show("You cannot report this item", "Error");
                    return;
                }

                mailItem = selection[1] as MailItem;

                // ALL OOM data extraction happens here, on the UI thread
                EmailReport report = ExtractEmailReport(selection[1], mailItem);
                Logger.Info("Email data extracted, item type: {0}", report.ItemType);

                // Delegate to orchestrator — the only await in this method
                await ReportOrchestrator.ExecuteAsync(
                    report,
                    Globals.ThisAddIn.Application,
                    selection[1],
                    mailItem).ConfigureAwait(false);
            }
            catch (System.Exception ex)
            {
                Logger.Error(ex, "Error during report processing");

                try
                {
                    MessageBox.Show(
                        "There was an error! An automatic email was sent to the support to resolve the issue.",
                        "Do not worry");
                }
                catch
                {
                    // Swallow — MessageBox can fail in degraded COM states
                }

                SendErrorEmail(ex);
            }
            finally
            {
                if (mailItem != null) { try { Marshal.ReleaseComObject(mailItem); } catch { } mailItem = null; }
                if (selection != null) { try { Marshal.ReleaseComObject(selection); } catch { } selection = null; }
            }
        }

        /// <summary>
        /// Extracts all Outlook OOM data into an immutable EmailReport.
        /// MUST be called on the UI thread — accesses COM objects.
        /// </summary>
        private EmailReport ExtractEmailReport(object selectedItem, MailItem mailItem)
        {
            // Detect item type
            string itemType;
            if (selectedItem is Outlook.MeetingItem) itemType = "MeetingItem";
            else if (selectedItem is Outlook.ContactItem) itemType = "ContactItem";
            else if (selectedItem is Outlook.AppointmentItem) itemType = "AppointmentItem";
            else if (selectedItem is Outlook.TaskItem) itemType = "TaskItem";
            else if (selectedItem is Outlook.MailItem) itemType = "MailItem";
            else itemType = "Unknown";

            bool isMailItem = (itemType == "MailItem");

            // Extract mail-specific data (null for non-MailItem)
            string subject = isMailItem ? mailItem.Subject : itemType;
            string headers = isMailItem ? mailItem.HeaderString() : null;
            string htmlBody = isMailItem ? mailItem.HTMLBody : null;

            // GoPhish detection (pure string parsing on extracted headers)
            string goPhishReportUrl = (headers != null)
                ? GoPhishIntegration.setReportURL(headers)
                : null;

            Logger.Info("Reported item type: {0}", itemType);
            Logger.Info("GoPhish header check: {0}", goPhishReportUrl != null ? "found" : "not found");

            // Pre-compute URL analysis from extracted HTML (no OOM access in UrlExtractor)
            UrlExtractionResult urlAnalysis = (htmlBody != null)
                ? UrlExtractor.ExtractUrls(htmlBody)
                : new UrlExtractionResult(
                    Array.Empty<string>(),
                    Array.Empty<string>());

            // Pre-compute attachment hashes (OOM access to Attachments collection)
            IReadOnlyList<AttachmentHashResult> attachmentHashes = isMailItem
                ? ExtractAttachmentHashes(mailItem)
                : (IReadOnlyList<AttachmentHashResult>)Array.Empty<AttachmentHashResult>();

            // Pre-compute formatted sections (OOM access in each method)
            string userInfoSection = GetCurrentUserInfos();
            string basicInfoSection = isMailItem ? GetBasicInfo(mailItem) : null;
            string pluginDetailsSection = GetPluginDetails();

            return new EmailReport(
                itemType: itemType,
                isMailItem: isMailItem,
                subject: subject,
                headers: headers,
                htmlBody: htmlBody,
                goPhishReportUrl: goPhishReportUrl,
                urlAnalysis: urlAnalysis,
                attachmentHashes: attachmentHashes,
                userInfoSection: userInfoSection,
                basicInfoSection: basicInfoSection,
                pluginDetailsSection: pluginDetailsSection);
        }

        /// <summary>
        /// Extracts hashes for all attachments. Handles COM object lifecycle.
        /// </summary>
        private IReadOnlyList<AttachmentHashResult> ExtractAttachmentHashes(MailItem mailItem)
        {
            Outlook.Attachments attachments = null;
            try
            {
                attachments = mailItem.Attachments;
                var hashes = new List<AttachmentHashResult>(attachments.Count);

                for (int i = 1; i <= attachments.Count; i++)
                {
                    Outlook.Attachment a = null;
                    try
                    {
                        a = attachments[i];
                        hashes.Add(AttachmentHasher.ComputeHashes(a));
                    }
                    finally
                    {
                        if (a != null) { try { Marshal.ReleaseComObject(a); } catch { } a = null; }
                    }
                }

                return hashes.AsReadOnly();
            }
            finally
            {
                if (attachments != null) { try { Marshal.ReleaseComObject(attachments); } catch { } attachments = null; }
            }
        }

        /// <summary>
        /// Sends an error notification email to support. Stays in Ribbon.cs because
        /// it needs OOM (Application.CreateItem). May be called from a background thread
        /// after an await — wrapped in try/catch for COMException safety.
        /// </summary>
        private void SendErrorEmail(System.Exception ex)
        {
            MailItem errorEmail = null;
            try
            {
                errorEmail = (MailItem)Globals.ThisAddIn.Application
                    .CreateItem(OlItemType.olMailItem);
                errorEmail.To = Properties.Settings.Default.support_email;
                errorEmail.Subject = "[Outlook Addin Error]";
                errorEmail.Body = "Addin error message: " + ex;
                errorEmail.Save();
                errorEmail.Send();
                Logger.Info("Error email sent to support: {0}", Properties.Settings.Default.support_email);
            }
            catch (System.Exception sendEx)
            {
                // May fail with COMException if called from background thread after await
                Logger.Error(sendEx, "Failed to send error email to support");
            }
            finally
            {
                if (errorEmail != null) { try { Marshal.ReleaseComObject(errorEmail); } catch { } errorEmail = null; }
            }
        }

        public String GetBasicInfo(MailItem mailItem)
        {
            Outlook.MAPIFolder parentFolder = null;

            try
            {
                parentFolder = mailItem.Parent as Outlook.MAPIFolder;
                string FolderLocation = parentFolder.FolderPath;
                string basicInfo = "---------- Basic Info ----------";
                basicInfo += "\n - Reported from: \"" + FolderLocation + "\" Folder";
                basicInfo += "\n - OS: " + Environment.OSVersion + " " + (Environment.Is64BitOperatingSystem ? "(64bit)" : "(32bit)");
                basicInfo += "\n - Agent: " + Globals.ThisAddIn.Application.Name + " "  + Globals.ThisAddIn.Application.Version;
                basicInfo += "\n - Suspecious emails reported: " + Properties.Settings.Default.suspecious_reports_counter;
                basicInfo += "\n - GoPhish campaigns reported: " + Properties.Settings.Default.gophish_reports_counter;
                basicInfo += "\n";
                return basicInfo;
            }
            finally
            {
                if (parentFolder != null) { try { Marshal.ReleaseComObject(parentFolder); } catch { } parentFolder = null; }
            }
        }


        public String GetCurrentUserInfos()
        {
            string str = "---------- User Information ----------";
            str += "\n - Domain:" + Environment.UserDomainName;
            str += "\n - Username:" + Environment.UserName;
            str += "\n - Machine name:" + Environment.MachineName;

            Outlook.NameSpace session = null;
            Outlook.Recipient currentUserRecipient = null;
            Outlook.AddressEntry addrEntry = null;
            Outlook.ExchangeUser currentUser = null;

            try
            {
                session = Globals.ThisAddIn.Application.Session;
                currentUserRecipient = session.CurrentUser;
                addrEntry = currentUserRecipient.AddressEntry;

                if (addrEntry.Type == "EX")
                {
                    currentUser = addrEntry.GetExchangeUser();
                    if (currentUser != null)
                    {
                        str += "\n - Name: " + currentUser.Name;
                        str += "\n - STMP address: " + currentUser.PrimarySmtpAddress;
                        str += "\n - Title: " + currentUser.JobTitle;
                        str += "\n - Department: " + currentUser.Department;
                        str += "\n - Location: " + currentUser.OfficeLocation;
                        str += "\n - Business phone: " + currentUser.BusinessTelephoneNumber;
                        str += "\n - Mobile phone: " + currentUser.MobileTelephoneNumber;
                    }
                }
            }
            finally
            {
                if (currentUser != null) { try { Marshal.ReleaseComObject(currentUser); } catch { } currentUser = null; }
                if (addrEntry != null) { try { Marshal.ReleaseComObject(addrEntry); } catch { } addrEntry = null; }
                if (currentUserRecipient != null) { try { Marshal.ReleaseComObject(currentUserRecipient); } catch { } currentUserRecipient = null; }
                if (session != null) { try { Marshal.ReleaseComObject(session); } catch { } session = null; }
            }

            return str + "\n";
        }



        public String GetPluginDetails()
        {
            string pluginDetails = "---------- Report Phishing Plugin ----------";
            pluginDetails += "\n - Version: " + Properties.Settings.Default.plugin_version;
            pluginDetails += "\n - Usage: Report phishing emails to the Information Security Team.";
            pluginDetails += "\n - Support: " + Properties.Settings.Default.support_email;
            
            pluginDetails += "\n - Developer: Abdulla Albreiki (aalbraiki@hotmail.com)"; // You may delete this line if you like :)
            return pluginDetails;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            try
            {
                return GetResourceText("PhishingReporter.Ribbon.xml");
            }
            catch (System.Exception ex)
            {
                Logger.Error(ex, "Unhandled exception in GetCustomUI callback");
                return null;
            }
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            try
            {
                this.ribbon = ribbonUI;
            }
            catch (System.Exception ex)
            {
                Logger.Error(ex, "Unhandled exception in Ribbon_Load callback");
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }

    public static class MailItemExtensions
    {
        private const string HeaderRegex =
            @"^(?<header_key>[-A-Za-z0-9]+)(?<seperator>:[ \t]*)" +
                "(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)";
        private const string TransportMessageHeadersSchema =
            "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

        public static string[] Headers(this MailItem mailItem, string name)
        {
            var headers = mailItem.HeaderLookup();
            if (headers.Contains(name))
                return headers[name].ToArray();
            return new string[0];
        }

        public static ILookup<string, string> HeaderLookup(this MailItem mailItem)
        {
            var headerString = mailItem.HeaderString();
            var headerMatches = Regex.Matches
                (headerString, HeaderRegex, RegexOptions.Multiline).Cast<Match>();
            return headerMatches.ToLookup(
                h => h.Groups["header_key"].Value,
                h => h.Groups["header_value"].Value);
        }

        public static string HeaderString(this MailItem mailItem)
        {
            return (string)mailItem.PropertyAccessor
                .GetProperty(TransportMessageHeadersSchema);
        }

    }
}
 