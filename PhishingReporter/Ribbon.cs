/*
 * Developer: Abdulla Albreiki
 * Github: https://github.com/0dteam
 * licensed under the GNU General Public License v3.0
 */
 
using Microsoft.Office.Core;
using PhishingReporter.Properties;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
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
        public void reportPhishing(Office.IRibbonControl control)
        {
            try
            {
                Logger.Info("Report phishing button clicked");
                var areYouSure = MessageBox.Show("Do you want to report this email to the Information Security Team as a potential phishing attempt?", "Are you sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if(areYouSure == DialogResult.Yes)
                {
                    Logger.Info("User confirmed report submission");
                    reportPhishingEmailToSecurityTeam(control);
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

        private void reportPhishingEmailToSecurityTeam(IRibbonControl control)
        {
            Logger.Info("Processing selected email for phishing report");

            Selection selection = null;
            MailItem mailItem = null;
            MailItem reportEmail = null;

            try
            {
                selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
                string reportedItemType = "NaN"; // email, contact, appointment ...etc
                string reportedItemHeaders = "NaN";

                if (selection.Count < 1) // no item is selected
                {
                    MessageBox.Show("Select an email before reporting.", "Error");
                }
                else if (selection.Count > 1) // many items selected
                {
                    MessageBox.Show("You can report 1 email at a time.", "Error");
                }
                else // only 1 item is selected
                {
                    if (selection[1] is Outlook.MeetingItem || selection[1] is Outlook.ContactItem || selection[1] is Outlook.AppointmentItem || selection[1] is Outlook.TaskItem || selection[1] is Outlook.MailItem)
                    {
                        // Identify the reported item type
                        if (selection[1] is Outlook.MeetingItem)
                        {
                            reportedItemType = "MeetingItem";
                        }
                        else if (selection[1] is Outlook.ContactItem)
                        {
                            reportedItemType = "ContactItem";
                        }
                        else if (selection[1] is Outlook.AppointmentItem)
                        {
                            reportedItemType = "AppointmentItem";
                        }
                        else if (selection[1] is Outlook.TaskItem)
                        {
                            reportedItemType = "TaskItem";
                        }
                        else if (selection[1] is Outlook.MailItem)
                        {
                            reportedItemType = "MailItem";
                        }

                        Logger.Info("Reported item type: {0}", reportedItemType);

                        // Prepare Reported Email
                        mailItem = (reportedItemType == "MailItem") ? selection[1] as MailItem : null;

                        reportEmail = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                        reportEmail.Attachments.Add(selection[1] as Object);

                        try
                        {
                            reportEmail.To = Properties.Settings.Default.infosec_email;
                            reportEmail.Subject = (reportedItemType == "MailItem") ? "[POTENTIAL PHISH] " + mailItem.Subject : "[POTENTIAL PHISH] " + reportedItemType;

                            // Get Email Headers
                            if (reportedItemType == "MailItem")
                            {
                                reportedItemHeaders = mailItem.HeaderString();
                            }
                            else
                            {
                                reportedItemHeaders = "Headers were not extracted because the reported item is not an email. It is " + reportedItemType;
                            }

                            // Check if the email is a simulated phishing campaign by Information Security Team
                            string simulatedPhishingURL = GoPhishIntegration.setReportURL(reportedItemHeaders);
                            Logger.Info("GoPhish header check: {0}", simulatedPhishingURL != null ? "found" : "not found");

                            if (simulatedPhishingURL != null)
                            {
                                GoPhishResult goPhishResult = GoPhishIntegration.sendReportNotificationToServer(simulatedPhishingURL);
                                Logger.Info("GoPhish notification result: {0}", goPhishResult);

                                // Update GoPhish Campaigns Reported counter
                                Properties.Settings.Default.gophish_reports_counter++;
                                Properties.Settings.Default.Save();

                                // Thanks
                                MessageBox.Show("Good job! You have reported a simulated phishing campaign sent by the Information Security Team.", "We have a winner!");
                            }
                            else
                            {
                                // Update Suspecious Emails Reported counter
                                Properties.Settings.Default.suspecious_reports_counter++;
                                Properties.Settings.Default.Save();

                                // Prepare the email body
                                reportEmail.Body = GetCurrentUserInfos();
                                reportEmail.Body += "\n";
                                reportEmail.Body += GetBasicInfo(mailItem);
                                reportEmail.Body += "\n";
                                reportEmail.Body += GetURLsAndAttachmentsInfo(mailItem);
                                reportEmail.Body += "\n";
                                reportEmail.Body += "---------- Headers ----------";
                                reportEmail.Body += "\n" + reportedItemHeaders;
                                reportEmail.Body += "\n";
                                reportEmail.Body += GetPluginDetails() + "\n\n";

                                Logger.Info("Report email composed for: {0}", mailItem.Subject);
                                reportEmail.Save();
                                reportEmail.Send();
                                Logger.Info("Report email sent to: {0}", Properties.Settings.Default.infosec_email);
                            }

                            // Delete the reported email
                            mailItem.Delete();
                            Logger.Info("Reported email deleted from mailbox");
                        }
                        catch (System.Exception ex)
                        {
                            Logger.Error(ex, "Error during report processing");
                            MessageBox.Show("There was an error! An automatic email was sent to the support to resolve the issue.", "Do not worry");

                            MailItem errorEmail = null;
                            try
                            {
                                errorEmail = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                                errorEmail.To = Properties.Settings.Default.support_email;
                                errorEmail.Subject = "[Outlook Addin Error]";
                                errorEmail.Body = ("Addin error message: " + ex);
                                errorEmail.Save();
                                errorEmail.Send();
                            }
                            finally
                            {
                                if (errorEmail != null) { try { Marshal.ReleaseComObject(errorEmail); } catch { } errorEmail = null; }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("You cannot report this item", "Error");
                    }
                }
            }
            finally
            {
                if (reportEmail != null) { try { Marshal.ReleaseComObject(reportEmail); } catch { } reportEmail = null; }
                if (mailItem != null) { try { Marshal.ReleaseComObject(mailItem); } catch { } mailItem = null; }
                if (selection != null) { try { Marshal.ReleaseComObject(selection); } catch { } selection = null; }
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

        public String GetURLsAndAttachmentsInfo(MailItem mailItem)
        {
            string result = "---------- URLs and Attachments ----------";

            // URL extraction delegated to UrlExtractor (QUAL-03, BUGF-01)
            var urlResult = UrlExtractor.ExtractUrls(mailItem.HTMLBody);

            result += "\n # of unique Domains: " + urlResult.UniqueDomains.Count;
            foreach (string domain in urlResult.UniqueDomains)
            {
                result += "\n --> Domain: " + domain.Replace(":", "[:]");
            }

            result += "\n\n # of URLs: " + urlResult.Urls.Count;
            foreach (string url in urlResult.Urls)
            {
                result += "\n --> URL: " + url.Replace(":", "[:]");
            }

            // Attachment hashing delegated to AttachmentHasher (QUAL-04, BUGF-05)
            Outlook.Attachments attachments = null;
            try
            {
                attachments = mailItem.Attachments;
                result += "\n\n # of Attachments: " + attachments.Count;

                for (int i = 1; i <= attachments.Count; i++)
                {
                    Attachment a = null;
                    try
                    {
                        a = attachments[i];
                        var hashResult = AttachmentHasher.ComputeHashes(a);
                        result += "\n --> Attachment: " + hashResult.FileName
                            + " (" + hashResult.SizeBytes + " bytes)"
                            + "\n\t\tMD5: " + hashResult.Md5
                            + "\n\t\tSha256: " + hashResult.Sha256 + "\n";
                    }
                    finally
                    {
                        if (a != null) { try { Marshal.ReleaseComObject(a); } catch { } a = null; }
                    }
                }
            }
            finally
            {
                if (attachments != null) { try { Marshal.ReleaseComObject(attachments); } catch { } attachments = null; }
            }

            return result;
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
 