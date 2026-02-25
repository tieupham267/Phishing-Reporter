/*
 * Developer: Abdulla Albreiki
 * Github: https://github.com/0dteam
 * licensed under the GNU General Public License v3.0
 */

using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Text;
using System;
using System.Security.Authentication;
using PhishingReporter;

namespace PhishingReporter
{
    /// <summary>
    /// Result of GoPhish campaign detection and reporting.
    /// Replaces magic strings "OK", "ERROR", and "NaN".
    /// </summary>
    internal enum GoPhishResult
    {
        /// <summary>No GoPhish header found in email -- not a simulated campaign.</summary>
        NotFound,

        /// <summary>GoPhish campaign detected and successfully reported to server.</summary>
        Reported,

        /// <summary>GoPhish campaign detected but reporting failed (network error, timeout).</summary>
        Error
    }

    static class GoPhishIntegration
    {
        private static readonly NLog.Logger Logger = AppLogger.Instance.GetCurrentClassLogger();

        static string GoPhishURL = PhishingReporter.Properties.Settings.Default.gophish_url + ":" + Properties.Settings.Default.gophish_listener_port;
        static string URLrequest = GoPhishURL + "/report?rid=USERID";
        static string GoPhishHeader = PhishingReporter.Properties.Settings.Default.gophish_custom_header;
        static string WebExpID = GoPhishHeader + @": [0-9a-zA-Z]+";
        static string WebExpPrefix = GoPhishHeader + @": ";

        // This function constructs GoPhish report url from a custom header in the simulated phishing campaign email
        public static string setReportURL(string headers)
        {
            Logger.Debug("Checking email headers for GoPhish campaign marker");
            // Extract GoPhish Custom Header (X-GOPHISH-AJSMN: USERID0123)
            var match = new Regex(WebExpID).Match(headers);

            foreach (var group in match.Groups)
            {
                if(group.ToString().Trim()!=string.Empty)
                {
                    // Extract User ID from the header (USERID0123)
                    string user_id = group.ToString().Replace(WebExpPrefix, string.Empty);

                    // Build reporting URL, something like this -> https[:]//GOPHISHURL:PORT/report?rid=USERID
                    string report_url = URLrequest.Replace(@"USERID", user_id);
                    Logger.Info("GoPhish campaign detected, report URL: {0}", report_url);
                    return report_url;
                }
            }

            // else, no header was found -> No report tracking URL
            Logger.Debug("No GoPhish header found in email");
            return null;
        }

        public const SslProtocols _strTls12 = (SslProtocols)0x00000C00;
        public const SecurityProtocolType Tls12 = (SecurityProtocolType)_strTls12;

        public static GoPhishResult sendReportNotificationToServer(string reportURL)
        {
            Logger.Info("Sending GoPhish report notification to: {0}", reportURL);
            ServicePointManager.SecurityProtocol = Tls12;

            try
            {
                var request = (HttpWebRequest)WebRequest.Create(reportURL);
                var response = (HttpWebResponse)request.GetResponse();
                string html = new StreamReader(response.GetResponseStream()).ReadToEnd();
                Logger.Info("GoPhish notification sent successfully");
                return GoPhishResult.Reported;
            }
            catch (System.Exception exc)
            {
                Logger.Error(exc, "GoPhish notification failed");
                return GoPhishResult.Error; // GoPhish Listener is not responding or there is no Internet connection.
            }
        }
    }
}

