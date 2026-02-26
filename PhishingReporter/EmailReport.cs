using System;
using System.Collections.Generic;

namespace PhishingReporter
{
    /// <summary>
    /// Immutable snapshot of all data extracted from an Outlook email item.
    /// Constructed on the UI thread from OOM objects; safe to use on any thread after construction.
    ///
    /// INVARIANT: This class contains ZERO COM type references. All properties are
    /// primitive types (string, int, bool) or immutable C# objects (UrlExtractionResult,
    /// IReadOnlyList&lt;AttachmentHashResult&gt;). Any Microsoft.Office.Interop.Outlook type
    /// in this class is a bug.
    /// </summary>
    internal sealed class EmailReport
    {
        // Item identity
        public string ItemType { get; }
        public bool IsMailItem { get; }

        // Mail-specific data (null/default if not a MailItem)
        public string Subject { get; }
        public string Headers { get; }
        public string HtmlBody { get; }

        // GoPhish detection result (null if not a GoPhish campaign email)
        public string GoPhishReportUrl { get; }

        // Pre-computed analysis results (immutable plain C# objects)
        public UrlExtractionResult UrlAnalysis { get; }
        public IReadOnlyList<AttachmentHashResult> AttachmentHashes { get; }

        // Pre-formatted report sections (extracted from OOM on UI thread)
        // These are pre-computed strings because the raw OOM properties (ExchangeUser,
        // MAPIFolder) are COM objects that cannot be stored safely across threads.
        public string UserInfoSection { get; }
        public string BasicInfoSection { get; }
        public string PluginDetailsSection { get; }

        public EmailReport(
            string itemType,
            bool isMailItem,
            string subject,
            string headers,
            string htmlBody,
            string goPhishReportUrl,
            UrlExtractionResult urlAnalysis,
            IReadOnlyList<AttachmentHashResult> attachmentHashes,
            string userInfoSection,
            string basicInfoSection,
            string pluginDetailsSection)
        {
            ItemType = itemType ?? throw new ArgumentNullException(nameof(itemType));
            IsMailItem = isMailItem;
            Subject = subject;
            Headers = headers;
            HtmlBody = htmlBody;
            GoPhishReportUrl = goPhishReportUrl;
            UrlAnalysis = urlAnalysis ?? throw new ArgumentNullException(nameof(urlAnalysis));
            AttachmentHashes = attachmentHashes ?? throw new ArgumentNullException(nameof(attachmentHashes));
            UserInfoSection = userInfoSection ?? throw new ArgumentNullException(nameof(userInfoSection));
            BasicInfoSection = basicInfoSection;
            PluginDetailsSection = pluginDetailsSection ?? throw new ArgumentNullException(nameof(pluginDetailsSection));
        }
    }
}
