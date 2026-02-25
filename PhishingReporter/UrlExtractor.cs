using System;
using System.Collections.Generic;
using System.Linq;
using HtmlAgilityPack;

namespace PhishingReporter
{
    /// <summary>
    /// Result of URL extraction from an email HTML body.
    /// Immutable data transfer object.
    /// </summary>
    internal sealed class UrlExtractionResult
    {
        public IReadOnlyList<string> Urls { get; }
        public IReadOnlyList<string> UniqueDomains { get; }

        public UrlExtractionResult(
            IReadOnlyList<string> urls,
            IReadOnlyList<string> uniqueDomains)
        {
            Urls = urls;
            UniqueDomains = uniqueDomains;
        }
    }

    /// <summary>
    /// Extracts URLs and domains from email HTML body.
    /// Replaces the broken Contains("a") filter from the original Ribbon.cs implementation.
    /// </summary>
    internal static class UrlExtractor
    {
        private static readonly NLog.Logger Logger =
            AppLogger.Instance.GetCurrentClassLogger();

        /// <summary>
        /// Extracts all href values from anchor tags in the provided HTML.
        /// </summary>
        /// <param name="emailHtmlBody">Raw HTML body of the email.</param>
        /// <returns>Extraction result with URLs and unique domains.</returns>
        public static UrlExtractionResult ExtractUrls(string emailHtmlBody)
        {
            if (string.IsNullOrEmpty(emailHtmlBody))
            {
                return new UrlExtractionResult(
                    Array.Empty<string>(),
                    Array.Empty<string>());
            }

            var doc = new HtmlDocument();
            doc.LoadHtml(emailHtmlBody);

            var urlNodes = doc.DocumentNode.SelectNodes("//a[@href]");
            if (urlNodes == null)
            {
                return new UrlExtractionResult(
                    Array.Empty<string>(),
                    Array.Empty<string>());
            }

            var urls = new List<string>();
            var domains = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var link in urlNodes)
            {
                string href = link.GetAttributeValue("href", "");
                if (string.IsNullOrWhiteSpace(href))
                    continue;

                // BUG FIX (BUGF-01): No Contains("a") filter.
                // All href values are captured regardless of content.
                urls.Add(href);

                // Domain extraction
                try
                {
                    domains.Add(new Uri(href).Host);
                }
                catch (UriFormatException)
                {
                    // Handle mailto: links -- extract domain from email address
                    int atIndex = href.IndexOf('@');
                    if (atIndex >= 0)
                    {
                        string emailDomain = href.Substring(atIndex + 1);
                        // Trim any trailing path/query from the domain
                        int slashIndex = emailDomain.IndexOf('/');
                        if (slashIndex >= 0)
                            emailDomain = emailDomain.Substring(0, slashIndex);

                        if (!string.IsNullOrWhiteSpace(emailDomain))
                            domains.Add(emailDomain);
                    }
                    else
                    {
                        Logger.Warn("Unparseable URL skipped for domain extraction: {0}", href);
                    }
                }
            }

            Logger.Info("Extracted {0} URLs and {1} unique domains from email body",
                urls.Count, domains.Count);

            return new UrlExtractionResult(
                urls.AsReadOnly(),
                domains.ToList().AsReadOnly());
        }
    }
}
