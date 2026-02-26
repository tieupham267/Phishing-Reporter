/*
 * Developer: Abdulla Albreiki
 * Github: https://github.com/0dteam
 * licensed under the GNU General Public License v3.0
 */

using System;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Polly;
using Polly.Retry;
using Polly.Timeout;

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

        // NETW-02: Static singleton prevents socket exhaustion
        private static readonly HttpClient HttpClientInstance;

        // NETW-03 + NETW-04: Resilience pipeline with retry + timeout
        private static readonly ResiliencePipeline<HttpResponseMessage> Pipeline;

        static GoPhishIntegration()
        {
            // NETW-02: Singleton HttpClient -- never dispose
            HttpClientInstance = new HttpClient
            {
                // Let Polly manage all timeouts (Pitfall 5: HttpClient.Timeout vs Polly conflict)
                Timeout = System.Threading.Timeout.InfiniteTimeSpan
            };

            // .NET Framework 4.8 DNS workaround: no SocketsHttpHandler available
            // ConnectionLeaseTimeout forces periodic connection recycling for DNS changes
            try
            {
                var baseUrl = Properties.Settings.Default.gophish_url;
                var port = Properties.Settings.Default.gophish_listener_port;
                if (!string.IsNullOrEmpty(baseUrl))
                {
                    var baseUri = new Uri(baseUrl + ":" + port);
                    var sp = ServicePointManager.FindServicePoint(baseUri);
                    sp.ConnectionLeaseTimeout = 60_000; // 1 minute
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "Failed to configure ServicePoint for GoPhish URL -- DNS recycling disabled");
            }

            // NETW-03 + NETW-04: Retry with exponential backoff + per-attempt timeout
            // Strategy order is FIFO: retry wraps timeout (timeout is per-attempt)
            Pipeline = new ResiliencePipelineBuilder<HttpResponseMessage>()
                .AddRetry(new RetryStrategyOptions<HttpResponseMessage>
                {
                    ShouldHandle = new PredicateBuilder<HttpResponseMessage>()
                        .Handle<HttpRequestException>()
                        .Handle<TimeoutRejectedException>(),
                    MaxRetryAttempts = 3,
                    Delay = TimeSpan.FromSeconds(1),
                    BackoffType = DelayBackoffType.Exponential,
                    UseJitter = true,
                    OnRetry = static args =>
                    {
                        var logger = AppLogger.Instance.GetCurrentClassLogger();
                        logger.Warn("GoPhish retry attempt {0} after {1}ms delay",
                            args.AttemptNumber, args.RetryDelay.TotalMilliseconds);
                        return default;
                    }
                })
                .AddTimeout(TimeSpan.FromSeconds(10)) // NETW-03: 10-second per-attempt timeout
                .Build();
        }

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

        /// <summary>
        /// Sends a report notification to the GoPhish server asynchronously.
        /// NETW-01: Does not block the UI thread.
        /// BUGF-04: HttpClient manages response lifecycle (no manual dispose needed).
        /// </summary>
        public static async Task<GoPhishResult> SendReportNotificationAsync(string reportUrl)
        {
            Logger.Info("Sending GoPhish report notification to: {0}", reportUrl);

            try
            {
                using (var response = await Pipeline.ExecuteAsync(
                    async ct => await HttpClientInstance.GetAsync(reportUrl, ct)
                        .ConfigureAwait(false),
                    CancellationToken.None).ConfigureAwait(false))
                {
                    Logger.Info("GoPhish notification result: HTTP {0}", (int)response.StatusCode);
                    return response.IsSuccessStatusCode
                        ? GoPhishResult.Reported
                        : GoPhishResult.Error;
                }
            }
            catch (TimeoutRejectedException)
            {
                Logger.Warn("GoPhish notification timed out after all retry attempts");
                return GoPhishResult.Error;
            }
            catch (HttpRequestException ex)
            {
                Logger.Error(ex, "GoPhish notification failed after all retry attempts");
                return GoPhishResult.Error;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "GoPhish notification unexpected error");
                return GoPhishResult.Error;
            }
        }
    }
}
