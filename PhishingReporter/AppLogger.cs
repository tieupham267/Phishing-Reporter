using System;
using NLog;
using NLog.Config;

namespace PhishingReporter
{
    /// <summary>
    /// Provides an isolated NLog LogFactory for the VSTO add-in.
    /// Uses a dedicated LogFactory instead of the global LogManager to avoid
    /// config conflicts with other add-ins loaded in the same Outlook process.
    /// Source: https://github.com/NLog/NLog/wiki/Configure-component-logging
    /// </summary>
    internal static class AppLogger
    {
        public static LogFactory Instance { get { return _instance.Value; } }

        private static readonly Lazy<LogFactory> _instance =
            new Lazy<LogFactory>(BuildLogFactory);

        private static LogFactory BuildLogFactory()
        {
            // Loaded from embedded resource instead of Assembly.Location, because VSTO
            // shadow-copies the DLL into AppData\Local\assembly\dl3\... without bringing
            // arbitrary content files along — a file-path-based load throws FileNotFound.
            var thisAssembly = System.Reflection.Assembly.GetExecutingAssembly();
            const string resourceName = "PhishingReporter.NLog.config";

            using (var stream = thisAssembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException(
                        "Embedded resource '" + resourceName + "' not found in PhishingReporter assembly.");
                }

                var logFactory = new LogFactory();
                using (var xmlReader = System.Xml.XmlReader.Create(stream))
                {
                    // basePath = AppDomain.CurrentDomain.BaseDirectory future-proofs any
                    // target that later uses a relative fileName. Current NLog.config only
                    // uses ${specialfolder} (absolute) renderers, so this has no effect today.
                    logFactory.Configuration = new XmlLoggingConfiguration(
                        xmlReader, AppDomain.CurrentDomain.BaseDirectory, logFactory);
                }
                return logFactory;
            }
        }
    }
}
