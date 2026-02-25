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
            var thisAssembly = System.Reflection.Assembly.GetExecutingAssembly();
            var configDir = System.IO.Path.GetDirectoryName(thisAssembly.Location);
            var configFilePath = System.IO.Path.Combine(configDir, "NLog.config");

            var logFactory = new LogFactory();
            logFactory.Configuration = new XmlLoggingConfiguration(configFilePath, logFactory);
            return logFactory;
        }
    }
}
