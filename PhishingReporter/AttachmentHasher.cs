using System;
using System.IO;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Outlook;

namespace PhishingReporter
{
    /// <summary>
    /// Hash results for a single attachment. Immutable.
    /// </summary>
    internal sealed class AttachmentHashResult
    {
        public string FileName { get; }
        public int SizeBytes { get; }
        public string Md5 { get; }
        public string Sha256 { get; }

        public AttachmentHashResult(string fileName, int sizeBytes, string md5, string sha256)
        {
            FileName = fileName;
            SizeBytes = sizeBytes;
            Md5 = md5;
            Sha256 = sha256;
        }
    }

    /// <summary>
    /// Computes MD5 and SHA256 hashes of Outlook attachment content.
    /// Uses temp files with guaranteed cleanup in a finally block (BUGF-05 fix).
    /// </summary>
    internal static class AttachmentHasher
    {
        private static readonly NLog.Logger Logger =
            AppLogger.Instance.GetCurrentClassLogger();

        /// <summary>
        /// Saves attachment to a temp file, computes hashes, and cleans up.
        /// The temp file is ALWAYS deleted, even if hashing throws.
        /// </summary>
        public static AttachmentHashResult ComputeHashes(Attachment attachment)
        {
            // Use GUID instead of DisplayName to avoid path-illegal characters and collisions
            string tempPath = Path.Combine(
                Path.GetTempPath(),
                "Outlook-Phishaddin-" + Guid.NewGuid().ToString("N") + ".tmp");

            try
            {
                attachment.SaveAsFile(tempPath);

                string md5Hash = ComputeMd5(tempPath);
                string sha256Hash = ComputeSha256(tempPath);

                Logger.Debug("Computed hashes for attachment: {0}", attachment.FileName);

                return new AttachmentHashResult(
                    attachment.FileName,
                    attachment.Size,
                    md5Hash,
                    sha256Hash);
            }
            finally
            {
                // BUGF-05: Guaranteed cleanup regardless of exceptions
                try
                {
                    if (File.Exists(tempPath))
                    {
                        File.Delete(tempPath);
                    }
                }
                catch (IOException ex)
                {
                    Logger.Warn(ex, "Failed to delete temp file: {0}", tempPath);
                }
            }
        }

        private static string ComputeMd5(string filePath)
        {
            using (var md5 = MD5.Create())
            using (var stream = File.OpenRead(filePath))
            {
                byte[] hash = md5.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }
        }

        private static string ComputeSha256(string filePath)
        {
            using (var sha256 = SHA256.Create())
            using (var stream = File.OpenRead(filePath))
            {
                byte[] hash = sha256.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }
        }
    }
}
