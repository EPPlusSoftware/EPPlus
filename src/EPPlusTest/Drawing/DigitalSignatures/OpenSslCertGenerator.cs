using System;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography.X509Certificates;

namespace EPPlusTest.Drawing.DigitalSignatures
{
    public static class OpenSslCertGenerator
    {
        public static X509Certificate2 CreateSelfSignedCert(string commonName = "SelfSignCert", int days = 1)
        {
            // Create unique temporary file paths in /tmp
            string tempDir = Path.GetTempPath();
            string id = Guid.NewGuid().ToString("N");
            string keyPath = Path.Combine(tempDir, $"{id}.key");
            string crtPath = Path.Combine(tempDir, $"{id}.crt");
            string pfxPath = Path.Combine(tempDir, $"{id}.pfx");

            try
            {
                // 1. Generate the private key and self-signed certificate
                RunOpenSsl($"req -x509 -newkey rsa:2048 -keyout \"{keyPath}\" -out \"{crtPath}\" -days {days} -nodes -subj \"/CN={commonName}\"");

                // 2. Bundle the key and certificate into a passwordless .pfx file
                RunOpenSsl($"pkcs12 -export -legacy -out \"{pfxPath}\" -inkey \"{keyPath}\" -in \"{crtPath}\" -passout pass:");

                // 3. Load the .pfx bytes into a standard .NET X509Certificate2
                byte[] pfxBytes = File.ReadAllBytes(pfxPath);
                return new X509Certificate2(pfxBytes, string.Empty, X509KeyStorageFlags.Exportable);
            }
            finally
            {
                // 4. Always delete temporary files from disk
                DeleteIfExists(keyPath);
                DeleteIfExists(crtPath);
                DeleteIfExists(pfxPath);
            }
        }

        private static void RunOpenSsl(string arguments)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "openssl",
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (var process = Process.Start(startInfo))
            {
                string errorOutput = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (process.ExitCode != 0)
                {
                    throw new InvalidOperationException($"OpenSSL failed (exit code {process.ExitCode}): {errorOutput}");
                }
            }
        }

        private static void DeleteIfExists(string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }
}
