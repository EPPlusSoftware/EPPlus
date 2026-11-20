using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class VbaIssues : TestBase
    {
        [ClassInitialize]
        public static void Init(TestContext context)
        {
        }
        [ClassCleanup]
        public static void Cleanup()
        {
        }
        [TestInitialize]
        public void Initialize()
        {
        }
        [TestMethod]
        public void s1096()
        {
            using (var p = OpenPackage("s1096.xlsm", true))
            {

                p.Workbook.Worksheets.Add("Sheet1");

                p.Workbook.CreateVBAProject();

                p.Workbook.CodeModule.Code =

                  "Private Sub Workbook_Open()\r\n  MsgBox \"Signed macro ran\"\r\nEnd Sub\r\n";

                //p.Workbook.VbaProject.Signature.Certificate = GetSelfCert(); // X509Certificate2 with private key
                p.Workbook.VbaProject.Signature.Certificate = LoadSelfCert(); // X509Certificate2 with private key
                p.Workbook.VbaProject.Signature.V3Signature.HashAlgorithm = OfficeOpenXml.VBA.VbaSignatureHashAlgorithm.SHA256;
                var fi = GetOutputFile("", "s1096.xlsm");
                p.SaveAs(fi);
             }
        }
        /// <summary>
        /// Use to creat a self-signed certificate for testing purposes. The certificate will be valid for 1 day.
        /// </summary>
        /// <returns></returns>
        X509Certificate2 GetSelfCert()
        {
            var requestedCert = new CertificateRequest("cn=SelfSignCert", RSA.Create(3072), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            var finalCert = requestedCert.CreateSelfSigned(DateTimeOffset.Now, DateTimeOffset.Now.AddDays(1));

            var certPrivate = finalCert.Export(X509ContentType.Pfx);
            var certPublic = finalCert.Export(X509ContentType.Cert);
            var newCert = new X509Certificate2(certPrivate, "", X509KeyStorageFlags.Exportable);

            byte[] pfx = newCert.Export(X509ContentType.Pfx);
            //File.WriteAllBytes($"c:\\temp\\mycert-saved.pfx", pfx);
            return newCert;
        }
        X509Certificate2 LoadSelfCert()
        {
            return new X509Certificate2("c:\\temp\\mycert-saved.pfx");
        }
    }
}
