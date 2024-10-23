using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.DigitalSignatures;

namespace EPPlusTest.Drawing.DigitalSignatures
{
    [TestClass]
    public class DigitalSignatureLineTests : TestBase
    {
        [TestMethod]
        public void CreateEmptySignatureLine()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine_Empty.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.AddSignatureLine();

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void CreateTwoEmptySignatureLine()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine_Empty2.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.AddSignatureLine();

                var sLine2 = ws.AddSignatureLine();

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void CreateSignatureLineWithSuggestedSigner()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine_SSigner.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.AddSignatureLine();
                sLine.Signer = "ASuggestedSigner";

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void CreateSignatureLineWithSuggestedSignerAndTitle()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine_SSignerTitle.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.AddSignatureLine();
                sLine.Signer = "ASuggestedSigner";
                sLine.Title = "ASuggestedTitle";

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void CreateSignatureLineWithALL()
        {
            using (ExcelPackage package = OpenPackage("DigSig_SignatureLine_ALL.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = package.Workbook.Worksheets.Add("SignatureLineWs");

                var sLine = ws.AddSignatureLine();
                sLine.Signer = "ASuggestedSigner";
                sLine.Title = "ASuggestedTitle";
                sLine.Email = "Example@Site.com";
                sLine.SigningInstructions = "Hey please sign this because x and y so it will be z";
                sLine.AllowComments = true;
                sLine.ShowSignDate = true;

                SaveAndCleanup(package);
            }
        }
    }
}
