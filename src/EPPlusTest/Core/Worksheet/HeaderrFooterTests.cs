using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class HeaderrFooterTests
    {
        [TestMethod]
        public void Issue_CopyWorksheet_HeaderFooter_Throws()
        {
            // Direct C# translation of the reporter's "Reproduction — throws" script.
            // Source is any workbook that has a header/footer picture.
            byte[] sourceBytes = CreateWorkbookWithHeaderFooterPicture();

            using (var sourceStream = new MemoryStream(sourceBytes))
            using (var source = new ExcelPackage(sourceStream))
            {
                var sourceSheet = source.Workbook.Worksheets[0];

                // The trigger: read a HeaderFooter property on the source first.
                var trigger = sourceSheet.HeaderFooter.OddFooter.CenteredText;

                using (var target = new ExcelPackage())
                {
                    // Before the fix: throws NullReferenceException from
                    // WorksheetCopyHelper.CopyHeaderFooterPictures.
                    target.Workbook.Worksheets.Add("Copied", sourceSheet);
                }
            }
        }

        [TestMethod]
        public void Issue_CopyWorksheet_HeaderFooter_SilentDataLoss()
        {
            byte[] sourceBytes = CreateWorkbookWithHeaderFooterPicture();

            byte[] outputBytes;
            using (var sourceStream = new MemoryStream(sourceBytes))
            using (var source = new ExcelPackage(sourceStream))
            {
                var sourceSheet = source.Workbook.Worksheets[0];

                // No HeaderFooter read here.

                using (var target = new ExcelPackage())
                {
                    var copiedSheet = target.Workbook.Worksheets.Add("Copied", sourceSheet);

                    Assert.AreEqual(1, copiedSheet.HeaderFooter.Pictures.Count,
                        "Header/footer picture missing right after copy.");

                    outputBytes = target.GetAsByteArray();
                }
            }

            // After save + reopen.
            using (var reopenStream = new MemoryStream(outputBytes))
            using (var reopen = new ExcelPackage(reopenStream))
            {
                var reopenedSheet = reopen.Workbook.Worksheets["Copied"];
                Assert.AreEqual(1, reopenedSheet.HeaderFooter.Pictures.Count,
                    "Header/footer picture was lost after save and reopen.");
            }
        }

        private static byte[] CreateWorkbookWithHeaderFooterPicture()
        {
            // Stands in for the reporter's "/path/to/any-workbook-with-a-footer.xlsx".
            using (var source = new ExcelPackage())
            {
                var ws = source.Workbook.Worksheets.Add("Sheet1");
                ws.HeaderFooter.OddFooter.CenteredText = "MyFooter";
                ws.HeaderFooter.OddFooter.InsertPicture(Properties.Resources.Test1, PictureAlignment.Centered);
                return source.GetAsByteArray();
            }
        }
    }
}
