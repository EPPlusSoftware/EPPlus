using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Sorting;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace EPPlusTest.Sorting
{
    [TestClass]
    public class SortCaseSensitivityTests : TestBase
    {
        [TestMethod]
        public void ShouldRespectCaseSensitivity()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("SortTest");
                
                // Data: APPLE then apple
                sheet.Cells["A1"].Value = "APPLE";
                sheet.Cells["A2"].Value = "apple";

                // 1. Case-Insensitive Sort (Should preserve original order APPLE, apple)
                var options = RangeSortOptions.Create();
                options.CompareOptions = CompareOptions.IgnoreCase;
                options.SortBy.Column(0);
                sheet.Cells["A1:A2"].Sort(options);

                Assert.AreEqual("APPLE", sheet.Cells["A1"].Text, "Insensitive sort failed to preserve order");
                Assert.AreEqual("apple", sheet.Cells["A2"].Text);

                // 2. Case-Sensitive Sort (Should flip to apple, APPLE because a < A in linguistic sort)
                options = RangeSortOptions.Create();
                options.CompareOptions = CompareOptions.None; // linguistic sensitive
                options.SortBy.Column(0);
                sheet.Cells["A1:A2"].Sort(options);

                Assert.AreEqual("apple", sheet.Cells["A1"].Text, "Sensitive sort failed to flip order");
                Assert.AreEqual("APPLE", sheet.Cells["A2"].Text);
            }
        }

        [TestMethod]
        public void ShouldRespectOrdinalCaseSensitivity()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("SortTestOrdinal");
                
                // Data: apple then APPLE
                sheet.Cells["A1"].Value = "apple";
                sheet.Cells["A2"].Value = "APPLE";

                // Ordinal Sort (Binary): A (65) < a (97). So APPLE should be first.
                var options = RangeSortOptions.Create();
                options.CompareOptions = CompareOptions.Ordinal;
                options.SortBy.Column(0);
                sheet.Cells["A1:A2"].Sort(options);

                Assert.AreEqual("APPLE", sheet.Cells["A1"].Text, "Ordinal sort failed to put uppercase first");
                Assert.AreEqual("apple", sheet.Cells["A2"].Text);
            }
        }
    }
}
