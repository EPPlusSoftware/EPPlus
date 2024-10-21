using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.RichData.RichValues;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.RichValues.LocalImage;
using OfficeOpenXml.Metadata.FutureMetadata;

namespace EPPlusTest.RichData
{
    [TestClass]
    public class DeleteTests
    {
        [TestMethod]
        public void FutureMetadataBlockShouldBeDeletedWithRichValue()
        {
            using var package = new ExcelPackage();
            var metadata = package.Workbook.Metadata;
            var richData = package.Workbook.RichData;
            var rv = new LocalImageRichValue(package.Workbook);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var store = new RichDataStore(sheet);
            store.AddRichData(1, 1, rv);
            var fmbk = metadata.FutureMetadataRichValue.Blocks.First();
            var relRv = fmbk.GetFirstTargetByType<ExcelRichValue>();
            Assert.AreEqual(rv.Id, relRv.Id);

            Assert.IsFalse(rv.Deleted);
            Assert.IsFalse(fmbk.Deleted);

            rv.DeleteMe();

            Assert.IsTrue(rv.Deleted);
            Assert.IsTrue(fmbk.Deleted);
        }
    }
}
