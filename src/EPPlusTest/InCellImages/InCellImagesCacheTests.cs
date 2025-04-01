using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using EPPlusTest.Properties;

namespace EPPlusTest.InCellImages
{
    [TestClass]
    public class InCellImagesCacheTests : TestBase
    {
        [TestMethod]
        public void ShouldReusePicture()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["A2"].Picture.Set(Resources.Png2ByteArray);
            var vm1 = sheet._metadataStore.GetValue(1, 1);
            var vm2 = sheet._metadataStore.GetValue(2, 1);
            Assert.AreEqual(vm1, vm2);
            SaveWorkbook("InCellPicturesReuseCache1.xlsx", package);
        }

        [TestMethod]
        public void ShouldNotReuseWhenDifferentPictures()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["A2"].Picture.Set(Resources.Png3ByteArray);
            var vm1 = sheet._metadataStore.GetValue(1, 1);
            var vm2 = sheet._metadataStore.GetValue(2, 1);
            Assert.AreNotEqual(vm1, vm2);
            SaveWorkbook("InCellPicturesReuseCache2.xlsx", package);
        }

        [TestMethod]
        public void ShouldNotRemoveRichDataWhenMoreReferencesExists()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["A2"].Picture.Set(Resources.Png2ByteArray);
            var vm1 = sheet._metadataStore.GetValue(1, 1);
            var vm2 = sheet._metadataStore.GetValue(2, 1);
            Assert.AreEqual(vm1, vm2);
            Assert.AreEqual(1, package.Workbook.RichData.Db.Values.Count());
            sheet.Cells["A1"].Picture.Remove();
            Assert.IsNull(sheet.Cells["A1"].Value);
            var vm3 = sheet._metadataStore.GetValue(1, 1);
            Assert.AreEqual(0u, vm3.vm);
            Assert.AreEqual(1, package.Workbook.RichData.Db.Values.Count());
            SaveWorkbook("InCellPicturesReuseCache3.xlsx", package);
        }

        [TestMethod]
        public void ShouldRemoveRichDataWhenLastReferenceRemoved()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["A2"].Picture.Set(Resources.Png2ByteArray);
            var vm1 = sheet._metadataStore.GetValue(1, 1);
            var vm2 = sheet._metadataStore.GetValue(2, 1);
            Assert.AreEqual(vm1, vm2);
            Assert.AreEqual(1, package.Workbook.RichData.Db.Values.Count());
            sheet.Cells["A1"].Picture.Remove();
            sheet.Cells["A2"].Picture.Remove();
            Assert.IsNull(sheet.Cells["A1"].Value);
            Assert.IsNull(sheet.Cells["A2"].Value);
            var vm3 = sheet._metadataStore.GetValue(1, 1);
            Assert.AreEqual(0u, vm3.vm);
            var vm4 = sheet._metadataStore.GetValue(2, 1);
            Assert.AreEqual(0u, vm4.vm);
            Assert.AreEqual(0, package.Workbook.RichData.Db.Values.Count, "RichDataValue still exists");
            SaveWorkbook("InCellPicturesReuseCache4.xlsx", package);
        }
    }
}
