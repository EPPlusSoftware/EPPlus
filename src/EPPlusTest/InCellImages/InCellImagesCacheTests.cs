using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.CellPictures;

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
        public void ShouldNotRemoveRichDataWhenMoreReferencesExistsWhenReading()
        {
            var ms = new MemoryStream();

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet 1");
                sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
                sheet.Cells["A2"].Picture.Set(Resources.Png2ByteArray);
                var vm1 = sheet._metadataStore.GetValue(1, 1);
                var vm2 = sheet._metadataStore.GetValue(2, 1);
                Assert.AreEqual(vm1, vm2);
                Assert.AreEqual(1, package.Workbook.RichData.Db.Values.Count());
                package.SaveAs(ms);
            }
            ms.Position = 0;
            ms.Seek(0, SeekOrigin.Begin);

            using (var package = new ExcelPackage(ms))
            {
                var ws = package.Workbook.Worksheets[0];

                var pic = ws.Cells[1, 1].Value as ExcelCellPicture;

                //Verify that the picture has been read into refs correctly
                PictureCacheKey key = null;
                if (pic != null)
                {
                    if (pic.PictureType == ExcelCellPictureTypes.LocalImage)
                    {
                        key = new LocalImageCacheKey(pic.ImageUri, pic.CalcOrigin, pic.AltText);
                    }
                    Assert.IsTrue(ws.Workbook.CellPictureReferenceCache.Contains(key));
                    var numberReferencesLeft = ws.Workbook.CellPictureReferenceCache.GetNumberOfReferences(key);
                    Assert.AreEqual(2, numberReferencesLeft);
                }

                ws.Cells["A1"].Picture.Remove();
                Assert.IsNull(ws.Cells["A1"].Value);
                var vm3 = ws._metadataStore.GetValue(1, 1);
                Assert.AreEqual(0u, vm3.vm);
                Assert.AreEqual(1, package.Workbook.RichData.Db.Values.Count());

                //Verify that the picture ref has been removed
                Assert.IsTrue(ws.Workbook.CellPictureReferenceCache.Contains(key));
                var numberOfRefs = ws.Workbook.CellPictureReferenceCache.GetNumberOfReferences(key);
                Assert.AreEqual(1, numberOfRefs);

                SaveWorkbook("InCellPicturesReuseCache3_OnRead.xlsx", package);
            }
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
