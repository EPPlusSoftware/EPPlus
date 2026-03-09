using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.CellPictures;
using OfficeOpenXml.RichData.RichValues.WebImages;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest.InCellImages
{
    [TestClass]
    public class WebImagesTests : TestBase
    {
        [TestMethod]
        public void LoadSimpleWorkbook1()
        {
            using var package = OpenTemplatePackage("ImageFunction1.xlsx");
            var sheet = package.Workbook.Worksheets.First();
            var webPic = sheet.Cells["A1"].Picture.Get();
            var uri = webPic.ImageUri;

            //sheet.Cells["A2"].Blip.Set(Resources.Png2ByteArray);
            var localPic = sheet.Cells["B1"].Picture.Get();
            var lpBytes = localPic.GetImageBytes();

            var imageBytes = webPic.GetImageBytes();

            SaveWorkbook("ImageFunction1_Output.xlsx", package);
        }

        [TestMethod]
        public void WebImagesShouldBeRemovedWhenOverwritten()
        {
            using var p = new ExcelPackage();
            var sheet = p.Workbook.Worksheets.Add("Sheet");
            sheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
            sheet.Cells["A2"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
            sheet.Calculate();
            Assert.IsTrue(sheet.Cells["A1"].Picture.Exists);
            Assert.IsTrue(sheet.Cells["A2"].Picture.Exists);
            Assert.AreEqual(ExcelCellPictureTypes.WebImage, sheet.Cells["A1"].Picture.Get().PictureType);
            Assert.AreEqual(ExcelCellPictureTypes.WebImage, sheet.Cells["A2"].Picture.Get().PictureType);
            Assert.AreEqual(1, p.Workbook._images.Count);
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            Assert.AreEqual(2, p.Workbook._images.Count);
            sheet.Cells["A2"].Picture.Set(Resources.Png2ByteArray);
            Assert.AreEqual(1, p.Workbook._images.Count);
            SaveWorkbook("WebImages_Removed.xlsx", p);

        }

        [TestMethod]
        public void DoublePictDelete()
        {
            using (var package = OpenTemplatePackage("DoublePictInCellWeb.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];

                //ws.Calculate();

                ws.Cells["B3"].Picture.Remove();

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void ShouldNotRemoveRichDataWhenMoreReferencesExistsWhenReading()
        {
            var ms = new MemoryStream();

            using (var p = new ExcelPackage())
            {
                var sheet = p.Workbook.Worksheets.Add("Sheet");
                sheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
                sheet.Cells["A2"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
                sheet.Calculate();
                Assert.IsTrue(sheet.Cells["A1"].Picture.Exists);
                Assert.IsTrue(sheet.Cells["A2"].Picture.Exists);
                Assert.AreEqual(ExcelCellPictureTypes.WebImage, sheet.Cells["A1"].Picture.Get().PictureType);
                Assert.AreEqual(ExcelCellPictureTypes.WebImage, sheet.Cells["A2"].Picture.Get().PictureType);
                p.SaveAs(ms);
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

                    key = new WebPictureCacheKey(pic.ExternalAddress, pic.AltText, pic.CalcOrigin, pic.Sizing ?? WebImageSizing.FitToCellMaintainRatio, null, null);

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
        public void ReCalculateWebImagesShouldWork()
        {
            using (var p = OpenPackage("ReacalculateWebImage.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet");
                ws.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
                ws.Calculate();
                Assert.IsTrue(ws.Cells["A1"].Picture.Exists);
                Assert.AreEqual(ExcelCellPictureTypes.WebImage, ws.Cells["A1"].Picture.Get().PictureType);

                var pic = ws.Cells[1, 1].Value as ExcelCellPicture;

                //Verify that the picture has added to refs correctly
                PictureCacheKey key = new WebPictureCacheKey(pic.ExternalAddress, pic.AltText, pic.CalcOrigin, pic.Sizing ?? WebImageSizing.FitToCellMaintainRatio, null, null);
                var numRefs = ws.Workbook.CellPictureReferenceCache.GetNumberOfReferences(key);
                Assert.AreEqual(1, numRefs);

                ws.Calculate();

                PictureCacheKey key2 = new WebPictureCacheKey(pic.ExternalAddress, pic.AltText, pic.CalcOrigin, pic.Sizing ?? WebImageSizing.FitToCellMaintainRatio, null, null);
                var numRefs2 = ws.Workbook.CellPictureReferenceCache.GetNumberOfReferences(key);
                Assert.AreEqual(1, numRefs2);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ShouldNotRemoveRichDataWhenMoreReferencesExistsWhenReadingWithCalculate()
        {
            var ms = new MemoryStream();

            using (var p = new ExcelPackage())
            {
                var sheet = p.Workbook.Worksheets.Add("Sheet");
                sheet.Cells["A1"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
                sheet.Cells["A2"].Formula = "IMAGE(\"https://epplussoftware.com/img/EPPlus-logo-full.png\")";
                sheet.Calculate();
                Assert.IsTrue(sheet.Cells["A1"].Picture.Exists);
                Assert.IsTrue(sheet.Cells["A2"].Picture.Exists);
                Assert.AreEqual(ExcelCellPictureTypes.WebImage, sheet.Cells["A1"].Picture.Get().PictureType);
                Assert.AreEqual(ExcelCellPictureTypes.WebImage, sheet.Cells["A2"].Picture.Get().PictureType);
                p.SaveAs(ms);
            }
            ms.Position = 0;
            ms.Seek(0, SeekOrigin.Begin);

            using (var package = new ExcelPackage(ms))
            {
                var ws = package.Workbook.Worksheets[0];

                //This causes issues as the NumberOfReferences is doubled while there's actually the same amount left
                ws.Calculate();

                var pic = ws.Cells[1, 1].Value as ExcelCellPicture;

                //Verify that the picture has been read into refs correctly
                PictureCacheKey key = null;
                if (pic != null)
                {

                    key = new WebPictureCacheKey(pic.ExternalAddress, pic.AltText, pic.CalcOrigin, pic.Sizing ?? WebImageSizing.FitToCellMaintainRatio, null, null);

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

                SaveWorkbook("InCellPicturesReuseCache3_WebImageOnRead.xlsx", package);
            }
        }
    }
}
