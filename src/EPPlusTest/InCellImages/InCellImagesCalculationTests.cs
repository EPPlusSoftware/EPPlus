using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using EPPlusTest.Properties;
using OfficeOpenXml.CellPictures;
using OfficeOpenXml.Metadata.FutureMetadata;

namespace EPPlusTest.InCellImages
{
    [TestClass]
    [DoNotParallelize]
    public class InCellImagesCalculationTests : TestBase
    {
        [TestMethod]
        public void ShouldAddRichDataWhenReferenceFromOtherCell()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["B1"].Formula = "A1";
            sheet.Calculate();
            var rd = package.Workbook.RichData;
            Assert.AreEqual(2, rd.Db.Values.Count, "RichData values count was not 2");
            var cellVal = sheet.Cells["B1"].Value;
            Assert.IsNotNull(cellVal, "B1 cell value was null");
            Assert.IsInstanceOfType(cellVal, typeof(ExcelCellPicture), "Cell B1 did not contain an instance of ExcelCellPicture");
            var b1Pic = cellVal as ExcelCellPicture;
            Assert.AreEqual(CalcOrigins.Reference, b1Pic.CalcOrigin, "CalcOrigin of the B1 picture was not Reference");
            var b1Bytes = b1Pic.GetImageBytes();
            Assert.AreEqual(Resources.Png2ByteArray.Length, b1Bytes.Length, "Image bytes of the cell picture was not the same as the added picture");
            Assert.AreEqual(1, package.Workbook.RichData.Db.RichValueRels.Count);
            SaveWorkbook("InCellImageCalculate1.xlsx", package);
        }

        [TestMethod]
        public void ShouldCauseValueErrorInSum()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Formula = "SUM(A1:A2)";
            sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A3"].Value);
        }

        [TestMethod]
        public void ShouldAddPictureInIfFunction()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Formula = "IF(TRUE(),A1,A2)";
            sheet.Calculate();
            Assert.IsTrue(sheet.Cells["A3"].Picture.Exists, "No picture present in cell A3 after calc");
            var pic = sheet.Cells["A3"].Picture.Get();
            Assert.AreEqual(Resources.Png2ByteArray.Length, pic.GetImageBytes().Length, "Length of A3 image bytes was not the same as the original picture");
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A3"].Value);
            SaveWorkbook("InCellImageCalculate2.xlsx", package);
        }

        [TestMethod]
        public void ShouldRemoveCalculatedImage1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["A3"].Formula = "IF(TRUE(),A1,A2)";
            sheet.Calculate();
            Assert.IsTrue(sheet.Cells["A3"].Picture.Exists, "No picture present in cell A3 after calc");
            Assert.AreEqual(2, package.Workbook.RichData.Db.Values.Count, $"RichData.Values.Count was {package.Workbook.RichData.Db.Values.Count}, not 2 as expected");
            var pic = sheet.Cells["A3"].Picture.Get();
            Assert.AreEqual(Resources.Png2ByteArray.Length, pic.GetImageBytes().Length, "Length of A3 image bytes was not the same as the original picture");
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A3"].Value);
            Assert.AreEqual(2, package.Workbook.RichData.Db.Values.Count);
            sheet.Cells["A3"].Formula = "IF(FALSE(),A1,A2)";
            sheet.Calculate();
            Assert.AreEqual(1, sheet.Cells["A3"].Value, $"Value of A3 was {sheet.Cells["A3"].Value}, not 1 as expected");
            Assert.AreEqual(1, package.Workbook.RichData.Db.Values.Count, $"RichData.Values.Count was {package.Workbook.RichData.Db.Values.Count} after re-calc, not 1 as expected");
            Assert.IsFalse(sheet.Cells["A3"].Picture.Exists);
            SaveWorkbook("InCellImageCalculate3.xlsx", package);
        }

        [TestMethod]
        public void ShouldOverwriteSpillError()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].Formula = "RANDARRAY(3,3)";
            sheet.Cells["C3"].Value = 1;
            sheet.Calculate();
            Assert.AreEqual("#SPILL!", sheet.Cells["A1"].Value.ToString());
            Assert.AreEqual(0, package.Workbook.RichData.Db.Values.Count);
            sheet.Cells["A1"].Picture.Set(Resources.Png2ByteArray);
            Assert.IsInstanceOfType(sheet.Cells["A1"].Value, typeof(ExcelCellPicture));
            Assert.IsTrue(sheet._metadataStore.GetValue(1, 1).vm > 0);
            Assert.AreEqual(1, package.Workbook.RichData.Db.Values.Count);
        }

        [TestMethod]
        public void ShouldOverwriteSpillErrorAndPreserveRichData()
        {
            using var package = OpenTemplatePackage("ExistingRichData1.xlsx");
            var sheet = package.Workbook.Worksheets.First();
            var mdrA1 = sheet._metadataStore.GetValue(1, 1);
            Assert.IsTrue(mdrA1.vm > 0, "Before: Value metadata was not set on cell A1");
            var mdrA6 = sheet._metadataStore.GetValue(6, 1);
            Assert.IsTrue(mdrA6.vm > 0, "Before: Value metadata was not set on cell A6");
            Assert.AreEqual(2, package.Workbook.Metadata.Db.ValueMetadata.Count, "Before: ValueMetadata.Count was not 2 as expected");
            Assert.AreEqual(2, package.Workbook.Metadata.Db.MetadataTypes.Count, "Before: MetadataTypes.Count was not 2 as expected");
            Assert.AreEqual("XLDAPR", package.Workbook.Metadata.Db.MetadataTypes.First().Name, "Before: First metadata type was XLDAPR not as expected");
            Assert.AreEqual("XLRICHVALUE", package.Workbook.Metadata.Db.MetadataTypes.Last().Name, "Before: Last metadata type was not XLRICHVALUE as expected");

            // now overwrite the existing spill error with a cell image
            sheet.Cells["B1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["A1"].Formula = "IF(TRUE(), B1, B2)";
            sheet.Calculate();
            sheet.Workbook.IndexStore.PrintRelations(@"c:\Temp");
            var pic = sheet.Cells["A1"].Picture.Get();
            Assert.IsNotNull(pic);
            // we should now have 3 metadata records, the pictures in cell A1 and B2 and the geography in cell A4
            Assert.AreEqual(3, package.Workbook.Metadata.Db.ValueMetadata.Count, "After: ValueMetadata.Count was not 3 as expected");
            // The XLDAPR metadata type should be deleted since the dynamic array error in A1 is overwritten by the picture
            Assert.AreEqual(1, package.Workbook.Metadata.Db.MetadataTypes.Count, "After: ValueMetadataTypes.Count was not 1 as expected");
            // The only remaining metadata type should be rich data.
            Assert.AreEqual(FutureMetadataBase.RICHDATA_NAME, package.Workbook.Metadata.Db.MetadataTypes.First(x => !x.Deleted).Name);
            // There should be only one FutureMetadata instance (for rich data) left
            Assert.AreEqual(1, package.Workbook.Metadata.Db.FutureMetadata.Count, "After: FutureMetadata.Count was not 0 as expected.");
            // There should be no cell value metadata blocks left
            Assert.AreEqual(0, package.Workbook.Metadata.Db.CellMetadata.Count, "After: CellMetadata.Count was not 0 as expected.");
            // There should be no cell metadata records left
            Assert.AreEqual(0, package.Workbook.Metadata.Db.CellMetadataRecords.Count, "After: CellMetadataRecords.Count was not 0 as expected.");

            SaveWorkbook("ExistingRichData1_Result.xlsx", package);

        }

        [TestMethod]
        public void ShouldOverwriteSpillErrorAndPreserveRichData2()
        {
            using var package = OpenTemplatePackage("ExistingRichData2.xlsx");
            var sheet = package.Workbook.Worksheets.First();
            // now overwrite the existing spill error with a cell image
            sheet.Cells["B1"].Picture.Set(Resources.Png2ByteArray);
            sheet.Cells["A1"].Formula = "IF(TRUE(), B1, B2)";
            sheet.Calculate();
            sheet.Workbook.IndexStore.PrintRelations(@"c:\Temp");
            var pic = sheet.Cells["A1"].Picture.Get();
            Assert.IsNotNull(pic);

            SaveWorkbook("ExistingRichData2_Result.xlsx", package);

        }

        [TestMethod]
        public void ExistingRichData3()
        {
            using var package = OpenTemplatePackage("ExistingRichData1.xlsx");
            var sheet = package.Workbook.Worksheets.First();
            sheet.Cells["A1"].Value = 1;
            SaveWorkbook("ExistingRichData3_Result.xlsx", package);
        }
    }
}
