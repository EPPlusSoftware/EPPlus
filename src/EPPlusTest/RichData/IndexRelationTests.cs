using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Metadata;
using OfficeOpenXml.Metadata.FutureMetadata;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;

namespace EPPlusTest.RichData
{
    [TestClass]
    public class IndexRelationTests
    {
        [TestMethod]
        public void ValidateRelationsLocalImage()
        {
            var pic1Bytes = Resources.Png2ByteArray;
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1"].SetCellPicture(pic1Bytes);

            // From ValueMetadataBlock
            var vmb = package.Workbook.Metadata.ValueMetadata.First();
            Assert.IsTrue(vmb.HasOutgoingRelationTo(RichDataEntities.MetadataType), "No relation from value metadata block to metadata type");
            Assert.IsTrue(vmb.HasOutgoingRelationTo(RichDataEntities.FutureMetadataRichDataBlock), "No relation from value metadatablock to future metadata block");

            // From metadata type
            var hasRichDataType = package.Workbook.Metadata.MetadataTypes.TryGetValue(FutureMetadataBase.RICHDATA_NAME, out ExcelMetadataType metadataType);
            Assert.IsTrue(hasRichDataType, "No existing metadata type for Rich Values");
            Assert.IsTrue(metadataType.HasOutgoingRelationTo(RichDataEntities.FutureMetadata), "No relation from metadata type to futuremetadata for richvalues");

            // From FutureMetadata block
            var fmb = package.Workbook.Metadata.FutureMetadataBlocks.First();
            Assert.IsTrue(fmb.HasOutgoingRelationTo(RichDataEntities.RichValue), "No relation from future metadata block to rich value");

            // From Rich Value
            var rv = package.Workbook.RichData.Values.First();
            Assert.IsTrue(rv.HasOutgoingRelationTo(RichDataEntities.RichStructure), "No relation from rich value to rich structure");
            Assert.IsTrue(rv.HasOutgoingRelationTo(RichDataEntities.RichValueRel), "No relation from rich value to rich value rel");


        }
    }
}
