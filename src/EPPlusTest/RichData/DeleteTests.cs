﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.RichData.RichValues;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.RichValues.LocalImage;
using OfficeOpenXml.Metadata.FutureMetadata;
using OfficeOpenXml.RichData.Structures;

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
            var relRv = fmbk.GetFirstOutgoingRelByType<ExcelRichValue>();
            Assert.AreEqual(rv.Id, relRv.Id);

            Assert.IsFalse(rv.Deleted);
            Assert.IsFalse(fmbk.Deleted);

            rv.DeleteMe();

            Assert.IsTrue(rv.Deleted);
            Assert.IsTrue(fmbk.Deleted);
        }

        [TestMethod]
        public void RichDataStructureShouldBeDeletedWithRichValue_LastRef()
        {
            using var package = new ExcelPackage();
            var richData = package.Workbook.RichData;
            var rv = new LocalImageRichValue(package.Workbook);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var store = new RichDataStore(sheet);
            store.AddRichData(1, 1, rv);
            var structure = richData.Structures.First();
            var relStructure = rv.GetFirstOutgoingRelByType<ExcelRichValueStructure>();
            Assert.AreEqual(structure.Id, relStructure.Id);

            Assert.IsFalse(rv.Deleted);
            Assert.IsFalse(structure.Deleted);

            rv.DeleteMe();

            Assert.IsTrue(rv.Deleted);
            Assert.IsTrue(structure.Deleted);
        }

        [TestMethod]
        public void RichDataStructureShouldNotBeDeletedWithRichValue_NotLastRef()
        {
            using var package = new ExcelPackage();
            var richData = package.Workbook.RichData;
            var rv1 = new LocalImageRichValue(package.Workbook);
            var rv2 = new LocalImageRichValue(package.Workbook);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var store = new RichDataStore(sheet);
            store.AddRichData(1, 1, rv1);
            store.AddRichData(1, 2, rv2);
            var structure = richData.Structures.First();
            var relStructure = rv1.GetFirstOutgoingRelByType<ExcelRichValueStructure>();
            Assert.AreEqual(structure.Id, relStructure.Id);
            Assert.AreEqual(1, richData.Structures.Count);

            Assert.IsFalse(rv1.Deleted);
            Assert.IsFalse(structure.Deleted);

            rv1.DeleteMe();

            Assert.IsTrue(rv1.Deleted);
            Assert.IsFalse(rv2.Deleted);
            Assert.IsFalse(structure.Deleted);
        }

        [TestMethod]
        public void FutureMetadataTypeShouldBeDeletedWithFutureMetadataBlock_LastRef()
        {
            using var package = new ExcelPackage();
            var metadata = package.Workbook.Metadata;
            var richData = package.Workbook.RichData;
            var rv = new LocalImageRichValue(package.Workbook);
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var store = new RichDataStore(sheet);
            Assert.AreEqual(0, metadata.MetadataTypes.Count);
            store.AddRichData(1, 1, rv);
            Assert.AreEqual(1, metadata.MetadataTypes.Count);
            var type = metadata.MetadataTypes[0];
            var bk = metadata.FutureMetadataBlocks.First();
            var relBk = rv.GetFirstIncomingRelByType<FutureMetadataBlock>();
            Assert.AreEqual(bk.Id, relBk.Id);

            Assert.IsFalse(bk.Deleted);
            Assert.IsFalse(rv.Deleted);
            Assert.IsFalse(type.Deleted);

            rv.DeleteMe();

            Assert.IsTrue(rv.Deleted);
            Assert.IsTrue(bk.Deleted);
            Assert.IsTrue(type.Deleted);
        }
    }
}
