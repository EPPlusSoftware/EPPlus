using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Drawing;
using System.IO;
using OfficeOpenXml.ConditionalFormatting;
using System.Data;
using FakeItEasy;
using System.Security.AccessControl;
namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableConditionalFormattingReadTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenTemplatePackage("PivotTableCF.xlsx");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;
            SaveWorkbook("PivotTableCFRead.xlsx", _pck);
            _pck.Dispose();
        }
        [TestMethod]
        public void ReadPivotTableCF()
        {
            var ws = _pck.Workbook.Worksheets["pt1"];
            var pt = ws.PivotTables[0];
            Assert.AreEqual(3, pt.ConditionalFormattings.Count);
            var cf1= pt.ConditionalFormattings[0];
            pt.CacheDefinition.Refresh();
            pt.Calculate();            
            Assert.AreEqual(1, cf1.Areas.Count);
            Assert.AreEqual(1, cf1.Areas[0].Conditions.DataFields.Count);
            Assert.AreEqual(2, cf1.Areas[0].Conditions.Fields.Count);
            
            Assert.AreEqual(6, cf1.Areas[0].Conditions.Fields[0].Items.Count);
            Assert.AreEqual(15, cf1.Areas[0].Conditions.Fields[0].Items[0].Index);
            Assert.AreEqual(2, cf1.Areas[0].Conditions.Fields[1].Items.Count);
        }
        [TestMethod]
        public void DeleteSourceDataShouldRemoveCF()
        {
            using (var pck = OpenTemplatePackage("PivotTableCF.xlsx"))
            {
                var ws = pck.Workbook.Worksheets["pt1"];
                pck.Workbook.Worksheets["Data"].Tables[0].DeleteRow(10, 180);
                var pt = ws.PivotTables[0];
                Assert.AreEqual(3, pt.ConditionalFormattings.Count);
                var cf1 = pt.ConditionalFormattings[0];
                pt.CacheDefinition.Refresh();
                pt.Calculate();

                SaveWorkbook("PivotTableDeleteCF.xlsx", pck);
            }
        }
    }
}

