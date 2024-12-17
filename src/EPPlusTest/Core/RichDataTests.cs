using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core
{
    [TestClass]
    public class RichDataTests : TestBase
    {
        [ClassInitialize]
        public static void Init(TestContext context)
        {
        }

        [TestMethod]
        public void RichDataReadTest()
        {
            using (var p = OpenTemplatePackage("RichData.xlsx"))
            {
                Assert.AreEqual(10, p.Workbook.RichData.Db.ValueTypes.Global.Count);
                Assert.AreEqual(3, p.Workbook.RichData.Db.Structures.Count);
                Assert.AreEqual(4, p.Workbook.RichData.Db.Values.Count);

                Assert.AreEqual(2, p.Workbook.Metadata.Db.MetadataTypes.Count);
                //Assert.AreEqual(1, p.Workbook.Metadata.FutureMetadata["XLDAPR"].Types.Count);
                //Assert.AreEqual(4, p.Workbook.Metadata.FutureMetadata["XLRICHVALUE"].Types.Count);
                Assert.AreEqual(1, p.Workbook.Metadata.Db.CellMetadata.Count);
                Assert.AreEqual(4, p.Workbook.Metadata.Db.ValueMetadata.Count);


                var ws = p.Workbook.Worksheets[0];

                Assert.IsInstanceOfType(ws.Cells["B1"].Value, typeof(ExcelErrorValue));
                Assert.AreEqual(((ExcelErrorValue)ws.Cells["B1"].Value).Type, eErrorType.Spill);

                Assert.IsInstanceOfType(ws.Cells["C1"].Value, typeof(ExcelErrorValue));
                Assert.AreEqual(((ExcelErrorValue)ws.Cells["C1"].Value).Type, eErrorType.Calc);

                Assert.IsInstanceOfType(ws.Cells["F1"].Value, typeof(ExcelErrorValue));
                Assert.AreEqual(((ExcelErrorValue)ws.Cells["F1"].Value).Type, eErrorType.Spill);

                Assert.IsInstanceOfType(ws.Cells["E10"].Value, typeof(ExcelErrorValue));
                Assert.AreEqual(((ExcelErrorValue)ws.Cells["E10"].Value).Type, eErrorType.Spill);

                SaveAndCleanup(p);
            }
        }
    }
}
