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

                var b1 = ws.Cells["B1"].Value;
                var c1 = ws.Cells["C1"].Value;
                Assert.IsInstanceOfType(b1, typeof(ExcelErrorValue));
                Assert.AreEqual(((ExcelErrorValue)b1).Type, eErrorType.Spill);

                Assert.IsInstanceOfType(c1, typeof(ExcelErrorValue));
                Assert.AreEqual(((ExcelErrorValue)c1).Type, eErrorType.Calc);

                Assert.IsInstanceOfType(ws.Cells["F1"].Value, typeof(ExcelErrorValue));
                Assert.AreEqual(((ExcelErrorValue)ws.Cells["F1"].Value).Type, eErrorType.Spill);

                Assert.IsInstanceOfType(ws.Cells["E10"].Value, typeof(ExcelErrorValue));
                Assert.AreEqual(((ExcelErrorValue)ws.Cells["E10"].Value).Type, eErrorType.Spill);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void VerifySpillError()
        {
            using (var p = OpenPackage("SpillError.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Formula = "RandArray(3,3)";
                ws.Cells["B3"].Value = 4;
                ws.Calculate();
                Assert.IsInstanceOfType(ws.Cells["A1"].Value, typeof(ExcelRichDataErrorValue));
                SaveAndCleanup(p);
            }
        }

    }
}
