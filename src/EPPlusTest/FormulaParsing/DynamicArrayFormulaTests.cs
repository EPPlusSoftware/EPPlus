using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System.Drawing;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]

    public class DynamicArrayFormulaTests : TestBase
    {
        private static ExcelPackage _package;
        private static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _package = OpenPackage("DynamicArrayFormulas.xlsx",true);
            _ws = _package.Workbook.Worksheets.Add("Data");
            LoadTestdata(_ws);
        }

        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_package);
        }
        [TestMethod]
        public void DynamicArrayFormulaSimpleRange()
        {
            _ws.Cells["F1"].Formula = "F10";
            _ws.Cells["F2"].Formula = "DateDif(A2:A11,A12:A21,\"d\")";  //Spill down
            _ws.Cells["E1"].Formula = "SUM(F1:F11)";
            _ws.Calculate();

            Assert.AreEqual(10D, _ws.GetValue(2, 6));
            Assert.AreEqual(10D, _ws.GetValue(1, 6));
            Assert.AreEqual(110D, _ws.GetValue(1, 5));
        }
        [TestMethod, Ignore]
        public void DynamicArrayFormulaReferencedBySharedFormula()
        {
            var ws = _package.Workbook.Worksheets.Add("SharedFormulaRef");
            ws.Cells["F1:N1"].Formula = "F2";
            ws.Cells["F2"].Formula = "Transpose(Data!A2:A10)"; //Spill Right
            ws.Calculate();
            ws.Cells["G2"].Value = 2; //Result in overwrite of the array formula.
            Assert.AreEqual(_ws.GetValue(2, 1), ws.GetValue(2, 6));
            Assert.AreEqual(_ws.GetValue(5, 1), ws.GetValue(2, 9));
            Assert.AreEqual(ConvertUtil.GetValueDouble(ws.GetValue(1, 9)), ConvertUtil.GetValueDouble(ws.GetValue(2, 9)));
            Assert.AreEqual(ConvertUtil.GetValueDouble(ws.GetValue(1, 14)), ConvertUtil.GetValueDouble(ws.GetValue(2, 14)));
        }

        [TestMethod]
        public void DynamicArrayFormulaFilterAndAnchorArrayFunction()
        {
            _ws.Cells["F17"].Value = 2000;
            _ws.Cells["F17"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Dotted);
            _ws.Cells["A1:D1"].Copy(_ws.Cells["F19"]);

            _ws.Cells["F19:I19"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            _ws.Cells["F19:I19"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            _ws.Cells["F19:I19"].Style.Font.Color.SetColor(Color.White);

            _ws.Cells["F19:I59"].Style.Font.Name = "Arial";
            _ws.Cells["F19:I59"].Style.Font.Size = 12;

            _ws.Cells["F20"].Formula = "Filter(A2:D100,D2:D100 > F17)";
            _ws.Cells["F19:I59"].AutoFitColumns();
            _ws.Cells["F100"].Formula = "AnchorArray(F20)"; //F20# in Excel GUI
            _ws.Calculate();
        }
        [TestMethod]
        public void DynamicArrayFormulaFilterAndAnchorArray()
        {
            _ws.Cells["K17"].Value = 20000;
            _ws.Cells["K17"].Style.Border.BorderAround(ExcelBorderStyle.Dotted);
            _ws.Cells["A1:D1"].Copy(_ws.Cells["K19"]);

            _ws.Cells["K19:N19"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            _ws.Cells["K19:N19"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            _ws.Cells["K19:N19"].Style.Font.Color.SetColor(Color.White);

            _ws.Cells["K19:N59"].Style.Font.Name = "Arial";
            _ws.Cells["K19:N59"].Style.Font.Size = 12;

            _ws.Cells["K20"].Formula = "Filter(A2:D100,D2:D100 > K17)";
            _ws.Cells["K19:N59"].AutoFitColumns();
            _ws.Calculate();
        }
    }
}
