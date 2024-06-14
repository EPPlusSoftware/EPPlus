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

            Assert.AreEqual("F2:F11", _ws.Cells["F2"].FormulaRange.Address);
            Assert.AreEqual("F2:F11", _ws.GetFormulaRange(2,6).Address);
        }

        [TestMethod]
        public void DynamicArrayFormulaReferencedBySharedFormula()
        {
            var ws = _package.Workbook.Worksheets.Add("SharedFormulaRef");
            ws.Cells["F1:N1"].Formula = "F2";
            ws.Cells["F2"].Formula = "Transpose(Data!A2:A10)"; //Spill Right
            ws.Calculate();
            Assert.AreEqual(_ws.GetValue(2, 1), ws.GetValue(2, 6));
            Assert.AreEqual(_ws.GetValue(2, 1), ws.GetValue(2, 6));
            Assert.AreEqual(_ws.GetValue(5, 1), ws.GetValue(2, 9));
            Assert.AreEqual(ConvertUtil.GetValueDouble(ws.GetValue(1, 9)), ConvertUtil.GetValueDouble(ws.GetValue(2, 9)));
            Assert.AreEqual(ConvertUtil.GetValueDouble(ws.GetValue(1, 14)), ConvertUtil.GetValueDouble(ws.GetValue(2, 14)));

            Assert.AreEqual("F2:N2", ws.Cells["F2"].FormulaRange.Address);
            Assert.AreEqual("F2:N2", ws.GetFormulaRange(2, 6).Address);
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

            Assert.AreEqual("F20:I59", _ws.Cells["F20"].FormulaRange.Address);
            Assert.AreEqual("F20:I59", _ws.GetFormulaRange(20, 6).Address);
            Assert.AreEqual("F100:I139", _ws.Cells["F100"].FormulaRange.Address);
            Assert.AreEqual("F100:I139", _ws.GetFormulaRange(100, 6).Address);
        }
        [TestMethod]
        public void DynamicArrayFormulaFilterAndAnchorArrayFunctionWithSpill()
        {
            _ws.Cells["Q17"].Value = 20000;
            _ws.Cells["Q17"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Dotted);
            _ws.Cells["A1:D1"].Copy(_ws.Cells["Q19"]);

            _ws.Cells["Q19:T19"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            _ws.Cells["Q19:T19"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            _ws.Cells["Q19:T19"].Style.Font.Color.SetColor(Color.White);
                       
            _ws.Cells["Q19:T59"].Style.Font.Name = "Arial";
            _ws.Cells["Q19:T59"].Style.Font.Size = 12;
                       
            _ws.Cells["Q20"].Formula = "Filter(A2:D100,D2:D100 > Q17, \"No matches found.\")";
            _ws.Cells["Q19:T59"].AutoFitColumns();
            _ws.Cells["Q100"].Formula = "AnchorArray(Q20)"; //F20# in Excel GUI

            _ws.Cells["R110"].Value = 88; //result in spill
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
        [TestMethod]
        public void DynamicArrayFormulaWithSpill()
        {
            var ws = _package.Workbook.Worksheets.Add("CalcSpill");
            ws.Cells["F1:N1"].Formula = "F2";
            ws.Cells["F2"].Formula = "Transpose(Data!A2:A10)"; //Spill Right
            ws.Cells["G2"].Value = 2;
            ws.Calculate();
            var v = ws.GetValue(2, 6);
            Assert.IsInstanceOfType(v, typeof(ExcelRichDataErrorValue));
            var spillError = (ExcelRichDataErrorValue)v;
            Assert.AreEqual(0, spillError.SpillRowOffset);
            Assert.AreEqual(1, spillError.SpillColOffset);
        }
        [TestMethod]
        public void DynamicArrayReadFromWorkbook()
        {
            using (var p = OpenTemplatePackage("ArrayFormulas.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                ws.Cells.ClearFormulaValues();
                ws.Cells["B6"].Value = 3;
                p.Workbook.Calculate();

                Assert.AreEqual(5D, ws.Cells["B1"].Value);
                Assert.AreEqual(5D, ws.Cells["C5"].Value);

                Assert.AreEqual(4D, ws.Cells["D1"].Value);
                Assert.AreEqual(4D, ws.Cells["D1"].Value);
                Assert.AreEqual(5D, ws.Cells["D2"].Value);
                Assert.IsNull(ws.Cells["D3"].Value);

                Assert.AreEqual(4D, ws.Cells["F1"].Value);
                Assert.AreEqual(5D, ws.Cells["F2"].Value);
                Assert.IsNull(ws.Cells["F3"].Value);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void DynamicArrayReadFromWorkbook_DeleteFullFormula()
        {
            using (var p = OpenTemplatePackage("ArrayFormulas.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                ws.Workbook.Worksheets[0].Cells["B1:B5"].Delete(eShiftTypeDelete.Left);
                SaveWorkbook("ArrayFormulas_Deleted.xlsx",p);
            }
        }
        [TestMethod]
        public void DynamicArrayReadFromWorkbook_InsertInsideFormula()
        {
            using (var p = OpenTemplatePackage("ArrayFormulas.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                ws.Workbook.Worksheets[0].InsertColumn(2, 1);
                ws.Calculate();
                SaveWorkbook("ArrayFormulas_Deleted.xlsx", p);
            }
        }
        [TestMethod]
        public void DynamicFunctionWithChart()
        {
            _ws.Cells[20, 20].Formula = "RandArray(5,5)";
            _ws.Calculate();

            var chart = _ws.Drawings.AddBarChart("Dynamic Chart", OfficeOpenXml.Drawing.Chart.eBarChartType.ColumnClustered);
            chart.StyleManager.SetChartStyle(OfficeOpenXml.Drawing.Chart.Style.ePresetChartStyle.ColumnChartStyle9);

            var address = _ws.Cells[20, 20].FormulaRange;
            Assert.AreEqual("T20:X24", address.Address);
            for (var c = address.Start.Column; c <= address.End.Column; c++)
            {
                chart.Series.Add(_ws.Cells[address.Start.Row, c, address.End.Row, c]);
            }

            chart.SetPosition(10, 0, 25, 0);

        }
    }
}
