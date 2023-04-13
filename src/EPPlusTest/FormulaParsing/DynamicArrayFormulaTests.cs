using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Exceptions;
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
        [TestMethod]
        public void DynamicArrayFormulaReferencedBySharedFormula()
        {
            var ws = _package.Workbook.Worksheets.Add("SharedFormulaRef");
            ws.Cells["F1:N1"].Formula = "F2";
            ws.Cells["F2"].Formula = "Transpose(Data!A2:A10)"; //Spill Right
            ws.Calculate();

            Assert.AreEqual(_ws.GetValue(2, 1), ws.GetValue(2, 6));
            Assert.AreEqual(_ws.GetValue(5, 1), ws.GetValue(2, 9));
            Assert.AreEqual(ConvertUtil.GetValueDouble(ws.GetValue(1, 9)), ConvertUtil.GetValueDouble(ws.GetValue(2, 9)));
            Assert.AreEqual(ConvertUtil.GetValueDouble(ws.GetValue(1, 14)), ConvertUtil.GetValueDouble(ws.GetValue(2, 14)));
        }

        [TestMethod]
        public void ArrayFormulaFilterFunction()
        {
            _ws.Cells["F17"].Value = 2000;
            _ws.Cells["F17"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Dotted);
            _ws.Cells["A1:D1"].Copy(_ws.Cells["F19"]);
            _ws.Cells["F19:I19"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            _ws.Cells["F19:I19"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            _ws.Cells["F19:I19"].Style.Font.Color.SetColor(Color.White);
            _ws.Cells["F19:I59"].Style.Font.Name = "Arial";
            _ws.Cells["F19:I59"].Style.Font.Size = 12;
            _ws.Cells["F20"].Formula = "_xlfn._xlws.Filter(A2:D100,D2:D100 > F17)";
            _ws.Cells["F19:I59"].AutoFitColumns();
            _ws.Cells["F100"].Formula = "_xlfn.AnchorArray(F20)";
            _ws.Calculate();

            //Assert.AreEqual(10D, _ws.GetValue(2, 6));
            //Assert.AreEqual(_ws.GetValue(2, 1), _ws.GetValue(3, 6));
            //Assert.AreEqual(_ws.GetValue(9, 1), _ws.GetValue(10, 6));

            //Assert.AreEqual(((DateTime)_ws.GetValue(10, 6)).ToOADate(), _ws.GetValue(1, 6));    //A1 is converted to AO-date. Correct?
        }
        //[TestMethod]
        //public void ArrayFormulaMultiplyRange()
        //{
        //    _ws.Cells["G1"].Formula = "G10";
        //    _ws.Cells["G2:G20"].CreateArrayFormula("(B2:B10 + 1) * 2");
        //    _ws.Calculate();

        //    Assert.AreEqual((_ws.Cells["B2"].GetValue<double>() + 1) * 2, _ws.Cells["G2"].Value);
        //    Assert.AreEqual((_ws.Cells["B10"].GetValue<double>() + 1) * 2, _ws.Cells["G10"].Value);
        //    Assert.AreEqual(_ws.Cells["G10"].Value, _ws.Cells["G1"].Value);

        //    Assert.AreEqual(((ExcelErrorValue)_ws.Cells["G11"].Value).Type, eErrorType.NA);
        //    Assert.AreEqual(((ExcelErrorValue)_ws.Cells["G20"].Value).Type, eErrorType.NA);
        //}
        //[TestMethod]
        //public void ArrayFormula_Transpose()
        //{
        //    _ws.Cells["G1"].Formula = "G10";
        //    _ws.Cells["H5:P5"].CreateArrayFormula("Transpose(C2:C10)");
        //    _ws.Calculate();

        //    Assert.AreEqual("Value 2", _ws.Cells["H5"].Value);
        //    Assert.AreEqual("Value 3", _ws.Cells["I5"].Value);
        //    Assert.AreEqual("Value 10", _ws.Cells["P5"].Value);
        //}
        //[TestMethod]
        //public void ArrayFormula_Round()
        //{
        //    _ws.Cells["F15:F20"].CreateArrayFormula("Round(D2:D10,-1)");
        //    _ws.Calculate();

        //    Assert.AreEqual(70D, _ws.Cells["F15"].Value);
        //    Assert.AreEqual(100D, _ws.Cells["F16"].Value);
        //    Assert.AreEqual(130D, _ws.Cells["F17"].Value);
        //    Assert.AreEqual(170D, _ws.Cells["F18"].Value);
        //    Assert.AreEqual(200D, _ws.Cells["F19"].Value);
        //    Assert.AreEqual(230D, _ws.Cells["F20"].Value);
        //    Assert.IsNull(_ws.Cells["F21"].Value);
        //}

        //[TestMethod]
        //[ExpectedException(typeof(CircularReferenceException))]
        //public void ArrayFormulaCircularRefererence()
        //{
        //    using (var p = new ExcelPackage())
        //    {
        //        var ws= p.Workbook.Worksheets.Add("CircularRef");
        //        ws.Cells["T1:T10"].CreateArrayFormula("Transpose(Q5:Z5)");
        //        ws.Calculate();
        //    }
        //}

    }
}
