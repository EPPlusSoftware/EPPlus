using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Table.PivotTable.Calculation;
using System.Collections.Generic;

namespace EPPlus.Core.Tests.Table.PivotTable
{
    [TestClass]
    public class GetPivotDataTests_CalculatedFields2
    {
        private ExcelPackage _package;
        private ExcelWorksheet _dataWs;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _dataWs = _package.Workbook.Worksheets.Add("Data");

            // Headers
            _dataWs.Cells["A1"].Value = "Department";
            _dataWs.Cells["B1"].Value = "Basic Pay";
            _dataWs.Cells["C1"].Value = "Overtime";
            _dataWs.Cells["D1"].Value = "Bonus";

            // Sales department
            _dataWs.Cells["A2"].Value = "Sales";
            _dataWs.Cells["B2"].Value = 5000d;
            _dataWs.Cells["C2"].Value = 500d;
            _dataWs.Cells["D2"].Value = 1000d;

            _dataWs.Cells["A3"].Value = "Sales";
            _dataWs.Cells["B3"].Value = 6000d;
            _dataWs.Cells["C3"].Value = 600d;
            _dataWs.Cells["D3"].Value = 1200d;

            // IT department
            _dataWs.Cells["A4"].Value = "IT";
            _dataWs.Cells["B4"].Value = 7000d;
            _dataWs.Cells["C4"].Value = 300d;
            _dataWs.Cells["D4"].Value = 800d;

            _dataWs.Cells["A5"].Value = "IT";
            _dataWs.Cells["B5"].Value = 7500d;
            _dataWs.Cells["C5"].Value = 400d;
            _dataWs.Cells["D5"].Value = 900d;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package?.Dispose();
        }

        [TestMethod]
        public void GetPivotData_CalculatedField_ShouldReturnCorrectValue()
        {
            // Arrange
            var ws = _package.Workbook.Worksheets.Add("Pivot");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _dataWs.Cells["A1:D5"], "TestPivot");

            pt.RowFields.Add(pt.Fields["Department"]);

            pt.Fields.AddCalculatedField("Total Comp", "'Basic Pay'+'Overtime'+'Bonus'");
            var df = pt.DataFields.Add(pt.Fields["Total Comp"]);
            df.Function = DataFieldFunctions.Sum;
            df.Name = "Sum of Total Comp"; 

            // Act
            pt.Calculate(refreshCache: true);

            var criteria = new List<PivotDataFieldItemSelection>
            {
                new PivotDataFieldItemSelection("Department", "Sales")
            };

            var result = pt.GetPivotData("Sum of Total Comp", criteria);

            // Assert
            Assert.IsNotNull(result, "Result should not be null");
            Assert.IsFalse(result is ExcelErrorValue,
                $"GetPivotData should not return #REF! error, got: {result}");

            // Sales: (5000+500+1000) + (6000+600+1200) = 14300
            Assert.AreEqual(14300d, (double)result, 0.001,
                "Calculated field value for Sales should be 14300");
        }

        [TestMethod]
        public void GetPivotData_CalculatedField_GrandTotal_ShouldWork()
        {
            // Arrange
            var ws = _package.Workbook.Worksheets.Add("Pivot");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _dataWs.Cells["A1:D5"], "TestPivot");

            pt.RowFields.Add(pt.Fields["Department"]);
            pt.Fields.AddCalculatedField("All Pay", "'Basic Pay'+'Overtime'+'Bonus'");
            var df = pt.DataFields.Add(pt.Fields["All Pay"]);
            df.Function = DataFieldFunctions.Sum;
            df.Name = "Sum of All Pay";

            // Act
            pt.Calculate(refreshCache: true);
            var grandTotal = pt.GetPivotData("Sum of All Pay");

            // Assert
            Assert.IsNotNull(grandTotal);
            Assert.IsFalse(grandTotal is ExcelErrorValue,
                string.Format("Grand total should not return error, got: {0}", grandTotal));

            // Total: Sales(14300) + IT(16900) = 31200
            Assert.AreEqual(31200d, (double)grandTotal, 0.001);
        }

        [TestMethod]
        public void GetPivotData_MultipleCalculatedFields_ShouldWorkIndependently()
        {
            // Arrange
            var ws = _package.Workbook.Worksheets.Add("Pivot");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _dataWs.Cells["A1:D5"], "TestPivot");

            pt.RowFields.Add(pt.Fields["Department"]);

            pt.Fields.AddCalculatedField("Base Plus OT", "'Basic Pay'+'Overtime'");
            pt.Fields.AddCalculatedField("Total Comp", "'Basic Pay'+'Overtime'+'Bonus'");

            var df1 = pt.DataFields.Add(pt.Fields["Base Plus OT"]);
            df1.Function = DataFieldFunctions.Sum;
            df1.Name = "Sum of Base Plus OT";

            var df2 = pt.DataFields.Add(pt.Fields["Total Comp"]);
            df2.Function = DataFieldFunctions.Sum;
            df2.Name = "Sum of Total Comp";

            // Act
            pt.Calculate(refreshCache: true);

            var criteriaIT = new List<PivotDataFieldItemSelection>
            {
                new PivotDataFieldItemSelection("Department", "IT")
            };

            var result1 = pt.GetPivotData("Sum of Base Plus OT", criteriaIT);
            var result2 = pt.GetPivotData("Sum of Total Comp", criteriaIT);

            // Assert
            Assert.IsFalse(result1 is ExcelErrorValue);
            Assert.IsFalse(result2 is ExcelErrorValue);

            // IT Base + OT: (7000+300) + (7500+400) = 15200
            Assert.AreEqual(15200d, (double)result1, 0.001);

            // IT Total: (7000+300+800) + (7500+400+900) = 16900
            Assert.AreEqual(16900d, (double)result2, 0.001);
        }

        [TestMethod]
        public void GetPivotData_CalculatedField_CalculatedItems_ShouldBePopulated()
        {
            // Arrange
            var ws = _package.Workbook.Worksheets.Add("Pivot");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _dataWs.Cells["A1:D5"], "TestPivot");

            pt.RowFields.Add(pt.Fields["Department"]);
            pt.Fields.AddCalculatedField("Sum Total", "'Basic Pay'+'Overtime'");
            var df = pt.DataFields.Add(pt.Fields["Sum Total"]);
            df.Function = DataFieldFunctions.Sum;
            df.Name = "Sum of Sum Total";

            // Act
            pt.Calculate(refreshCache: true);

            // Assert
            Assert.IsTrue(pt.IsCalculated, "Pivot table should be marked as calculated");
            Assert.IsNotNull(pt.CalculatedItems, "CalculatedItems should not be null");

            var dfIndex = pt.DataFields.IndexOf(df);
            Assert.IsTrue(dfIndex >= 0, "DataField should be in collection");
            Assert.IsTrue(dfIndex < pt.CalculatedItems.Count,
                "Should have CalculatedItems entry for this index");

            var store = pt.CalculatedItems[dfIndex];
            Assert.IsNotNull(store, "Store should not be null");
            Assert.IsTrue(store.Count > 0,
                "CalculatedItems store should be populated (this was empty before the fix)");
        }

        [TestMethod]
        public void GetPivotData_CalculatedFieldMixedWithRegular_ShouldWork()
        {
            // Arrange
            var ws = _package.Workbook.Worksheets.Add("Pivot");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _dataWs.Cells["A1:D5"], "TestPivot");

            pt.RowFields.Add(pt.Fields["Department"]);

            // Add regular data field first
            var dfBasic = pt.DataFields.Add(pt.Fields["Basic Pay"]);
            dfBasic.Function = DataFieldFunctions.Sum;
            dfBasic.Name = "Sum of Basic Pay";

            // Add calculated field
            pt.Fields.AddCalculatedField("Total", "'Basic Pay'+'Overtime'+'Bonus'");
            var dfCalc = pt.DataFields.Add(pt.Fields["Total"]);
            dfCalc.Function = DataFieldFunctions.Sum;
            dfCalc.Name = "Sum of Total";

            // Act
            pt.Calculate(refreshCache: true);

            var criteriaSales = new List<PivotDataFieldItemSelection>
            {
                new PivotDataFieldItemSelection("Department", "Sales")
            };

            var basicResult = pt.GetPivotData("Sum of Basic Pay", criteriaSales);
            var calcResult = pt.GetPivotData("Sum of Total", criteriaSales);

            // Assert
            Assert.IsFalse(basicResult is ExcelErrorValue, "Regular field should work");
            Assert.IsFalse(calcResult is ExcelErrorValue, "Calculated field should work");

            Assert.AreEqual(11000d, (double)basicResult, 0.001); // 5000 + 6000
            Assert.AreEqual(14300d, (double)calcResult, 0.001);
        }

        [TestMethod]
        public void GetPivotData_CalculatedFieldWithComplexFormula_ShouldWork()
        {
            // Arrange
            var ws = _package.Workbook.Worksheets.Add("Pivot");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _dataWs.Cells["A1:D5"], "TestPivot");

            pt.RowFields.Add(pt.Fields["Department"]);

            // Complex formula with multiple operations
            pt.Fields.AddCalculatedField("Complex Calc",
                "'Basic Pay' * 2 + 'Overtime' - 'Bonus'");
            var df = pt.DataFields.Add(pt.Fields["Complex Calc"]);
            df.Function = DataFieldFunctions.Sum;
            df.Name = "Sum of Complex Calc";

            // Act
            pt.Calculate(refreshCache: true);

            var criteriaIT = new List<PivotDataFieldItemSelection>
            {
                new PivotDataFieldItemSelection("Department", "IT")
            };

            var result = pt.GetPivotData("Sum of Complex Calc", criteriaIT);

            // Assert
            Assert.IsFalse(result is ExcelErrorValue);

            // IT: (7000*2+300-800) + (7500*2+400-900) 
            //   = (14000+300-800) + (15000+400-900)
            //   = 13500 + 14500 = 28000
            Assert.AreEqual(28000d, (double)result, 0.001);
        }

        [TestMethod]
        public void GetPivotData_CalculatedFieldOnlyWithFormula_NoRegularFields()
        {
            // Arrange
            var ws = _package.Workbook.Worksheets.Add("Pivot");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _dataWs.Cells["A1:D5"], "TestPivot");

            pt.RowFields.Add(pt.Fields["Department"]);

            // Only add calculated field, no regular data fields
            pt.Fields.AddCalculatedField("Only Calc", "'Basic Pay'+'Overtime'");
            var df = pt.DataFields.Add(pt.Fields["Only Calc"]);
            df.Function = DataFieldFunctions.Sum;
            df.Name = "Sum of Only Calc";

            // Act
            pt.Calculate(refreshCache: true);

            var criteriaIT = new List<PivotDataFieldItemSelection>
            {
                new PivotDataFieldItemSelection("Department", "IT")
            };

            var result = pt.GetPivotData("Sum of Only Calc", criteriaIT);

            // Assert
            Assert.IsNotNull(result);
            Assert.IsFalse(result is ExcelErrorValue,
                "Should work even when pivot has only calculated fields");

            // IT: (7000+300) + (7500+400) = 15200
            Assert.AreEqual(15200d, (double)result, 0.001);
        }

        [TestMethod]
        public void GetPivotData_CalculatedField_DifferentDepartments()
        {
            // Arrange
            var ws = _package.Workbook.Worksheets.Add("Pivot");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _dataWs.Cells["A1:D5"], "TestPivot");

            pt.RowFields.Add(pt.Fields["Department"]);
            pt.Fields.AddCalculatedField("Total Pay", "'Basic Pay'+'Overtime'+'Bonus'");
            var df = pt.DataFields.Add(pt.Fields["Total Pay"]);
            df.Function = DataFieldFunctions.Sum;
            df.Name = "Sum of Total Pay";

            // Act
            pt.Calculate(refreshCache: true);

            var criteriaSales = new List<PivotDataFieldItemSelection>
            {
                new PivotDataFieldItemSelection("Department", "Sales")
            };

            var criteriaIT = new List<PivotDataFieldItemSelection>
            {
                new PivotDataFieldItemSelection("Department", "IT")
            };

            var salesResult = pt.GetPivotData("Sum of Total Pay", criteriaSales);
            var itResult = pt.GetPivotData("Sum of Total Pay", criteriaIT);

            // Assert
            Assert.IsFalse(salesResult is ExcelErrorValue);
            Assert.IsFalse(itResult is ExcelErrorValue);

            // Sales: 14300, IT: 16900
            Assert.AreEqual(14300d, (double)salesResult, 0.001);
            Assert.AreEqual(16900d, (double)itResult, 0.001);
        }
    }
}