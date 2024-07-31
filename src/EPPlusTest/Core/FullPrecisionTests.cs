using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core
{
    [TestClass]
    public class FullPrecisionTests : TestBase
    {
        [ClassInitialize]
        public static void Init(TestContext context)
        {
        }
        [TestMethod]
        public void Full_Precision_Should_Round_Values_On_Set_Test()
        {
            using(var p=new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                ws.Cells["A1"].Value = "String Value";
                ws.Cells["A2"].Value = 123.456789;
                
                ws.Cells["A3"].Value = 123.456789;
                ws.Cells["A4"].Value = -123.456789;
                ws.Cells["A3:A4"].Style.Numberformat.Format = "#,##0.00;-#,##0.000;0.0";

                p.Workbook.FullPrecision = false;

                Assert.AreEqual("String Value", ws.Cells["A1"].Value);
                Assert.AreEqual(123.456789, ws.Cells["A2"].Value);
                Assert.AreEqual(123.46, ws.Cells["A3"].Value);
                Assert.AreEqual(-123.457, ws.Cells["A4"].Value);
            }
        }
        [TestMethod]
        public void Should_Round_Values_When_Set_Value_On_Range_With_Full_Precision_False_Test()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                p.Workbook.FullPrecision = false;

                ws.Cells["A3:A4"].Style.Numberformat.Format = "#,##0.00;-#,##0.000;0.0";

                ws.Cells["A1"].Value = "String Value";
                ws.Cells["A2"].Value = 123.456789;

                ws.Cells["A3"].Value = 123.456789;
                ws.Cells["A4"].Value = -123.456789;

                Assert.AreEqual("String Value", ws.Cells["A1"].Value);
                Assert.AreEqual(123.456789, ws.Cells["A2"].Value);
                Assert.AreEqual(123.46, ws.Cells["A3"].Value);
                Assert.AreEqual(-123.457, ws.Cells["A4"].Value);
            }
        }
        [TestMethod]
        public void Should_Round_Values_When_Set_Number_Format_On_Range_With_Full_Precision_False_Test()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                p.Workbook.FullPrecision = false;

                ws.Cells["A1"].Value = "String Value";
                ws.Cells["A2"].Value = 123.456789;

                ws.Cells["A3"].Value = 123.456789;
                ws.Cells["A4"].Value = -123.456789;

                ws.Cells["A3:A4"].Style.Numberformat.Format = "#,##0.00;-#,##0.000;0.0";

                Assert.AreEqual("String Value", ws.Cells["A1"].Value);
                Assert.AreEqual(123.456789, ws.Cells["A2"].Value);
                Assert.AreEqual(123.46, ws.Cells["A3"].Value);
                Assert.AreEqual(-123.457, ws.Cells["A4"].Value);
            }
        }
        [TestMethod]
        public void Should_Round_LoadFromCollection_With_FullPrecision_False_Test()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                p.Workbook.FullPrecision = false;

                var data = new List<Tuple<string, decimal>>();
                data.Add(new Tuple<string, decimal>("String 1", 123.456789M));
                data.Add(new Tuple<string, decimal>("String 2", 123.456789M));
                data.Add(new Tuple<string, decimal>("String 3", -123.456789M));
                ws.Cells["A1"].LoadFromCollection(data);
                ws.Cells["B2:B3"].Style.Numberformat.Format = "#,##0.00;-#,##0.000;0.0";

                Assert.AreEqual("String 1", ws.Cells["A1"].Value);
                Assert.AreEqual("String 2", ws.Cells["A2"].Value);
                Assert.AreEqual("String 3", ws.Cells["A3"].Value);
                Assert.AreEqual(123.456789M, ws.Cells["B1"].Value);
                Assert.AreEqual(123.46D, ws.Cells["B2"].Value);
                Assert.AreEqual(-123.457D, ws.Cells["B3"].Value);
            }
        }
        [TestMethod]
        public void Should_Round_LoadArray_With_FullPrecision_False_Test()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                p.Workbook.FullPrecision = false;

                var item1 = new object[]
                {
                    "String 1",
                    123.456789D,
                };
                var item2 = new object[]
                {
                    "String 2",
                    123.456789D
                };
                var item3 = new object[]
                {
                    "String 3",
                    -123.456789D
                };

                ws.Cells["A1"].LoadFromArrays(new List<object[]>() { item1, item2, item3 });
                ws.Cells["B2:B3"].Style.Numberformat.Format = "#,##0.00;-#,##0.000;0.0";

                Assert.AreEqual("String 1", ws.Cells["A1"].Value);
                Assert.AreEqual("String 2", ws.Cells["A2"].Value);
                Assert.AreEqual("String 3", ws.Cells["A3"].Value);
                Assert.AreEqual(123.456789D, ws.Cells["B1"].Value);
                Assert.AreEqual(123.46D, ws.Cells["B2"].Value);
                Assert.AreEqual(-123.457D, ws.Cells["B3"].Value);
            }
        }
        [TestMethod]
        public void Should_Round_LoadFromDataTable_With_FullPrecision_False_Test()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                p.Workbook.FullPrecision = false;

                DataTable dt = CreateDataTable();

                ws.Cells["A1"].LoadFromDataTable(dt);
                ws.Cells["B2:B3"].Style.Numberformat.Format = "#,##0.00;-#,##0.000;0.0";

                Assert.AreEqual("String 1", ws.Cells["A1"].Value);
                Assert.AreEqual("String 2", ws.Cells["A2"].Value);
                Assert.AreEqual("String 3", ws.Cells["A3"].Value);
                Assert.AreEqual(123.456789D, ws.Cells["B1"].Value);
                Assert.AreEqual(123.46D, ws.Cells["B2"].Value);
                Assert.AreEqual(-123.457D, ws.Cells["B3"].Value);
            }
        }
        [TestMethod]
        public void Should_Round_LoadFromDataReader_With_FullPrecision_False_Test()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                p.Workbook.FullPrecision = false;

                DataTable dt = CreateDataTable();
                var dr = new DataTableReader(dt);

                ws.Cells["A1"].LoadFromDataReader(dr, false);
                ws.Cells["B2:B3"].Style.Numberformat.Format = "#,##0.00;-#,##0.000;0.0";

                Assert.AreEqual("String 1", ws.Cells["A1"].Value);
                Assert.AreEqual("String 2", ws.Cells["A2"].Value);
                Assert.AreEqual("String 3", ws.Cells["A3"].Value);
                Assert.AreEqual(123.456789D, ws.Cells["B1"].Value);
                Assert.AreEqual(123.46D, ws.Cells["B2"].Value);
                Assert.AreEqual(-123.457D, ws.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void Should_Round_LoadFromText_With_FullPrecision_False_Test()
        {
            using (var p = new ExcelPackage())
            {
                p.Workbook.FullPrecision = false;
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                var text = new StringBuilder();

                text.AppendLine("String 1,123.456789");
                text.AppendLine("String 2,123.456789");
                text.AppendLine("String 3,-123.456789");

                ws.Cells["A1"].LoadFromText(text.ToString());
                ws.Cells["B2:B3"].Style.Numberformat.Format = "#,##0.00;-#,##0.000;0.0";

                Assert.AreEqual("String 1", ws.Cells["A1"].Value);
                Assert.AreEqual("String 2", ws.Cells["A2"].Value);
                Assert.AreEqual("String 3", ws.Cells["A3"].Value);
                Assert.AreEqual(123.456789D, ws.Cells["B1"].Value);
                Assert.AreEqual(123.46D, ws.Cells["B2"].Value);
                Assert.AreEqual(-123.457D, ws.Cells["B3"].Value);
            }
        }
        [TestMethod]
        public void Should_Round_LoadFromText_FixedWidth_With_FullPrecision_False_Test()
        {
            using (var p = new ExcelPackage())
            {
                p.Workbook.FullPrecision = false;
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                var text = new StringBuilder();

                text.AppendLine("String 1 123.456789");
                text.AppendLine("String 2 123.456789");
                text.AppendLine("String 3-123.456789");

                var option = new ExcelTextFormatFixedWidth();
                option.SetColumnPositions(19, 0, 8);
                ws.Cells["A1"].LoadFromText(text.ToString(), option);
                ws.Cells["B2:B3"].Style.Numberformat.Format = "#,##0.00;-#,##0.000;0.0";

                Assert.AreEqual("String 1", ws.Cells["A1"].Value);
                Assert.AreEqual("String 2", ws.Cells["A2"].Value);
                Assert.AreEqual("String 3", ws.Cells["A3"].Value);
                Assert.AreEqual(123.456789D, ws.Cells["B1"].Value);
                Assert.AreEqual(123.46D, ws.Cells["B2"].Value);
                Assert.AreEqual(-123.457D, ws.Cells["B3"].Value);
            }
        }
        private static DataTable CreateDataTable()
        {
            DataTable dt = new DataTable();

            dt.Columns.Add(new DataColumn("String", typeof(string)));
            dt.Columns.Add(new DataColumn("Double", typeof(double)));

            var r1 = dt.NewRow();
            var r2 = dt.NewRow();
            var r3 = dt.NewRow();

            r1.ItemArray = new object[] { "String 1", 123.456789D };
            r2.ItemArray = new object[] { "String 2", 123.456789D };
            r3.ItemArray = new object[] { "String 3", -123.456789D };

            dt.Rows.Add(r1);
            dt.Rows.Add(r2);
            dt.Rows.Add(r3);
            return dt;
        }

    }
}
