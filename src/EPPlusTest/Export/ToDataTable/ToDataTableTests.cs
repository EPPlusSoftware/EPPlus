using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToDataTable;
using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Export.ToDataTable
{
    [TestClass]
    public class ToDataTableTests
    {
        [TestMethod]
        public void ToDataTableShouldReturnDataTable_WithDefaultOptions()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Name";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = "John Doe";
                var dt = sheet.Cells["A1:B2"].ToDataTable();
                Assert.AreEqual("dataTable1", dt.TableName);
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual(1, dt.Rows[0]["Id"]);
                Assert.AreEqual("John Doe", dt.Rows[0]["Name"]);
            }
        }

        [TestMethod]
        public void ToDataTableShouldReturnDataTable_WithOneMapping()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Name";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = "John Doe";
                var options = ToDataTableOptions.Create(o =>
                {
                    o.PredefinedMappingsOnly = true;
                    o.Mappings.Add(1, "Name");
                });
                var dt = sheet.Cells["A1:B2"].ToDataTable(options);
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual(1, dt.Columns.Count);
                Assert.AreEqual(typeof(string), dt.Columns[0].DataType);
                Assert.AreEqual("John Doe", dt.Rows[0]["Name"]);
            }
        }

        [TestMethod]
        public void ToDataTableShouldSetNamespace()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Name";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = "John Doe";
                var options = ToDataTableOptions.Create(o =>
                {
                    o.DataTableNamespace = "ns1";
                });
                var dt = sheet.Cells["A1:B2"].ToDataTable(options);
                Assert.AreEqual("ns1", dt.Namespace);
            }
        }

        [TestMethod]
        public void ToDataTableShouldHandleDateTime()
        {
            using (var package = new ExcelPackage())
            {
                var date = DateTime.UtcNow;
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Date";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = date;
                var dt = sheet.Cells["A1:B2"].ToDataTable();
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual(1, dt.Rows[0]["Id"]);
                Assert.AreEqual(date, dt.Rows[0]["Date"]);
            }
        }
        [TestMethod]
        public void ToDataTableShouldHandleIntToString()
        {
            using (var package = new ExcelPackage())
            {
                var date = DateTime.UtcNow;
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Date";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = date;
                var options = ToDataTableOptions.Create(o =>
                {
                    o.Mappings.Add(0,"Id",typeof(string));
                });
                var dt = sheet.Cells["A1:B2"].ToDataTable(options);
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual("1", dt.Rows[0]["Id"]);
                Assert.AreEqual(date, dt.Rows[0]["Date"]);
            }
        }


        [TestMethod]
        public void ToDataTableShouldHandleIntAndBool()
        {
            using (var package = new ExcelPackage())
            {
                var date = DateTime.UtcNow;
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "IsBool";
                sheet.Cells["A2"].Value = 3;
                sheet.Cells["B2"].Value = true;
                var dt = sheet.Cells["A1:B2"].ToDataTable();
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual(3, dt.Rows[0]["Id"]);
                Assert.IsTrue((bool)dt.Rows[0]["IsBool"]);
                Assert.AreEqual(typeof(int), dt.Columns[0].DataType);
            }
        }

        [TestMethod]
        public void ToDataTableShouldHandleDateTimeWithMapping()
        {
            using (var package = new ExcelPackage())
            {
                var date = DateTime.UtcNow;
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Date";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = date;
                var options = ToDataTableOptions.Create(o =>
                {
                    o.Mappings.Add(1, "MyDate", typeof(DateTime));
                });
                var dt = sheet.Cells["A1:B2"].ToDataTable(options);
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual(1, dt.Rows[0]["Id"]);
                Assert.AreEqual(date.ToOADate(), ((DateTime)dt.Rows[0]["MyDate"]).ToOADate());
            }
        }
    }
}
