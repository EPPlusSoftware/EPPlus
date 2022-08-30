using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToDataTable;
using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections.Generic;
using System.Data;
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
        public void ToDataTableShouldReturnDataTable_WithOneRow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = "John Doe";
                var dt = sheet.Cells["A1:B1"].ToDataTable(x => x.FirstRowIsColumnNames = false);
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual(1, dt.Rows[0][0]);
                Assert.AreEqual("John Doe", dt.Rows[0][1]);
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
                var dt = sheet.Cells["A1:B2"].ToDataTable(o =>
                {
                    o.PredefinedMappingsOnly = true;
                    o.Mappings.Add(1, "Name");
                });
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual(1, dt.Columns.Count);
                Assert.AreEqual(typeof(string), dt.Columns[0].DataType);
                Assert.AreEqual("John Doe", dt.Rows[0]["Name"]);
            }
        }

        [TestMethod]
        public void ToDataTableShouldReturnDataTable_WithColumnMapping()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Name";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = "John Doe";

                var col = new DataColumn();
                col.ColumnName = "Name";
                col.DataType = typeof(string);
                var dt = sheet.Cells["A1:B2"].ToDataTable(o =>
                {
                    o.PredefinedMappingsOnly = true;
                    o.Mappings.Add(1, col);
                });
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
        public void ToDataTableShouldSetPrimaryKeys()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "FirstName";
                sheet.Cells["C1"].Value = "LastName";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = "John";
                sheet.Cells["C2"].Value = "Doe";
                
                // One column
                var dt = sheet.Cells["A1:C2"].ToDataTable(o => {
                    o.SetPrimaryKey("Id");
                });
                Assert.AreEqual("Id", dt.PrimaryKey[0].ColumnName);
                
                // two columns
                dt = sheet.Cells["A1:C2"].ToDataTable(o => {
                    o.SetPrimaryKey("Id", "LastName");
                });
                Assert.AreEqual("Id", dt.PrimaryKey[0].ColumnName);
                Assert.AreEqual("LastName", dt.PrimaryKey[1].ColumnName);

                // one column by index
                dt = sheet.Cells["A1:C2"].ToDataTable(o => {
                    o.SetPrimaryKey(0);
                });
                Assert.AreEqual("Id", dt.PrimaryKey[0].ColumnName);

                // two columns by index
                dt = sheet.Cells["A1:C2"].ToDataTable(o => {
                    o.SetPrimaryKey(0, 2);
                });
                Assert.AreEqual("Id", dt.PrimaryKey[0].ColumnName);
                Assert.AreEqual("LastName", dt.PrimaryKey[1].ColumnName);
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
        public void ToDataTableShouldHandleTransform()
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
                    o.Mappings.Add(0, "Id", typeof(string), true, c => "Id: " + c.ToString());
                });
                var dt = sheet.Cells["A1:B2"].ToDataTable(options);
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual("Id: 1", dt.Rows[0]["Id"]);
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

        [TestMethod]
        public void ToDataTableShouldHandleExcelErrors()
        {
            using (var package = new ExcelPackage())
            {
                var date = DateTime.UtcNow;
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Name";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = "Bob";
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B3"].Value = ExcelErrorValue.Create(eErrorType.Value);
                
                // Default strategy: Count error as blank cell value
                var dt = sheet.Cells["A1:B3"].ToDataTable();
                Assert.AreEqual(2, dt.Rows.Count);
                Assert.AreEqual(3, dt.Rows[1]["Id"]);
                Assert.AreEqual(DBNull.Value, dt.Rows[1]["Name"]);

                dt = sheet.Cells["A1:B3"].ToDataTable(o => o.ExcelErrorParsingStrategy = ExcelErrorParsingStrategy.IgnoreRowWithErrors);
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual(1, dt.Rows[0]["Id"]);
            }
        }

        [TestMethod]
        public void ToDataTableShouldSkipLinesStart()
        {
            using (var package = new ExcelPackage())
            {
                var date = DateTime.UtcNow;
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Name";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = "Bob";
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B3"].Value = "Rob";

                // Default strategy: Count error as blank cell value
                var dt = sheet.Cells["A1:B3"].ToDataTable(o => o.SkipNumberOfRowsStart = 1);
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual(3, dt.Rows[0]["Id"]);
                Assert.AreEqual("Rob", dt.Rows[0]["Name"]);
            }
        }
        
        [TestMethod]
        public void ToDataTableShouldSkipEmptyRows()
        {
            using (var package = new ExcelPackage())
            {
                var date = DateTime.UtcNow;
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Name";
                sheet.Cells["A3"].Value = 1;
                sheet.Cells["B3"].Value = "Bob";
                sheet.Cells["A4"].Value = 3;
                sheet.Cells["B4"].Value = "Rob";

                // Default strategy: Count error as blank cell value
                var dt = sheet.Cells["A1:B4"].ToDataTable(o => o.EmptyRowStrategy = EmptyRowsStrategy.Ignore);
                Assert.AreEqual(2, dt.Rows.Count);
                Assert.AreEqual(1, dt.Rows[0]["Id"]);
                Assert.AreEqual("Rob", dt.Rows[1]["Name"]);

                sheet.Cells.Clear();
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Name";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = "Bob";
                sheet.Cells["A4"].Value = 3;
                sheet.Cells["B4"].Value = "Rob";

                dt = sheet.Cells["A1:B4"].ToDataTable(o => o.EmptyRowStrategy = EmptyRowsStrategy.StopAtFirst);
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual(1, dt.Rows[0]["Id"]);
            }
        }

        [TestMethod]
        public void ToDataTableShouldSkipLinesEnd()
        {
            using (var package = new ExcelPackage())
            {
                var date = DateTime.UtcNow;
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Name";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = "Bob";
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B3"].Value = "Rob";

                // Default strategy: Count error as blank cell value
                var dt = sheet.Cells["A1:B3"].ToDataTable(o => o.SkipNumberOfRowsEnd = 1);
                Assert.AreEqual(1, dt.Rows.Count);
                Assert.AreEqual(1, dt.Rows[0]["Id"]);
                Assert.AreEqual("Bob", dt.Rows[0]["Name"]);
            }
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ToDataTableShouldHandleAllowNulls()
        {
            using (var package = new ExcelPackage())
            {
                var date = DateTime.UtcNow;
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Name";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = "Bob";
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B3"].Value = null;

                var dt = sheet.Cells["A1:B3"].ToDataTable(o =>
                {
                    o.Mappings.Add(1, "Name", typeof(string), false); 
                });
            }
        }

        [TestMethod]
        public void ToDataTableWithExistingTable_UseOnlyDefinedCols()
        {
            using (var package = new ExcelPackage())
            {
                var date = DateTime.UtcNow;
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Name";
                sheet.Cells["C1"].Value = "Email";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = "Bob";
                sheet.Cells["C2"].Value = "Bobs email";
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B3"].Value = "Rob";
                sheet.Cells["C3"].Value = "Robs email";


                var table = new DataTable("dt1", "ns1");
                table.Columns.Add("Id", typeof(int));
                table.Columns.Add("Email", typeof(string));

                sheet.Cells["A1:C3"].ToDataTable(table);
                Assert.AreEqual("Bobs email", table.Rows[0]["Email"]);

                
            }
        }
    }
}
