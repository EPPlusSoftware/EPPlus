using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromDataTableTests
    {
        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("test");
            _dataSet = new DataSet();
            _table = _dataSet.Tables.Add("table");
            _table.Columns.Add("Id", typeof(string));
            _table.Columns.Add("Name", typeof(string));
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;
        private DataSet _dataSet;
        private DataTable _table;

        [TestMethod]
        public void ShouldLoadTable()
        {
            _table.Rows.Add("1", "Test name");
            _worksheet.Cells["A1"].LoadFromDataTable(_table, false);
            Assert.AreEqual("1", _worksheet.Cells["A1"].Value);
        }
        [TestMethod]
        public void ShouldLoadTableTransposed()
        {
            _table.Rows.Add("1", "Testname 1");
            _table.Rows.Add("2", "Testname 2");
            _table.Rows.Add("3", "Testname 3");
            var r = _worksheet.Cells["A1"].LoadFromDataTable(_table, false, TableStyles.None, true);
            Assert.AreEqual("A1:C2", r.Address);
            Assert.AreEqual("1", _worksheet.Cells["A1"].Value);
            Assert.AreEqual("2", _worksheet.Cells["B1"].Value);
            Assert.AreEqual("3", _worksheet.Cells["C1"].Value);
            Assert.AreEqual("Testname 1", _worksheet.Cells["A2"].Value);
            Assert.AreEqual("Testname 2", _worksheet.Cells["B2"].Value);
            Assert.AreEqual("Testname 3", _worksheet.Cells["C2"].Value);
        }

        [TestMethod]
        public void CreateAndFillDataTable()
        {
            var table = new DataTable("Astronauts");
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("FirstName", typeof(string));
            table.Columns.Add("LastName", typeof(string));
            table.Columns["FirstName"].Caption = "First name";
            table.Columns["LastName"].Caption = "Last name";
            // add some data
            table.Rows.Add(1, "Bob", "Behnken");
            table.Rows.Add(2, "Doug", "Hurley");

            //create a workbook with a spreadsheet and load the data table
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Astronauts");
                sheet.Cells["A1"].LoadFromDataTable(table);
            }
        }

        [TestMethod]
        public void ShouldLoadTableWithTableStyle()
        {
            _table.Rows.Add("1", "Test name");
            _worksheet.Cells["A1"].LoadFromDataTable(_table, false, TableStyles.Dark1);
            Assert.AreEqual(1, _worksheet.Tables.Count);
        }

        [TestMethod]
        public void ShouldLoadTableWithTableStyleTransposed()
        {
            _table.Rows.Add("1", "Test name");
            _worksheet.Cells["A1"].LoadFromDataTable(_table, false, TableStyles.Dark1, true);
            Assert.AreEqual(1, _worksheet.Tables.Count);
        }

        [TestMethod]
        public void ShouldUseCaptionForHeader()
        {
            _table.Columns["Id"].Caption = "An Id";
            _table.Columns["Name"].Caption = "A name";
            _worksheet.Cells["A1"].LoadFromDataTable(_table, true);
            Assert.AreEqual("An Id", _worksheet.Cells["A1"].Value);
        }
        [TestMethod]
        public void ShouldUseCaptionForHeaderTransposed()
        {
            _table.Columns["Id"].Caption = "An Id";
            _table.Columns["Name"].Caption = "A name";
            _worksheet.Cells["A1"].LoadFromDataTable(_table, true, TableStyles.None, true);
            Assert.AreEqual("An Id", _worksheet.Cells["A1"].Value);
            Assert.AreEqual("A name", _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void ShouldUseColumnNameForHeaderIfNoCaption()
        {
            _worksheet.Cells["A1"].LoadFromDataTable(_table, true);
            Assert.AreEqual("Id", _worksheet.Cells["A1"].Value);
        }
        [TestMethod]
        public void ShouldUseColumnNameForHeaderIfNoCaptionTransposed()
        {
            _worksheet.Cells["A1"].LoadFromDataTable(_table, true, TableStyles.None, true);
            Assert.AreEqual("Id", _worksheet.Cells["A1"].Value);
            Assert.AreEqual("Name", _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void ShouldLoadXmlFromDataset()
        {
            var dataSet = new DataSet();
            var xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                        "<Astronauts>" +
                        "<Astronaut Id=\"1\">" +
                        "<FirstName>Bob</FirstName>" +
                        "<LastName>Behnken</LastName>" +
                        "</Astronaut>" +
                        "<Astronaut Id=\"2\">" +
                        "<FirstName>Doug</FirstName>" +
                        "<LastName>Hurley</LastName>" +
                        "</Astronaut>" +
                        "</Astronauts>";
            var reader = XmlReader.Create(new StringReader(xml));
            dataSet.ReadXml(reader);
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var table = dataSet.Tables["Astronaut"];
                // default the Id ends up last in the column order. This moves it to the first position.
                table.Columns["Id"].SetOrdinal(0);
                // Set caption for the headers
                table.Columns["FirstName"].Caption = "First name";
                table.Columns["LastName"].Caption = "Last name";
                // call LoadFromDataTable, print headers and use the Dark1 table style
                sheet.Cells["A1"].LoadFromDataTable(table, true, TableStyles.Dark1);
                // AutoFit column with for the entire range
                sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Row].AutoFitColumns();
                //package.SaveAs(new FileInfo(@"c:\temp\astronauts.xlsx"));
            }
        }

        [TestMethod]
        public void ShouldUseLambdaConfig()
        {
            _table.Rows.Add("1", "Test name");
            _worksheet.Cells["A1"].LoadFromDataTable(_table, c =>
            {
                c.PrintHeaders = true;
                c.TableStyle = TableStyles.Dark1;
            });
            Assert.AreEqual("Id", _worksheet.Cells["A1"].Value);
            Assert.AreEqual("1", _worksheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void ShouldSetDbNullToNull()
        {
            _table.Rows.Add("1", DBNull.Value);
            _worksheet.Cells["A1"].LoadFromDataTable(_table, c =>
            {
                c.PrintHeaders = true;
                c.TableStyle = TableStyles.Dark1;
            });
            Assert.IsNull(_worksheet.Cells["B2"].Value);
        }

        [TestMethod]
        public void ShouldSetNullToNull()
        {
            _table.Rows.Add("1", null);
            _worksheet.Cells["A1"].LoadFromDataTable(_table, c =>
            {
                c.PrintHeaders = true;
                c.TableStyle = TableStyles.Dark1;
            });
            Assert.IsNull(_worksheet.Cells["B2"].Value);
        }

        [TestMethod]
        public void ShouldReplaceWithNullIfDbNull()
        {
            _table.Rows.Add("1", null);
            _worksheet.Cells["B2"].Value = 2;
            _worksheet.Cells["A1"].LoadFromDataTable(_table, c =>
            {
                c.PrintHeaders = true;
                c.TableStyle = TableStyles.Dark1;
            });
            Assert.IsNull(_worksheet.Cells["B2"].Value);
        }
        [TestMethod]
        public void ShouldReplaceWithNullIfDbNullTranspose()
        {
            _table.Rows.Add("1", null);
            _worksheet.Cells["B2"].Value = 2;
            _worksheet.Cells["A1"].LoadFromDataTable(_table, c =>
            {
                c.PrintHeaders = true;
                c.TableStyle = TableStyles.Dark1;
                c.Transpose = true;
            });
            Assert.IsNull(_worksheet.Cells["B2"].Value);
        }
    }
}
