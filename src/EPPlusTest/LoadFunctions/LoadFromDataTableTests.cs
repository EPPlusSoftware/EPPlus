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
                sheet.Cells["A1"].LoadFromDataTable(table, true);
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
        public void ShouldUseCaptionForHeader()
        {
            _table.Columns["Id"].Caption = "An Id";
            _table.Columns["Name"].Caption = "A name";
            _worksheet.Cells["A1"].LoadFromDataTable(_table, true);
            Assert.AreEqual("An Id", _worksheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldUseColumnNameForHeaderIfNoCaption()
        {
            _worksheet.Cells["A1"].LoadFromDataTable(_table, true);
            Assert.AreEqual("Id", _worksheet.Cells["A1"].Value);
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
                sheet.Cells["A1"].LoadFromDataTable(dataSet.Tables["Astronaut"], true, TableStyles.Dark1);
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
    }
}
