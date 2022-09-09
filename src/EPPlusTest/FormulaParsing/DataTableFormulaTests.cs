using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class DataTableFormulaTests : TestBase
    {
        [TestMethod]
        public void CheckSaveWhatif_DataTable()
        {
            using (var p = OpenTemplatePackage("Whatif-DataTable.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                Assert.AreEqual(4900D, ws.Cells["F5"].Value);
                Assert.AreEqual(2900D, ws.Cells["T20"].Value);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void CheckSaveWhatif_CopyWorksheetInsertRow()
        {
            using (var p = OpenTemplatePackage("Whatif-DataTable.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var copy = p.Workbook.Worksheets.Add("Copy", ws);
                copy.InsertRow(2, 1);
                copy.InsertRow(7, 1);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void CheckSaveWhatif_InsertInsideRow()
        {
            using (var p = OpenTemplatePackage("Whatif-DataTable.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var copy = p.Workbook.Worksheets.Add("Copy", ws);
                copy.InsertRow(3, 1);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void CheckSaveWhatif_CopyWorksheetInsertColumn()
        {
            using (var p = OpenTemplatePackage("Whatif-DataTable.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var copy = p.Workbook.Worksheets.Add("Copy", ws);
                copy.InsertColumn(2, 1);
                copy.InsertColumn(8, 1);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void CheckSaveWhatif_InsertInsideColumn()
        {
            using (var p = OpenTemplatePackage("Whatif-DataTable.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var copy = p.Workbook.Worksheets.Add("Copy", ws);
                copy.InsertColumn(4, 1);
                SaveAndCleanup(p);
            }
        }

    }
}
