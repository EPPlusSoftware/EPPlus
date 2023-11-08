using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromCollectionDictionaryPropertyTests
    {
        [EpplusTable]
        public class TestClass
        {
            [EpplusTableColumn(Order = 2)] 
            public string Name { get; set; }

            [EPPlusDictionaryColumn(Order = 1, ColumnHeaders = new string[] { "A", "B", "C"})]
            public Dictionary<string, object> Columns { get; set; } 
        }

        [TestMethod]
        public void ShouldReadColumnsAndValuesFromDictionaryProperty()
        {
            var item1 = new TestClass
            {
                Name = "test 1",
                Columns = new Dictionary<string, object> { { "A", 1 }, { "B", 2 } }
            };
            var items = new List<TestClass> { item1 };
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(items, c => c.PrintHeaders = true);
                Assert.AreEqual("A", sheet.Cells["A1"].Value);
                Assert.AreEqual("B", sheet.Cells["B1"].Value);
                Assert.AreEqual("C", sheet.Cells["C1"].Value);
                Assert.AreEqual("Name", sheet.Cells["D1"].Value);
                Assert.AreEqual(1, sheet.Cells["A2"].Value);
                Assert.AreEqual(2, sheet.Cells["B2"].Value);
                Assert.IsNull(sheet.Cells["C2"].Value);
                Assert.AreEqual("test 1", sheet.Cells["D2"].Value);
            }

        }
    }
}
