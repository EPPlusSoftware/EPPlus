using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.LoadFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FakeItEasy;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromCollectionDictionaryPropertyTests
    {
        [EpplusTable]
        public class TestClass
        {
            [EpplusTableColumn(Order = 3)] 
            public string Name { get; set; }

            [EPPlusDictionaryColumn(Order = 2, KeyId = "1")]
            public Dictionary<string, object> Columns { get; set; }
        }

        public class TestClass2 : TestClass
        {
            [EPPlusDictionaryColumn(Order = 1, KeyId = "2")]
            public Dictionary<string, object> Columns2 { get; set; }
        }

        [EpplusTable]
        public class  TestClass3
        {
            [EpplusTableColumn(Order = 1)]
            public string Name { get; set; }

            [EPPlusDictionaryColumn(Order = 2)]
            public Dictionary<string, object> Columns { get; set; }
        }

        [EpplusTable]
        public class TestClass4
        {
            [EpplusTableColumn(Order = 1)]
            public string Name { get; set; }

            [EPPlusDictionaryColumn(Order = 2, ColumnHeaders = new string[] { "C", "B", "A" })]
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
            var keys1 = new string[] { "A", "B", "C" };
            var keys2 = new string[] { "C", "D", "E" };
            var items = new List<TestClass> { item1 };
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.RegisterDictionaryKeys("1", keys1);
                    c.RegisterDictionaryKeys("2", keys2);
                });
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

        [TestMethod]
        public void ShouldReadColumnsAndValuesFromDictionaryProperty2()
        {
            var item1 = new TestClass2
            {
                Name = "test 1",
                Columns = new Dictionary<string, object> { { "A", 3 } },
                Columns2 = new Dictionary<string, object> { { "C", 1 }, { "D", 2 } }
            };
            var keys1 = new string[] { "A", "B", "C" };
            var keys2 = new string[] { "C", "D", "E" };
            var items = new List<TestClass2> { item1 };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.RegisterDictionaryKeys("1", keys1);
                    c.RegisterDictionaryKeys("2", keys2);
                });
                Assert.AreEqual("C", sheet.Cells["A1"].Value);
                Assert.AreEqual("D", sheet.Cells["B1"].Value);
                Assert.AreEqual("E", sheet.Cells["C1"].Value);
                Assert.AreEqual("A", sheet.Cells["D1"].Value);
                Assert.AreEqual("B", sheet.Cells["E1"].Value);
                Assert.AreEqual("C", sheet.Cells["F1"].Value);
                Assert.AreEqual("Name", sheet.Cells["G1"].Value);
                Assert.AreEqual(1, sheet.Cells["A2"].Value);
                Assert.AreEqual(2, sheet.Cells["B2"].Value);
                Assert.IsNull(sheet.Cells["C2"].Value);
                Assert.AreEqual(3, sheet.Cells["D2"].Value);
                Assert.IsNull(sheet.Cells["E2"].Value);
                Assert.AreEqual("test 1", sheet.Cells["G2"].Value);
            }
        }

        [TestMethod]
        public void ShouldReadColumnsAndValuesFromDictionaryProperty3()
        {
            var item1 = new TestClass2
            {
                Name = "test 1",
                Columns = new Dictionary<string, object> { { "A", 3 } },
                Columns2 = new Dictionary<string, object> { { "C", 1 }, { "D", 2 } }
            };
            var keys1 = new string[] { "C", "B", "A" };
            var keys2 = new string[] { "C", "D", "E" };
            var items = new List<TestClass2> { item1 };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.RegisterDictionaryKeys("1", keys1);
                    c.RegisterDictionaryKeys("2", keys2);
                });
                Assert.AreEqual("C", sheet.Cells["A1"].Value);
                Assert.AreEqual("D", sheet.Cells["B1"].Value);
                Assert.AreEqual("E", sheet.Cells["C1"].Value);
                Assert.AreEqual("C", sheet.Cells["D1"].Value);
                Assert.AreEqual("B", sheet.Cells["E1"].Value);
                Assert.AreEqual("A", sheet.Cells["F1"].Value);
                Assert.AreEqual("Name", sheet.Cells["G1"].Value);
                Assert.AreEqual(1, sheet.Cells["A2"].Value);
                Assert.AreEqual(2, sheet.Cells["B2"].Value);
                Assert.IsNull(sheet.Cells["C2"].Value);
                Assert.IsNull(sheet.Cells["E2"].Value);
                Assert.AreEqual(3, sheet.Cells["F2"].Value);
                Assert.AreEqual("test 1", sheet.Cells["G2"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseDefaultKeys()
        {
            var item1 = new TestClass3
            {
                Name = "test 1",
                Columns = new Dictionary<string, object> { { "A", 3 }, { "B", 2 } }
            };
            var keys1 = new string[] { "C", "B", "A" };
            var items = new List<TestClass3> { item1 };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.RegisterDictionaryKeys(keys1);
                });
                Assert.AreEqual("Name", sheet.Cells["A1"].Value);
                Assert.AreEqual("C", sheet.Cells["B1"].Value);
                Assert.AreEqual("B", sheet.Cells["C1"].Value);
                Assert.AreEqual("A", sheet.Cells["D1"].Value);
                Assert.AreEqual("test 1", sheet.Cells["A2"].Value);
                Assert.IsNull(sheet.Cells["B2"].Value);
                Assert.AreEqual(2, sheet.Cells["C2"].Value);
                Assert.AreEqual(3, sheet.Cells["D2"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseHeadersFromAttribute()
        {
            var item1 = new TestClass4
            {
                Name = "test 1",
                Columns = new Dictionary<string, object> { { "A", 3 }, { "B", 2 } }
            };
            var items = new List<TestClass4> { item1 };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                });
                Assert.AreEqual("Name", sheet.Cells["A1"].Value);
                Assert.AreEqual("C", sheet.Cells["B1"].Value);
                Assert.AreEqual("B", sheet.Cells["C1"].Value);
                Assert.AreEqual("A", sheet.Cells["D1"].Value);
                Assert.AreEqual("test 1", sheet.Cells["A2"].Value);
                Assert.IsNull(sheet.Cells["B2"].Value);
                Assert.AreEqual(2, sheet.Cells["C2"].Value);
                Assert.AreEqual(3, sheet.Cells["D2"].Value);
            }
        }
    }
}
