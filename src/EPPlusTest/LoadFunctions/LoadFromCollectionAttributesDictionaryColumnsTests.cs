using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Table;

namespace EPPlusTest.LoadFunctions
{
    [EpplusTable(PrintHeaders = true, TableStyle = TableStyles.Medium1)]
    public class DictionaryColumnsTestClass
    {
        [EpplusTableColumn(Order = 1)]
        public string Name { get; set; }

        [EpplusDictionaryColumns(Order = 2, Keys = new string[] { "Q1", "Q2", "Q3", "Q4" })]
        public Dictionary<string, int> Quarters { get; set; } = new Dictionary<string, int>();
    }
    [TestClass]
    public class LoadFromCollectionAttributesDictionaryColumnsTests
    {
        public void Test1()
        {
            var items = new List<DictionaryColumnsTestClass>();
            var d1 = new DictionaryColumnsTestClass();
            d1.Name = "Bob";
            d1.Quarters["Q1"] = 12;
            d1.Quarters["Q3"] = 45;
            items.Add(d1);
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].LoadFromCollection(items);
        }
    }
}
