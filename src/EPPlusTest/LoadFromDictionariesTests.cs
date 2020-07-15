using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest
{
    [TestClass]
    public class LoadFromDictionariesTests
    {
        [TestMethod]
        public void ShouldLoad()
        {
            var items = new List<IDictionary<string, object>>()
            {
                new Dictionary<string, object>()
                { 
                    { "Id", 1 },
                    { "Name", "TestName 1" }
                },
                new Dictionary<string, object>()
                {
                    { "Id", 2 },
                    { "Name", "TestName 2" }
                }
            };
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromDictionaries(items, true, TableStyles.None, null);
            }
        }
    }
}
