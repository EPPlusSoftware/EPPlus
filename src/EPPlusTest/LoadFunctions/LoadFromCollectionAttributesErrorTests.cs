using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.LoadFunctions
{
    public class NestedAttrOnString
    {
        [EpplusTableColumn(Order = 1)]
        public string Name { get; set; }

        [EpplusNestedTableColumn(Order = 2)]
        public string Value { get; set; }
    }

    [TestClass]
    public class LoadFromCollectionAttributesErrorTests
    {
        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void NestedTypeAttributeOnString()
        {
            var coll = new List<NestedAttrOnString>
            {
                new NestedAttrOnString
                {
                    Name = "foo",
                    Value = "bar",
                }
            };
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(coll);
            }
        }
    }
}
