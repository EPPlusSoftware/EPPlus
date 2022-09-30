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
    public class ToCollectionTests
    {
        public struct Category
        {
            public int CatId { get; set; }
            public string Name { get; set; }
            public string Description { get; set; }
        }
        public class TestDto
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public double Ratio { get; set; }
            public Category Category { get; set; }
        }

        [TestMethod]
        public void ToCollection_RowList()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = "Name";
                sheet.Cells["C1"].Value = "Ratio";
                sheet.Cells["D1"].Value = "CategoryId";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = "John Doe";
                sheet.Cells["C2"].Value = 12.38;
                sheet.Cells["D2"].Value = 1;
                sheet.Cells["A3"].Value = 2;
                sheet.Cells["B3"].Value = "Jane Doe";
                sheet.Cells["C3"].Value = 68.44;
                sheet.Cells["D3"].Value = 3;

                var list = sheet.Cells["A2:D3"].ToCollection((List<object> l) => 
                {
                    var dto = new TestDto();
                    dto.Id = (int)l[0];
                    dto.Name = l[1].ToString();
                    dto.Ratio = (double)l[2];
                    dto.Category = new Category() { CatId = (int)l[3] };
                    return dto;
                });

                Assert.AreEqual(2, list.Count);
                Assert.AreEqual(sheet.Cells["A2"].Value, list[0].Id);
                Assert.AreEqual(sheet.Cells["B2"].Text, list[0].Name);
                Assert.AreEqual(sheet.Cells["C2"].Value, list[0].Ratio);
                Assert.AreEqual(sheet.Cells["D2"].Value, list[0].Category.CatId);

                Assert.AreEqual(sheet.Cells["A3"].Value, list[1].Id);
                Assert.AreEqual(sheet.Cells["B3"].Text, list[1].Name);
                Assert.AreEqual(sheet.Cells["C3"].Value, list[1].Ratio);
                Assert.AreEqual(sheet.Cells["D3"].Value, list[1].Category.CatId);
            }
        }
    }
}
