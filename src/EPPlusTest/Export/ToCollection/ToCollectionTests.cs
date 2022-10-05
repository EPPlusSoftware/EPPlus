using EPPlusTest.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToCollection;
using OfficeOpenXml.Export.ToDataTable;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
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
    public class ToCollectionTests : TestBase
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
            public DateTime TimeStamp { get; set; }
            public Category Category { get; set; }
            public string FormattedRatio { get; set; }
            public string FormattedTimeStamp { get; set; }
        }

        [TestMethod]
        public void ToCollection_Index()
        {
            using(var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionIndex");
                var list = sheet.Cells["A2:E3"].ToCollection(
                    row => 
                    {
                        var dto = new TestDto();
                        dto.Id = row.GetValue<int>(0);
                        dto.Name = row.GetValue<string>(1);
                        dto.Ratio = row.GetValue<double>(2);
                        dto.TimeStamp = row.GetValue<DateTime>(3);
                        dto.Category = new Category() { CatId = row.GetValue<int>(4) };
                        dto.FormattedRatio = row.GetText(2);
                        dto.FormattedTimeStamp = row.GetText(3);
                        return dto;
                    });

                Assert.AreEqual(2, list.Count);
                Assert.AreEqual(sheet.Cells["A2"].Value, list[0].Id);
                Assert.AreEqual(sheet.Cells["B2"].Text, list[0].Name);
                Assert.AreEqual(sheet.Cells["C2"].Value, list[0].Ratio);
                Assert.AreEqual(sheet.Cells["D2"].Value, list[0].TimeStamp);
                Assert.AreEqual(sheet.Cells["E2"].Value, list[0].Category.CatId);
                Assert.AreEqual(sheet.Cells["C2"].Text, list[0].FormattedRatio);
                Assert.AreEqual(sheet.Cells["D2"].Text, list[0].FormattedTimeStamp);

                Assert.AreEqual(sheet.Cells["A3"].Value, list[1].Id);
                Assert.AreEqual(sheet.Cells["B3"].Text, list[1].Name);
                Assert.AreEqual(sheet.Cells["C3"].Value, list[1].Ratio);
                Assert.AreEqual(sheet.Cells["D3"].Value, list[1].TimeStamp);
                Assert.AreEqual(sheet.Cells["E3"].Value, list[1].Category.CatId);                
                Assert.AreEqual(sheet.Cells["C3"].Text, list[1].FormattedRatio);
                Assert.AreEqual(sheet.Cells["D3"].Text, list[1].FormattedTimeStamp);
            }
        }

        [TestMethod]
        public void ToCollection_ColumnNames()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionName");
                var list = sheet.Cells["A1:E3"].ToCollection(x =>
                {
                    var dto = new TestDto();
                    dto.Id = x.GetValue<int>("id");
                    dto.Name = x.GetValue<string>("Name");
                    dto.Ratio = x.GetValue<double>("Ratio");
                    dto.TimeStamp = x.GetValue<DateTime>("TimeStamp");
                    dto.Category = new Category() { CatId = x.GetValue<int>("CategoryId") };
                    dto.FormattedRatio = x.GetText("Ratio");
                    dto.FormattedTimeStamp = x.GetText("TimeStamp");
                    return dto;
                }, x => x.HeaderRow = 0);

                Assert.AreEqual(2, list.Count);
                Assert.AreEqual(sheet.Cells["A2"].Value, list[0].Id);
                Assert.AreEqual(sheet.Cells["B2"].Text, list[0].Name);
                Assert.AreEqual(sheet.Cells["C2"].Value, list[0].Ratio);
                Assert.AreEqual(sheet.Cells["D2"].Value, list[0].TimeStamp);
                Assert.AreEqual(sheet.Cells["E2"].Value, list[0].Category.CatId);
                Assert.AreEqual(sheet.Cells["C3"].Text, list[0].FormattedRatio);
                Assert.AreEqual(sheet.Cells["D3"].Text, list[0].FormattedTimeStamp);

                Assert.AreEqual(sheet.Cells["A3"].Value, list[1].Id);
                Assert.AreEqual(sheet.Cells["B3"].Text, list[1].Name);
                Assert.AreEqual(sheet.Cells["C3"].Value, list[1].Ratio);
                Assert.AreEqual(sheet.Cells["D3"].Value, list[1].TimeStamp);
                Assert.AreEqual(sheet.Cells["E3"].Value, list[1].Category.CatId);
                
                Assert.AreEqual(sheet.Cells["C3"].Text, list[1].FormattedRatio);
                Assert.AreEqual(sheet.Cells["D3"].Text, list[1].FormattedTimeStamp);
            }
        }
        [TestMethod]
        public void ToCollection_CustomHeaders()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionName");
                var list = sheet.Cells["A2:E3"].ToCollection((ToCollectionRow row) =>
                {
                    var dto = new TestDto();
                    dto.Id = row.GetValue<int>("The Id");                    
                    dto.Name = row.GetValue<string>("The Name");
                    dto.Ratio = row.GetValue<double>("The Ratio");
                    dto.TimeStamp = row.GetValue<DateTime>("End Date");                    
                    dto.Category = new Category() { CatId = row.GetValue<int>("Category Id") };
                    return dto;
                }, x => x.SetCustomHeaders("The Id", "The Name", "The Ratio", "End Date", "Category Id"));

                Assert.AreEqual(2, list.Count);
                Assert.AreEqual(sheet.Cells["A2"].Value, list[0].Id);
                Assert.AreEqual(sheet.Cells["B2"].Text, list[0].Name);
                Assert.AreEqual(sheet.Cells["C2"].Value, list[0].Ratio);
                Assert.AreEqual(sheet.Cells["D2"].GetValue<DateTime>(), list[0].TimeStamp);
                Assert.AreEqual(sheet.Cells["E2"].Value, list[0].Category.CatId);

                Assert.AreEqual(sheet.Cells["A3"].Value, list[1].Id);
                Assert.AreEqual(sheet.Cells["B3"].Text, list[1].Name);
                Assert.AreEqual(sheet.Cells["C3"].Value, list[1].Ratio);
                Assert.AreEqual(sheet.Cells["D3"].GetValue<DateTime>(), list[1].TimeStamp);
                Assert.AreEqual(sheet.Cells["E3"].Value, list[1].Category.CatId);
            }
        }

        private ExcelWorksheet LoadTestData(ExcelPackage p, string wsName)
        {
            var sheet = p.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Id";
            sheet.Cells["B1"].Value = "Name";
            sheet.Cells["C1"].Value = "Ratio";
            sheet.Cells["D1"].Value = "TimeStamp";
            sheet.Cells["E1"].Value = "CategoryId";
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["B2"].Value = "John Doe";
            sheet.Cells["C2"].Value = 1012.38;
            sheet.Cells["D2"].Value = new DateTime(2022,10,1,13,15,30);
            sheet.Cells["E2"].Value = 1;
            sheet.Cells["A3"].Value = 2;
            sheet.Cells["B3"].Value = "Jane Doe";
            sheet.Cells["C3"].Value = 9968.44;
            sheet.Cells["D3"].Value = new DateTime(2022, 11, 1, 18, 15, 30);
            sheet.Cells["E3"].Value = 3;
            sheet.Cells["C2:C3"].Style.Numberformat.Format = "#,##0.0";
            sheet.Cells["D2:D3"].Style.Numberformat.Format = "yyyy-MM-dd HH:MM";
            return sheet;

        }
    }
}
