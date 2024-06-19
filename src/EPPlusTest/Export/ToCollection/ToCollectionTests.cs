using EPPlusTest.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Export.ToCollection;
using OfficeOpenXml.Export.ToCollection.Exceptions;
using OfficeOpenXml.Export.ToDataTable;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace EPPlusTest.Export.ToCollection
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
            [DisplayName("Identity")]            
            public int Id { get; set; }
            [EpplusTableColumn(Order = 1, Header = "First name")]
            public string Name { get; set; }
            public double Ratio { get; set; }
            public DateTime TimeStamp { get; set; }
            public Category Category { get; set; }
            public string FormattedRatio { get; set; }
            public string FormattedTimeStamp { get; set; }
        }
#region Range
        [TestMethod]
        public void ToCollection_Index()
        {
            using(var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionIndex");
                var list = sheet.Cells["A2:E3"].ToCollectionWithMappings(
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
                var list = sheet.Cells["A1:E3"].ToCollectionWithMappings(x =>
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
        public void ToCollection_AutoMapInMapping()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionName");
                var list = sheet.Cells["A1:E3"].ToCollectionWithMappings((ToCollectionRow row) =>
                {
                    var dto = new TestDto();
                    row.Automap(dto);
                    dto.Category = new Category() { CatId = row.GetValue<int>("CategoryId") };
                    dto.FormattedRatio = row.GetText("Ratio");
                    dto.FormattedTimeStamp = row.GetText("TimeStamp");
                    return dto;
                }, x => x.HeaderRow=0);

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
        [TestMethod]
        public void ToCollection_CustomHeaders()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionHeaders");
                var list = sheet.Cells["A2:E3"].ToCollectionWithMappings((ToCollectionRow row) =>
                {
                    var dto = new TestDto();
                    dto.Id = row.GetValue<int>("Custom-Id");
                    dto.Name = row.GetText("Custom-Name");
                    dto.Category = new Category() { CatId = row.GetValue<int>("Custom-CategoryId") };
                    dto.Ratio = row.GetValue<double>("Custom-Ratio");
                    dto.FormattedRatio = row.GetText("Custom-Ratio");
                    dto.FormattedTimeStamp = row.GetText("Custom-TimeStamp");
                    dto.TimeStamp = row.GetValue<DateTime>("Custom-TimeStamp");
                    return dto;
                }, x => x.SetCustomHeaders("Custom-Id", "Custom-Name", "Custom-Ratio", "Custom-TimeStamp", "Custom-CategoryId"));

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
#if(NET40_OR_GREATER)
        [TestMethod]
        public void ToCollection_AutoMap()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionAuto");
                sheet.Cells["A1"].Value = "Identity";
                sheet.Cells["B1"].Value = "First name";
                var list = sheet.Cells["A1:E3"].ToCollection<TestDto>();

                Assert.AreEqual(2, list.Count);
                Assert.AreEqual(sheet.Cells["A2"].Value, list[0].Id);
                Assert.AreEqual(sheet.Cells["B2"].Text, list[0].Name);
                Assert.AreEqual(sheet.Cells["C2"].Value, list[0].Ratio);
                Assert.AreEqual(sheet.Cells["D2"].Value, list[0].TimeStamp);

                Assert.AreEqual(sheet.Cells["A3"].Value, list[1].Id);
                Assert.AreEqual(sheet.Cells["B3"].Text, list[1].Name);
                Assert.AreEqual(sheet.Cells["C3"].Value, list[1].Ratio);
                Assert.AreEqual(sheet.Cells["D3"].Value, list[1].TimeStamp);
            }
        }
        [TestMethod]
        public void ToCollection_AutoMapInCallback()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionAuto");
                sheet.Cells["A1"].Value = "Identity";
                sheet.Cells["B1"].Value = "First name";
                var list = sheet.Cells["A1:E3"].ToCollectionWithMappings(x =>
                {
                    var item = new TestDto();
                    x.Automap(item); 
                    item.Category = new Category() { CatId = x.GetValue<int>("CategoryId") };
                    item.FormattedRatio = x.GetText("Ratio");
                    item.FormattedTimeStamp = x.GetText("TimeStamp");
                    return item;
                }, x=>x.HeaderRow=0);

                Assert.AreEqual(2, list.Count);
                Assert.AreEqual(sheet.Cells["A2"].Value, list[0].Id);
                Assert.AreEqual(sheet.Cells["B2"].Text, list[0].Name);
                Assert.AreEqual(sheet.Cells["C2"].Value, list[0].Ratio);
                Assert.AreEqual(sheet.Cells["D2"].Value, list[0].TimeStamp);

                Assert.AreEqual(sheet.Cells["A3"].Value, list[1].Id);
                Assert.AreEqual(sheet.Cells["B3"].Text, list[1].Name);
                Assert.AreEqual(sheet.Cells["C3"].Value, list[1].Ratio);
                Assert.AreEqual(sheet.Cells["D3"].Value, list[1].TimeStamp);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(EPPlusDataTypeConvertionException))]
        public void ToCollection_AutoMap_EPPlusDataTypeConvertionException()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionAuto");
                sheet.Cells["A1"].Value = "Identity";
                sheet.Cells["B1"].Value = "First name";
                sheet.Cells["A2"].Value = "Error";
                var list = sheet.Cells["A1:E3"].ToCollectionWithMappings(x =>
                {
                    var item = new TestDto();
                    x.Automap(item);
                    item.Category = new Category() { CatId = x.GetValue<int>("CategoryId") };
                    item.FormattedRatio = x.GetText("Ratio");
                    item.FormattedTimeStamp = x.GetText("TimeStamp");
                    return item;
                }, x => x.HeaderRow = 0);
            }
        }
        [TestMethod]
        public void ToCollection_AutoMapInCallback_SetDefaultValueOnConversionError()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionAuto");
                sheet.Cells["A1"].Value = "Identity";
                sheet.Cells["B1"].Value = "First name";
                sheet.Cells["A2"].Value = "Error";
                var list = sheet.Cells["A1:E3"].ToCollectionWithMappings(x =>
                {
                    var item = new TestDto();
                    x.Automap(item);
                    item.Category = new Category() { CatId = x.GetValue<int>("CategoryId") };
                    item.FormattedRatio = x.GetText("Ratio");
                    item.FormattedTimeStamp = x.GetText("TimeStamp");
                    return item;
                }, x => { x.HeaderRow = 0; x.ConversionFailureStrategy = ToCollectionConversionFailureStrategy.SetDefaultValue; });

                Assert.AreEqual(2, list.Count);
                Assert.AreEqual(default(int), list[0].Id); //sheet.Cells["A2"].Value is invalid for int as it contains a string
                Assert.AreEqual(sheet.Cells["B2"].Text, list[0].Name);
                Assert.AreEqual(sheet.Cells["C2"].Value, list[0].Ratio);
                Assert.AreEqual(sheet.Cells["D2"].Value, list[0].TimeStamp);

                Assert.AreEqual(sheet.Cells["A3"].Value, list[1].Id);
                Assert.AreEqual(sheet.Cells["B3"].Text, list[1].Name);
                Assert.AreEqual(sheet.Cells["C3"].Value, list[1].Ratio);
                Assert.AreEqual(sheet.Cells["D3"].Value, list[1].TimeStamp);
            }
        }

        [TestMethod]
        public void ToCollection_Transposed()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["C1"].Value = 2;
                sheet.Cells["D1"].Value = 3;
                sheet.Cells["E1"].Value = 4;
                sheet.Cells["F1"].Value = 5;
                sheet.Cells["G1"].Value = 6;
                sheet.Cells["A2"].Value = "Name";
                sheet.Cells["B2"].Value = "Scott";
                sheet.Cells["C2"].Value = "Mats";
                sheet.Cells["D2"].Value = "Jimmy";
                sheet.Cells["E2"].Value = "Cameron";
                sheet.Cells["F2"].Value = "Luther";
                sheet.Cells["G2"].Value = "Josh";

                var list = sheet.Cells["A1:G2"].ToCollectionWithMappings((ToCollectionRow row) =>
                {
                    var dto = new TestDto();
                    dto.Id = row.GetValue<int>("Id");
                    dto.Name = row.GetText("Name");
                    return dto;
                }, x => {
                            x.DataIsTransposed = true;
                            x.HeaderRow = 0;
                        }
                );

                Assert.AreEqual(6, list.Count);
                Assert.AreEqual(sheet.Cells["B1"].Value, list[0].Id);
                Assert.AreEqual(sheet.Cells["B2"].Value, list[0].Name);
                Assert.AreEqual(sheet.Cells["G1"].Value, list[5].Id);
                Assert.AreEqual(sheet.Cells["G2"].Value, list[5].Name);
            }
        }


#endif
        #endregion
        #region Table
        [TestMethod]
        public void ToCollectionTable_AutoMap()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionAuto", true);
                sheet.Cells["A1"].Value = "Identity";
                sheet.Cells["B1"].Value = "First name";
                var list = sheet.Tables[0].ToCollection<TestDto>();

                Assert.AreEqual(2, list.Count);
                Assert.AreEqual(sheet.Cells["A2"].Value, list[0].Id);
                Assert.AreEqual(sheet.Cells["B2"].Text, list[0].Name);
                Assert.AreEqual(sheet.Cells["C2"].Value, list[0].Ratio);
                Assert.AreEqual(sheet.Cells["D2"].Value, list[0].TimeStamp);

                Assert.AreEqual(sheet.Cells["A3"].Value, list[1].Id);
                Assert.AreEqual(sheet.Cells["B3"].Text, list[1].Name);
                Assert.AreEqual(sheet.Cells["C3"].Value, list[1].Ratio);
                Assert.AreEqual(sheet.Cells["D3"].Value, list[1].TimeStamp);
            }
        }

        [TestMethod]
        public void ToCollectionTable_Index()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionIndex", true);
                var list = sheet.Tables[0].ToCollection(
                    row =>
                    {
                        var dto = new TestDto();
                        dto.Id = row.GetValue<int>(0);
                        dto.Name = row.GetValue<string>(1);
                        dto.Ratio = row.GetValue<double>(2);
                        dto.TimeStamp = row.GetValue<DateTime > (3);
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
        public void ToCollectionTable_AutoMapInCallback()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionAuto", true);
                sheet.Cells["A1"].Value = "Identity";
                sheet.Cells["B1"].Value = "First name";
                var list = sheet.Tables[0].ToCollection(x =>
                {
                    var item = new TestDto();
                    x.Automap(item);
                    item.Category = new Category() { CatId = x.GetValue<int>("CategoryId") };
                    item.FormattedRatio = x.GetText("Ratio");
                    item.FormattedTimeStamp = x.GetText("TimeStamp");
                    return item;
                });

                Assert.AreEqual(2, list.Count);
                Assert.AreEqual(sheet.Cells["A2"].Value, list[0].Id);
                Assert.AreEqual(sheet.Cells["B2"].Text, list[0].Name);
                Assert.AreEqual(sheet.Cells["C2"].Value, list[0].Ratio);
                Assert.AreEqual(sheet.Cells["D2"].Value, list[0].TimeStamp);

                Assert.AreEqual(sheet.Cells["A3"].Value, list[1].Id);
                Assert.AreEqual(sheet.Cells["B3"].Text, list[1].Name);
                Assert.AreEqual(sheet.Cells["C3"].Value, list[1].Ratio);
                Assert.AreEqual(sheet.Cells["D3"].Value, list[1].TimeStamp);
            }
        }
        [TestMethod]
        public void ToCollectionTable_ColumnNames()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionName",true);
                var list = sheet.Tables[0].ToCollection(x =>
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
        [ExpectedException(typeof(EPPlusDataTypeConvertionException))]
        public void ToCollectionTable_ColumnNamesConversionFailure()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionName", true);
                sheet.Cells["C2"].Value = "str";
                var list = sheet.Tables[0].ToCollection(x =>
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
                });
            }
        }
        [TestMethod]
        public void ToCollectionTable_ColumnNamesConversionDefault()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = LoadTestData(p, "LoadFromCollectionName", true);
                sheet.Cells["C2"].Value = "str";
                var list = sheet.Tables[0].ToCollectionWithMappings(x =>
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
                }, x => x.ConversionFailureStrategy = ToCollectionConversionFailureStrategy.SetDefaultValue);
                
                Assert.AreEqual(2, list.Count);
                Assert.AreEqual(sheet.Cells["A2"].Value, list[0].Id);
                Assert.AreEqual(sheet.Cells["B2"].Text, list[0].Name);
                Assert.AreEqual(0D,list[0].Ratio);
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

        #endregion
        private ExcelWorksheet LoadTestData(ExcelPackage p, string wsName, bool addTable = false)
        {
            var sheet = p.Workbook.Worksheets.Add(wsName);
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
            if(addTable)
            {
                var t=sheet.Tables.Add(sheet.Cells["A1:E3"], $"tbl{wsName}");
                t.ShowTotal = true;
                t.Columns["Id"].TotalsRowLabel = "Totals";
                t.Columns["Ratio"].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
                t.Columns["TimeStamp"].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Count;
            }
            return sheet;

        }
    }
}
