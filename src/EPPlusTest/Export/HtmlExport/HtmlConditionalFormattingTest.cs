using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.ConditionalFormatting;
using System.Drawing;

namespace EPPlusTest.Export.HtmlExport
{
    [TestClass]
    public class HtmlConditionalFormattingTest : TestBase
    {
        [TestMethod]
        public void ExportingTableFileShouldWork()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("noStyleWs");
                var range = sheet.Cells["A1:D5"];
                var table = sheet.Tables.Add(range, "noStyleRange");

                sheet.Cells["B1:D5"].Formula = "ROW()";
                sheet.Cells["B3"].Formula = "0";


                var cf = sheet.Cells["B2:D5"].ConditionalFormatting.AddThreeColorScale();

                cf.LowValue.Type = eExcelConditionalFormattingValueObjectType.Min;
                cf.HighValue.Type = eExcelConditionalFormattingValueObjectType.Max;
                cf.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;

                cf.LowValue.Color = Color.Teal;
                cf.HighValue.Color = Color.Green;
                cf.MiddleValue.Color = Color.Blue;

                sheet.Calculate();

                var settings = new HtmlTableExportSettings();
                var context = new ExporterContext();
                context.InitializeQuadTree(range);

                var list = AttributeTranslator.GetClassAttributeFromStyle(sheet.Cells["B3"], false, settings, string.Empty, context);

                Assert.AreEqual(2, list.Count);
                Assert.AreEqual(list[0], "epp-ar");
                Assert.AreEqual(list[1], "background-color:#"+ Color.Teal.ToArgb().ToString("x8").Substring(2)+";");
            }
        }
    }
}
