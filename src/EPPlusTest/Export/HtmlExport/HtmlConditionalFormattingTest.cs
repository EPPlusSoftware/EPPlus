using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.ConditionalFormatting;
using System.Drawing;
using System.IO;
using System.Text;

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

                var classString = AttributeTranslator.GetClassAttributeFromStyle(sheet.Cells["B3"], false, settings, string.Empty, context);
                var stylesAndExtras = AttributeTranslator.GetConditionalFormattings(sheet.Cells["B3"], settings, context, ref classString);

                Assert.AreEqual("epp-ar", classString);
                var expectedString = "background-color:#" + Color.Teal.ToArgb().ToString("x8").Substring(2) + ";";
                Assert.AreEqual(expectedString, stylesAndExtras[0]);
            }
        }
        [TestMethod]
        public void ExportingHtmlTemplate()
        {
            using (var package = OpenTemplatePackage("CF_IconSetsCompareTemplate.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];

                //var model = new ExportViewModel();
                var exporter = ws.Cells["A1:AC108"].CreateHtmlExporter();

                var settings = exporter.Settings;
                settings.Pictures.Include = ePictureInclude.Include;
                //settings.Pictures.KeepOriginalSize = true;
                settings.Minify = false;
                settings.SetColumnWidth = true;
                settings.SetRowHeight = true;
                settings.Pictures.AddNameAsId = true;

                //var Css = exporter.GetCssString();
                //var Html = exporter.GetHtmlString();

                // Create the file, or overwrite if the file exists.
                using (FileStream fs = File.Create("C:\\epplusTest\\Testoutput\\CF_IconSetsCompareTemplate.html"))
                {
                    byte[] info = new UTF8Encoding(true).GetBytes(exporter.GetSinglePage());
                    // Add some information to the file.
                    fs.Write(info, 0, info.Length);
                }
            }
        }
    }
}
