using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.ConditionalFormatting;
using System.Drawing;
using System.IO;
using System.Text;
using OfficeOpenXml.Drawing;

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

        [TestMethod]
        public void ExportingHtmlCFsWithThemeColor()
        {
            using (var p = OpenPackage("AdvancedCFsWithThemeColor.xlsx", true))
            {
                var sheet = p.Workbook.Worksheets.Add("ConditionalFormattingSheet");

                var range = sheet.Cells["A1:A30"];
                var range2 = sheet.Cells["B1:B30"];
                var range3 = sheet.Cells["C1:C30"];

                sheet.Cells["A1:B30"].Formula = "ROW()";
                sheet.Cells["C1:C30"].Formula = "ROW()-10";

                sheet.Calculate();

                var twoColor = range.ConditionalFormatting.AddTwoColorScale();
                var threeColor = range2.ConditionalFormatting.AddThreeColorScale();
                var databar = range3.ConditionalFormatting.AddDatabar(Color.Aqua);

                twoColor.LowValue.ColorSettings.SetColor(eThemeSchemeColor.Accent4);
                twoColor.HighValue.ColorSettings.SetColor(eThemeSchemeColor.Accent6);

                threeColor.LowValue.ColorSettings.SetColor(eThemeSchemeColor.Accent1);
                threeColor.MiddleValue.ColorSettings.SetColor(eThemeSchemeColor.Text1);
                threeColor.HighValue.ColorSettings.SetColor(eThemeSchemeColor.Background2);

                databar.FillColor.SetColor(eThemeSchemeColor.Accent6);
                databar.BorderColor.SetColor(eThemeSchemeColor.Background2);
                databar.AxisColor.SetColor(eThemeSchemeColor.Accent2);
                databar.NegativeBorderColor.SetColor(eThemeSchemeColor.Accent4);
                databar.NegativeFillColor.SetColor(eThemeSchemeColor.Hyperlink);

                var exporter = sheet.Cells["A1:D30"].CreateHtmlExporter();

                var settings = exporter.Settings;
                settings.Pictures.Include = ePictureInclude.Include;

                settings.Minify = false;
                settings.SetColumnWidth = true;
                settings.SetRowHeight = true;
                settings.Pictures.AddNameAsId = true;

                var result = exporter.GetSinglePage();

                // Create the file, or overwrite if the file exists.
                using (FileStream fs = File.Create("C:\\epplusTest\\Testoutput\\CF_AdvancedThemeColorExport.html"))
                {
                    byte[] info = new UTF8Encoding(true).GetBytes(result);
                    // Add some information to the file.
                    fs.Write(info, 0, info.Length);
                }
                var expected = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\ntable.epplus-table{\r\n  font-family:Calibri;\r\n  font-size:11pt;\r\n  border-spacing:0;\r\n  border-collapse:collapse;\r\n  word-wrap:break-word;\r\n  white-space:nowrap;\r\n}\r\n.epp-hidden {\r\n  display:none;\r\n}\r\n.epp-al {\r\n  text-align:left;\r\n}\r\n.epp-ar {\r\n  text-align:right;\r\n}\r\n.epp-dcw {\r\n  width:64px;\r\n}\r\n.epp-drh {\r\n  height:20px;\r\n}\r\ntd.epp-image-cell {\r\n  vertical-align:middle;\r\n  text-align:center;\r\n}\r\n.epp-db-shared{\r\n  position:relative;\r\n  position:relative;\r\n  overflow:hidden;\r\n  background-image:url(data:image/svg+xml;base64,PHN2ZyB2ZXJzaW9uPScxLjEnIHZpZXdCb3g9JzAgMCAxNSAxMDAnIHhtbG5zPSdodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2Zyc+PGcgZmlsbD0nIzE0MDkwNCc+PHJlY3QgaWQ9J3N0cmlwZScgd2lkdGg9JzE1cHgnIGhlaWdodD0nNzUlJy8+PC9nPjwvc3ZnPg==);\r\n  background-size:5px 10px;\r\n  background-repeat:repeat-y;\r\n  background-position:-30px 0%;\r\n}\r\n.epp-db-shared::after{\r\n  content:\"\";\r\n  position:absolute;\r\n  width:100%;\r\n  height:calc(100% - 3px);\r\n  z-index:-1;\r\n  top:0%;\r\n  bottom:0%;\r\n  background-repeat:no-repeat;\r\n  background-size:100% 100%;\r\n}\r\n.epp-dxf1-pos, .epp-dxf1-neg{\r\n  z-index:0;\r\n  background-position:31.034% 0%;\r\n  background-image:url(data:image/svg+xml;base64,PHN2ZyB2ZXJzaW9uPScxLjEnIHZpZXdCb3g9JzAgMCAxNSAxMDAnIHhtbG5zPSdodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2Zyc+PGcgZmlsbD0nI2VkN2QzMSc+PHJlY3QgaWQ9J3N0cmlwZScgd2lkdGg9JzE1cHgnIGhlaWdodD0nNzUlJy8+PC9nPjwvc3ZnPg==);\r\n}\r\n.epp-dxf1-pos::after{\r\n  background-image:url(data:image/svg+xml;base64,PHN2ZyB2ZXJzaW9uPScxLjEnIHhtbG5zPSdodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZycgcHJlc2VydmVBc3BlY3RSYXRpbz0nbm9uZSc+PGRlZnM+PGxpbmVhckdyYWRpZW50IGlkPSdHcmFkaWVudDEnPjxzdG9wIGNsYXNzPSdzdG9wMScgb2Zmc2V0PScwJScgLz48c3RvcCBjbGFzcz0nc3RvcDInIG9mZnNldD0nOTAlJyAvPjwvbGluZWFyR3JhZGllbnQ+PHN0eWxlPiAjcmVjdDEgeyBmaWxsOiB1cmwoI0dyYWRpZW50MSk7IH0gLnN0b3AxIHsgc3RvcC1jb2xvcjogIzcwYWQ0NzsgfSAuc3RvcDIgeyBzdG9wLWNvbG9yOiB3aGl0ZTsgfSA8L3N0eWxlPjwvZGVmcz48cmVjdCBpZD0ncmVjdDEnIHdpZHRoPScxMDAlJyBoZWlnaHQ9JzEwMCUnIHN0cm9rZT0nI2U3ZTZlNicgc3Ryb2tlLXdpZHRoPScycHgnLz48L3N2Zz4=);\r\n  background-position:1px;\r\n  width:68.966%;\r\n  left:31.034%;\r\n}\r\n.epp-dxf1-neg::after{\r\n  background-image:url(data:image/svg+xml;base64,PHN2ZyB2ZXJzaW9uPScxLjEnIHhtbG5zPSdodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZycgcHJlc2VydmVBc3BlY3RSYXRpbz0nbm9uZSc+PGRlZnM+PGxpbmVhckdyYWRpZW50IGlkPSdHcmFkaWVudDEnPjxzdG9wIGNsYXNzPSdzdG9wMScgb2Zmc2V0PScwJScgLz48c3RvcCBjbGFzcz0nc3RvcDInIG9mZnNldD0nOTAlJyAvPjwvbGluZWFyR3JhZGllbnQ+PHN0eWxlPiAjcmVjdDEgeyBmaWxsOiB1cmwoI0dyYWRpZW50MSk7IH0gLnN0b3AxIHsgc3RvcC1jb2xvcjogIzA1NjNjMTsgfSAuc3RvcDIgeyBzdG9wLWNvbG9yOiB3aGl0ZTsgfSA8L3N0eWxlPjwvZGVmcz48cmVjdCBpZD0ncmVjdDEnIHdpZHRoPScxMDAlJyBoZWlnaHQ9JzEwMCUnIHN0cm9rZT0nI2ZmYzAwMCcgc3Ryb2tlLXdpZHRoPScycHgnLz48L3N2Zz4=);\r\n  background-position:1px;\r\n  width:31.034%;\r\n  right:68.966%;\r\n  transform:scale(-1, 1);\r\n}\r\n.epp-C1-db::after{\r\n  background-size:100% 100%;\r\n}\r\n.epp-C2-db::after{\r\n  background-size:88.889% 100%;\r\n}\r\n.epp-C3-db::after{\r\n  background-size:77.778% 100%;\r\n}\r\n.epp-C4-db::after{\r\n  background-size:66.667% 100%;\r\n}\r\n.epp-C5-db::after{\r\n  background-size:55.556% 100%;\r\n}\r\n.epp-C6-db::after{\r\n  background-size:44.444% 100%;\r\n}\r\n.epp-C7-db::after{\r\n  background-size:33.333% 100%;\r\n}\r\n.epp-C8-db::after{\r\n  background-size:22.222% 100%;\r\n}\r\n.epp-C9-db::after{\r\n  background-size:11.111% 100%;\r\n}\r\n.epp-C10-db::after{\r\n  background-size:0% 100%;\r\n}\r\n.epp-C11-db::after{\r\n  background-size:5% 100%;\r\n}\r\n.epp-C12-db::after{\r\n  background-size:10% 100%;\r\n}\r\n.epp-C13-db::after{\r\n  background-size:15% 100%;\r\n}\r\n.epp-C14-db::after{\r\n  background-size:20% 100%;\r\n}\r\n.epp-C15-db::after{\r\n  background-size:25% 100%;\r\n}\r\n.epp-C16-db::after{\r\n  background-size:30% 100%;\r\n}\r\n.epp-C17-db::after{\r\n  background-size:35% 100%;\r\n}\r\n.epp-C18-db::after{\r\n  background-size:40% 100%;\r\n}\r\n.epp-C19-db::after{\r\n  background-size:45% 100%;\r\n}\r\n.epp-C20-db::after{\r\n  background-size:50% 100%;\r\n}\r\n.epp-C21-db::after{\r\n  background-size:55% 100%;\r\n}\r\n.epp-C22-db::after{\r\n  background-size:60% 100%;\r\n}\r\n.epp-C23-db::after{\r\n  background-size:65% 100%;\r\n}\r\n.epp-C24-db::after{\r\n  background-size:70% 100%;\r\n}\r\n.epp-C25-db::after{\r\n  background-size:75% 100%;\r\n}\r\n.epp-C26-db::after{\r\n  background-size:80% 100%;\r\n}\r\n.epp-C27-db::after{\r\n  background-size:85% 100%;\r\n}\r\n.epp-C28-db::after{\r\n  background-size:90% 100%;\r\n}\r\n.epp-C29-db::after{\r\n  background-size:95% 100%;\r\n}\r\n.epp-C30-db::after{\r\n  background-size:100% 100%;\r\n}\r\n</style></head>\r\n<body>\r\n<table class=\"epplus-table\" role=\"table\">\r\n  <colgroup>\r\n    <col class=\"epp-dcw\" span=\"1\"/>\r\n    <col class=\"epp-dcw\" span=\"1\"/>\r\n    <col class=\"epp-dcw\" span=\"1\"/>\r\n    <col class=\"epp-dcw\" span=\"1\"/>\r\n  </colgroup>\r\n  <thead role=\"rowgroup\">\r\n    <tr role=\"row\" class=\"epp-drh\">\r\n      <th data-datatype=\"number\" class=\"epp-ar\" style=\"background-color:#ffc000;\">1</th>\r\n      <th data-datatype=\"number\" class=\"epp-ar\" style=\"background-color:#4472c4;\">1</th>\r\n      <th data-datatype=\"number\" class=\"epp-ar epp-db-shared epp-dxf1-neg epp-C1-db\">-9</th>\r\n      <th data-datatype=\"string\" class=\"epp-al\"></th>\r\n    </tr>\r\n  </thead>\r\n  <tbody role=\"rowgroup\">\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"2\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#fbbf02;\">2</td>\r\n      <td data-value=\"2\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#3f6ab6;\">2</td>\r\n      <td data-value=\"-8\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-neg epp-C2-db\">-8</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"3\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#f5bf05;\">3</td>\r\n      <td data-value=\"3\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#3a62a9;\">3</td>\r\n      <td data-value=\"-7\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-neg epp-C3-db\">-7</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"4\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#f1be07;\">4</td>\r\n      <td data-value=\"4\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#365a9b;\">4</td>\r\n      <td data-value=\"-6\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-neg epp-C4-db\">-6</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"5\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#ebbd0a;\">5</td>\r\n      <td data-value=\"5\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#31528d;\">5</td>\r\n      <td data-value=\"-5\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-neg epp-C5-db\">-5</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"6\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#e7bd0c;\">6</td>\r\n      <td data-value=\"6\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#2d4b81;\">6</td>\r\n      <td data-value=\"-4\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-neg epp-C6-db\">-4</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"7\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#e1bc0f;\">7</td>\r\n      <td data-value=\"7\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#284374;\">7</td>\r\n      <td data-value=\"-3\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-neg epp-C7-db\">-3</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"8\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#ddbb11;\">8</td>\r\n      <td data-value=\"8\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#233b66;\">8</td>\r\n      <td data-value=\"-2\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-neg epp-C8-db\">-2</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"9\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#d7bb14;\">9</td>\r\n      <td data-value=\"9\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#1f3358;\">9</td>\r\n      <td data-value=\"-1\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-neg epp-C9-db\">-1</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"10\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#d3ba16;\">10</td>\r\n      <td data-value=\"10\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#1a2b4a;\">10</td>\r\n      <td data-value=\"0\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-neg epp-C10-db\">0</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"11\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#ceba18;\">11</td>\r\n      <td data-value=\"11\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#15233d;\">11</td>\r\n      <td data-value=\"1\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C11-db\">1</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"12\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#c9b91b;\">12</td>\r\n      <td data-value=\"12\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#101b2f;\">12</td>\r\n      <td data-value=\"2\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C12-db\">2</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"13\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#c4b81d;\">13</td>\r\n      <td data-value=\"13\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#0c1321;\">13</td>\r\n      <td data-value=\"3\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C13-db\">3</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"14\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#bfb720;\">14</td>\r\n      <td data-value=\"14\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#070b14;\">14</td>\r\n      <td data-value=\"4\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C14-db\">4</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"15\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#bab722;\">15</td>\r\n      <td data-value=\"15\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#020306;\">15</td>\r\n      <td data-value=\"5\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C15-db\">5</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"16\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#b5b625;\">16</td>\r\n      <td data-value=\"16\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#070707;\">16</td>\r\n      <td data-value=\"6\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C16-db\">6</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"17\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#b0b627;\">17</td>\r\n      <td data-value=\"17\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#171717;\">17</td>\r\n      <td data-value=\"7\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C17-db\">7</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"18\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#abb52a;\">18</td>\r\n      <td data-value=\"18\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#272727;\">18</td>\r\n      <td data-value=\"8\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C18-db\">8</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"19\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#a6b42c;\">19</td>\r\n      <td data-value=\"19\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#373737;\">19</td>\r\n      <td data-value=\"9\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C19-db\">9</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"20\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#a1b32f;\">20</td>\r\n      <td data-value=\"20\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#484747;\">20</td>\r\n      <td data-value=\"10\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C20-db\">10</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"21\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#9cb331;\">21</td>\r\n      <td data-value=\"21\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#585757;\">21</td>\r\n      <td data-value=\"11\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C21-db\">11</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"22\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#98b233;\">22</td>\r\n      <td data-value=\"22\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#686868;\">22</td>\r\n      <td data-value=\"12\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C22-db\">12</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"23\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#92b236;\">23</td>\r\n      <td data-value=\"23\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#787878;\">23</td>\r\n      <td data-value=\"13\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C23-db\">13</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"24\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#8eb138;\">24</td>\r\n      <td data-value=\"24\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#888888;\">24</td>\r\n      <td data-value=\"14\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C24-db\">14</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"25\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#88b03b;\">25</td>\r\n      <td data-value=\"25\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#989898;\">25</td>\r\n      <td data-value=\"15\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C25-db\">15</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"26\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#84b03d;\">26</td>\r\n      <td data-value=\"26\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#a6a6a6;\">26</td>\r\n      <td data-value=\"16\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C26-db\">16</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"27\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#7eaf40;\">27</td>\r\n      <td data-value=\"27\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#b6b6b6;\">27</td>\r\n      <td data-value=\"17\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C27-db\">17</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"28\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#7aae42;\">28</td>\r\n      <td data-value=\"28\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#c7c6c6;\">28</td>\r\n      <td data-value=\"18\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C28-db\">18</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"29\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#74ae45;\">29</td>\r\n      <td data-value=\"29\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#d7d6d6;\">29</td>\r\n      <td data-value=\"19\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C29-db\">19</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n    <tr role=\"row\" scope=\"row\" class=\"epp-drh\">\r\n      <td data-value=\"30\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#70ad47;\">30</td>\r\n      <td data-value=\"30\" role=\"cell\" class=\"epp-ar\" style=\"background-color:#e7e6e6;\">30</td>\r\n      <td data-value=\"20\" role=\"cell\" class=\"epp-ar epp-db-shared epp-dxf1-pos epp-C30-db\">20</td>\r\n      <td role=\"cell\"></td>\r\n    </tr>\r\n  </tbody>\r\n</table>\r\n</body>\r\n</html>";
                Assert.AreEqual(expected, result);

                SaveAndCleanup(p);
            }
        }
    }
}
