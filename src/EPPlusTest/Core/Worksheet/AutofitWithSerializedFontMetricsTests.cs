using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts;
using OfficeOpenXml.System.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class AutofitWithSerializedFontMetricsTests : TestBase
    {
        [DataTestMethod]
        //[DataRow("Calibri")]
        [DataRow("Arial")]
        //[DataRow("Arial Black")]
        [DataRow("Times New Roman")]
        [DataRow("Courier New")]
        //[DataRow("Liberation Serif")]
        [DataRow("Verdana")]
        //[DataRow("Cambria")]
        //[DataRow("Georgia")]
        //[DataRow("Corbel")]
        //[DataRow("Century Gothic")]
        public void AutofitWithSerializedFonts(string fontFamily)
        {
            using (var package = new ExcelPackage())
            {
                for(var style = FontSubFamilies.Regular; style <= FontSubFamilies.BoldItalic; style++)
                {
                    var sheet = package.Workbook.Worksheets.Add(style.ToString());
                    var range = sheet.Cells[1, 1, 5, 10];
                    range.Style.Font.Name = fontFamily;
                    range.Style.Font.Size = 9f;
                    range.Style.Font.Italic = (style == FontSubFamilies.Italic || style == FontSubFamilies.BoldItalic);
                    range.Style.Font.Bold = (style == FontSubFamilies.Bold || style == FontSubFamilies.BoldItalic);
                    var rnd = new Random();
                    for (var col = 1; col < 10; col++)
                    {
                        for (var row = 1; row < 5; row++)
                        {
                            var sb = new StringBuilder();
                            var maxLength = 35 - (col * 2);
                            var nLetters = rnd.Next(4, maxLength);
                            for (var x = 0; x < nLetters; x++)
                            {
                                var n = 65;
                                if (x % 2 == 0)
                                {
                                    n = rnd.Next(65, 90);
                                }
                                else
                                {
                                    n = rnd.Next(97, 122);
                                }

                                sb.Append((char)n);
                            }
                            sheet.Cells[row, col].Value = sb.ToString();
                        }
                    }
                    var sw = new Stopwatch();
                    sw.Start();
                    sheet.Columns[1, 9].AutoFit();
                    sw.Stop();
                    var ms = sw.ElapsedMilliseconds;
                }
                
                SaveWorkbook($"Autofit_SerializedFont_{fontFamily.Replace(" ", string.Empty)}.xlsx", package);
            }
        }

        [DataTestMethod]
        //[DataRow("Calibri")]
        //[DataRow("Arial")]
        //[DataRow("Arial Black")]
        //[DataRow("Times New Roman")]
        //[DataRow("Courier New")]
        //[DataRow("Liberation Serif")]
        //[DataRow("Verdana")]
        //[DataRow("Cambria")]
        //[DataRow("Georgia")]
        [DataRow("Corbel")]
        //[DataRow("Century Gothic")]
        public void AutofitWithSerializedFontsChinese(string fontFamily)
        {

            using (var package = new ExcelPackage())
            {
                for (var style = FontSubFamilies.Regular; style <= FontSubFamilies.BoldItalic; style++)
                {
                    var sheet = package.Workbook.Worksheets.Add(style.ToString());
                    var range = sheet.Cells[1, 1, 5, 10];
                    range.Style.Font.Name = fontFamily;
                    range.Style.Font.Size = 9f;
                    range.Style.Font.Italic = (style == FontSubFamilies.Italic || style == FontSubFamilies.BoldItalic);
                    range.Style.Font.Bold = (style == FontSubFamilies.Bold || style == FontSubFamilies.BoldItalic);
                    var rnd = new Random();
                    for (var col = 1; col < 10; col++)
                    {
                        for (var row = 1; row < 5; row++)
                        {
                            var sb = new StringBuilder();
                            var maxLength = 35 - (col * 2);
                            var nLetters = rnd.Next(4, maxLength);
                            for (var x = 0; x < nLetters; x++)
                            {
                                var n = 65;
                                if (x % 2 == 0)
                                {
                                    n = rnd.Next(65, 90);
                                }
                                else
                                {
                                    n = rnd.Next(97, 122);
                                }

                                sb.Append((char)n);
                            }
                            sheet.Cells[row, col].Value = sb.ToString();
                        }
                    }
                    var sw = new Stopwatch();
                    sw.Start();
                    sheet.Columns[1, 9].AutoFit();
                    sw.Stop();
                    var ms = sw.ElapsedMilliseconds;
                }

                SaveWorkbook($"Autofit_SerializedFont_{fontFamily.Replace(" ", string.Empty)}.xlsx", package);
            }
        }

#if Core
        [TestMethod, Ignore]
        public void AutoFitSystemDrawing()
        {
            using(var package = new ExcelPackage())
            {
                package.Workbook.TextSettings.FallbackTextMeasurer = new OfficeOpenXml.SkiaSharp.Text.SkiaSharpTextMeasurer();
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "abc 123 SDFÖLKJE !wueriopiquwejklöpasdfj";
                sheet.Cells["A1"].Style.Font.Name = "Times New Roman";
                sheet.Columns.AutoFit();
                SaveWorkbook("Autofit_Candara.xlsx", package);
            }
        }
#endif
    }
}
