using EPPlus.Fonts.OpenType.GenericFontWidths;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class AutofitWithSerializedFontMetricsTests : TestBase
    {
        private const float AutofitCorpusFontSize = 9f;

        private static readonly string[] AutofitCorpusHeaders =
       {
            "Narrow", "Wide", "Words", "Digits", "Punctuation", "Sentence", "Long mixed", "Short", "East Asian"
        };

        private static readonly string[,] AutofitCorpus =
       {
            // Narrow glyphs - the bottom width classes.
            { "illij", "lililili", "iilltjfi lliftj", "jlitfi.ijltf ilfjti" },
            // Wide glyphs - the top width classes.
            { "WMQ", "WWMMQQ", "MWQOGWM WQMOW", "WMOQGWM QWMOGW WMQOG" },
            // Ordinary words, the case that actually matters most.
            { "Name", "Stockholm", "Invoice number", "Quarterly revenue report" },
            // Digits get their own scaling factor in the measurer.
            { "23", "1234567", "1 234 567,89", "0123456789 0123456789" },
            { ".,!-", "-.,!?:;", "!!! ??? ...", "Hello, world! - (again); yes?" },
            { "One two", "The quick brown fox", "Jumps over the lazy dog", "Pack my box with five dozen jugs" },
            { "Ab1.", "Xy9! Zq2?", "Order 4711 - shipped, 12 pcs", "Ref: AB-1234/2026 (rev 3) - approved 12,5%" },
            { "A", "Hi", "OK", "End" },
            // Measured as full width regardless of font; the half width Katakana block is half.
            { "日本語", "日本語のテキスト", "ﾊﾝｶｸ ｶﾀｶﾅ", "日本語とﾊﾝｶｸの混在テキスト" }
        };

        [TestMethod]
        [DataRow("Calibri")]
        [DataRow("Aptos Narrow")]
        [DataRow("Aptos Display")]
        [DataRow("Arial")]
        [DataRow("Arial Black")]
        [DataRow("Times New Roman")]
        [DataRow("Courier New")]
        [DataRow("Liberation Serif")]
        [DataRow("Verdana")]
        [DataRow("Cambria")]
        [DataRow("Cambria Math")]
        [DataRow("Georgia")]
        [DataRow("Corbel")]
        [DataRow("Century Gothic")]
        [DataRow("Rockwell")]
        [DataRow("Trebuchet MS")]
        [DataRow("Tw Cen MT")]
        [DataRow("Tw Cen MT Condensed")]
        [DataRow("Segoe UI")]
        public void AutofitWithSerializedFonts(string fontFamily)
        {
            var columns = AutofitCorpus.GetLength(0);
            var rows = AutofitCorpus.GetLength(1);

            using (var package = new ExcelPackage())
            {
                var measurements = package.Workbook.Worksheets.Add("Measurements");
                measurements.Cells[1, 1].Value = "Font";
                measurements.Cells[1, 2].Value = "Style";
                measurements.Cells[1, 3].Value = "Column";
                measurements.Cells[1, 4].Value = "Category";
                measurements.Cells[1, 5].Value = "Widest cell";
                measurements.Cells[1, 6].Value = "EPPlus width (chars)";
                measurements.Cells[1, 7].Value = "Excel width (chars)";
                measurements.Cells[1, 1, 1, 7].Style.Font.Bold = true;
                var measurementRow = 2;

                for (var style = FontSubFamilies.Regular; style <= FontSubFamilies.BoldItalic; style++)
                {
                    var sheet = package.Workbook.Worksheets.Add(style.ToString());
                    var range = sheet.Cells[1, 1, rows + 1, columns];
                    range.Style.Font.Name = fontFamily;
                    range.Style.Font.Size = AutofitCorpusFontSize;
                    range.Style.Font.Italic = style == FontSubFamilies.Italic || style == FontSubFamilies.BoldItalic;
                    range.Style.Font.Bold = style == FontSubFamilies.Bold || style == FontSubFamilies.BoldItalic;

                    for (var col = 0; col < columns; col++)
                    {
                        sheet.Cells[1, col + 1].Value = AutofitCorpusHeaders[col];
                        for (var row = 0; row < rows; row++)
                        {
                            sheet.Cells[row + 2, col + 1].Value = AutofitCorpus[col, row];
                        }
                    }

                    sheet.Columns[1, columns].AutoFit();

                    for (var col = 0; col < columns; col++)
                    {
                        var widest = string.Empty;
                        for (var row = 0; row < rows; row++)
                        {
                            if (AutofitCorpus[col, row].Length > widest.Length)
                            {
                                widest = AutofitCorpus[col, row];
                            }
                        }

                        measurements.Cells[measurementRow, 1].Value = fontFamily;
                        measurements.Cells[measurementRow, 2].Value = style.ToString();
                        measurements.Cells[measurementRow, 3].Value = col + 1;
                        measurements.Cells[measurementRow, 4].Value = AutofitCorpusHeaders[col];
                        measurements.Cells[measurementRow, 5].Value = widest;
                        measurements.Cells[measurementRow, 6].Value = Math.Round(sheet.Column(col + 1).Width, 2);
                        measurementRow++;
                    }
                }

                // Column 7 is left empty on purpose - fill it in from Excel after running
                // Excel's own autofit on the same columns, so the two sit side by side.
                measurements.Cells[1, 1, measurementRow - 1, 7].AutoFitColumns();

                SaveWorkbook($"Autofit_SerializedFont_{fontFamily.Replace(" ", string.Empty)}.xlsx", package);
            }
        }

        [TestMethod, Ignore]
        [DataRow("Calibri", 1)]
        //[DataRow("Calibri Light", 2)]
        //[DataRow("Arial", 3)]
        //[DataRow("Arial Black", 4)]
        //[DataRow("Arial Narrow", 5)]
        //[DataRow("Bookman Old Style", 6)]
        //[DataRow("Calisto MT", 7)]
        //[DataRow("Times New Roman", 8)]
        //[DataRow("Courier New", 9)]
        //[DataRow("Liberation Serif", 10)]
        //[DataRow("Verdana", 11)]
        //[DataRow("Cambria", 12)]
        //[DataRow("Georgia", 13)]
        //[DataRow("Corbel", 14)]
        //[DataRow("Garamond", 15)]
        //[DataRow("Gill Sans MT", 16)]
        //[DataRow("Impact", 17)]
        //[DataRow("Century Gothic", 18)]
        //[DataRow("Century Schoolbook", 19)]
        //[DataRow("Rockwell", 20)]
        //[DataRow("Rockwell Condensed", 21)]
        //[DataRow("Trebuchet MS", 22)]
        //[DataRow("Tw Cen MT", 23)]
        //[DataRow("Tw Cen MT Condensed", 24)]
        [DataRow("Aptos Narrow", 25)]
        [DataRow("Aptos Display", 26)]
        public void AutofitWithSerializedFonts2(string fontFamily, int run)
        {
            var report = new ExcelPackage(@"c:\Temp\fontreport2.xlsx");
            var reportSheet = !report.Workbook.Worksheets.Any() ? report.Workbook.Worksheets.Add("Report") : report.Workbook.Worksheets["Report"];
            var reportColOffset = 3;
            var reportRow = (run - 1) * 5 + 2;
            var shortList = new List<string>
            {
                "One",
                "12,3456",
                "Hello"
            };
            var mediumList = new List<string>
            {
                "A little longer",
                "5435.1234556",
                "Something else"
            };
            var longList = new List<string>
            {
                "A little longer than the previous example",
                "5435.1234556",
                "Something else that is even longer 12345567 than above"
            };
            var reallyLongList = new List<string>
            {
                "A little longer than the previous example, 333333333333954838!!!!!!!!!!!!!!!!,,,,,",
                "5435.1234556321 - 4.32413254353",
                "Something else that is even longer 12345567 than above, 136542.5439587432 (really, really long)"
            };
            var reallyReallyLongList = new List<string>
            {
                "A little longer than the previous example, 333333333333954838!!!!!!!!!!!!!!!!,,,,,",
                "5435.1234556321 - 4.32413254353",
                "Something else that is even longer 12345567 than above, 136542.5439587432 (really, really long),,,,,,,,,,,.............&%¤#/%¤)%(/#/%#(%/&¤#`??.3123212321"
            };
            var lists = new List<List<string>>
            {
                shortList,
                mediumList,
                longList,
                reallyLongList,
                reallyReallyLongList
            };
            using (var package = new ExcelPackage())
            {
                package.Settings.TextSettings.PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
                var newFont = true;
                for (var style = FontSubFamilies.Regular; style <= FontSubFamilies.BoldItalic; style++)
                {
                    var sheet = package.Workbook.Worksheets.Add(style.ToString());
                    var range = sheet.Cells[1, 1, 5, 10];
                    range.Style.Font.Name = fontFamily;
                    range.Style.Font.Size = 9f;
                    range.Style.Font.Italic = (style == FontSubFamilies.Italic || style == FontSubFamilies.BoldItalic);
                    range.Style.Font.Bold = (style == FontSubFamilies.Bold || style == FontSubFamilies.BoldItalic);
                    var rnd = new Random();
                    for (var col = 1; col < lists.Count + 1; col++)
                    {
                        for (var row = 1; row < 4; row++)
                        {
                            var s = lists[col - 1][row - 1];
                            sheet.Cells[row, col].Value = s;
                        }
                    }
                    var sw = new Stopwatch();
                    sw.Start();
                    sheet.Columns[1, 9].AutoFit();
                    if(newFont)
                    {
                        reportSheet.Cells[reportRow, 1].Value = range.Style.Font.Name;
                        newFont = false;
                    }
                    reportSheet.Cells[reportRow, 2].Value = style.ToString();
                    for (var col = 1; col < lists.Count + 1; col++)
                    {
                        reportSheet.Cells[reportRow, col + reportColOffset].Value = sheet.Columns[col].Width;
                    }
                    reportRow++;
                    sw.Stop();
                    var ms = sw.ElapsedMilliseconds;
                }

                SaveWorkbook($"Autofit_SerializedFont_{fontFamily.Replace(" ", string.Empty)}.xlsx", package);
                report.Save();
                report.Dispose();
            }
        }

        [TestMethod, Ignore]
        [DataRow("Calibri", 1)]
        [DataRow("Arial", 2)]
        [DataRow("Arial Black", 3)]
        [DataRow("Times New Roman", 4)]
        [DataRow("Courier New", 5)]
        [DataRow("Liberation Serif", 6)]
        [DataRow("Verdana", 7)]
        [DataRow("Cambria", 8)]
        [DataRow("Cambria Math", 9)]
        [DataRow("Georgia", 10)]
        [DataRow("Corbel", 11)]
        [DataRow("Century Gothic", 12)]
        [DataRow("Rockwell", 13)]
        [DataRow("Trebuchet MS", 14)]
        [DataRow("Tw Cen MT", 15)]
        [DataRow("Tw Cen MT Condensed", 16)]
        [DataRow("MS Gothic", 17)]
        public void AutofitWithSerializedFonts_JP(string fontFamily, int run)
        {
            var report = new ExcelPackage(@"c:\Temp\fontreport_jp.xlsx");
            var reportSheet = !report.Workbook.Worksheets.Any() ? report.Workbook.Worksheets.Add("Report") : report.Workbook.Worksheets["Report"];
            var reportColOffset = 3;
            var reportRow = (run - 1) * 5 + 2;
            var shortList = new List<string>
            {
                "新しい最新スタイルです",
                "ルの拡張サポート",
                "ピボット テー"
            };
            var mediumList = new List<string>
            {
                "数式計算エンジンの改良点とサポートされる新しい関数",
                "5435.1234556",
                "Something else"
            };
            var longList = new List<string>
            {
                "A little longer than the previous example",
                "5435.1234556",
                "ェクトが完了すると、コードを管理する開発者のライセンスのみが必要"
            };
            var reallyLongList = new List<string>
            {
                "A little longer than the previous example, 333333333333954838!!!!!!!!!!!!!!!!,,,,,",
                "5435.1234556321 - 4.32413254353",
                "EPPlusは3000万回以上ダウンロードされています。世界中の何千もの企業がスプレッドシートデータを管理するために使用しています。"
            };
            var reallyReallyLongList = new List<string>
            {
                "A little longer than the previous example, 333333333333954838!!!!!!!!!!!!!!!!,,,,,",
                "5435.1234556321 - 4.32413254353",
                "場合など)、会社は、ユーザーでもあるため、そのサービスの内部ユーザー (開発者) の数をカバーするサブスクリプションをサブスクライブする必要があります。"
            };
            var lists = new List<List<string>>
            {
                shortList,
                mediumList,
                longList,
                reallyLongList,
                reallyReallyLongList
            };
            using (var package = new ExcelPackage())
            {
                package.Settings.TextSettings.PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
                var newFont = true;
                for (var style = FontSubFamilies.Regular; style <= FontSubFamilies.BoldItalic; style++)
                {
                    var sheet = package.Workbook.Worksheets.Add(style.ToString());
                    var range = sheet.Cells[1, 1, 5, 10];
                    range.Style.Font.Name = fontFamily;
                    range.Style.Font.Size = 24f;
                    range.Style.Font.Italic = (style == FontSubFamilies.Italic || style == FontSubFamilies.BoldItalic);
                    range.Style.Font.Bold = (style == FontSubFamilies.Bold || style == FontSubFamilies.BoldItalic);
                    var rnd = new Random();
                    for (var col = 1; col < lists.Count + 1; col++)
                    {
                        for (var row = 1; row < 4; row++)
                        {
                            var s = lists[col - 1][row - 1];
                            sheet.Cells[row, col].Value = s;
                        }
                    }
                    var sw = new Stopwatch();
                    sw.Start();
                    sheet.Columns[1, 9].AutoFit();
                    if (newFont)
                    {
                        reportSheet.Cells[reportRow, 1].Value = range.Style.Font.Name;
                        newFont = false;
                    }
                    reportSheet.Cells[reportRow, 2].Value = style.ToString();
                    for (var col = 1; col < lists.Count + 1; col++)
                    {
                        reportSheet.Cells[reportRow, col + reportColOffset].Value = sheet.Columns[col].Width;
                    }
                    reportRow++;
                    sw.Stop();
                    var ms = sw.ElapsedMilliseconds;
                }

                SaveWorkbook($"JP_Autofit_SerializedFont_{fontFamily.Replace(" ", string.Empty)}.xlsx", package);
                report.Save();
                report.Dispose();
            }
        }
        [TestMethod]
        public void LoadFontSizeFromResource()
        {
            using (var p = new ExcelPackage())
            {
                var expectedLoaded = 897;
                if (FontSize._isLoaded == false)
                {
                    var expectedDefault = 25;
                    Assert.AreEqual(expectedDefault, FontSize.FontHeights.Count);
                    Assert.AreEqual(expectedDefault, FontSize.FontWidths.Count);
                }
                FontSize.LoadAllFontsFromResource();
                Assert.AreEqual(expectedLoaded, FontSize.FontHeights.Count);
                Assert.AreEqual(expectedLoaded, FontSize.FontWidths.Count);
            }
        }

        [TestMethod, Ignore]
        [DataRow("Calibri")]
        [DataRow("Arial")]
        [DataRow("Times New Roman")]
        public void MeasureSpecificFont(string font)
        {
            using (var package = new ExcelPackage())
            {
                package.Settings.TextSettings.PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
                var sheet = package.Workbook.Worksheets.Add("text");
                var sheet2 = package.Workbook.Worksheets.Add("measures");
                var sheet3 = package.Workbook.Worksheets.Add("numbers");
                sheet.Cells["A1:A50"].Style.Font.Name = font;
                sheet.Cells["A1:A50"].Style.Font.Italic = true;
                var chars = "aabcdeefghijklmnopqrrssttuvxyzåäö   AABCDEEFGHIJKLMNOPQRSSTTUVXYZÅÄÖ      !!,,,,,,,,, 112233445566778899.....";
                var numbers = "11122233344455566677788899900000000........,,,,,,,       ";
                var rnd = new Random();
                for (var x = 0; x < 60; x++)
                {
                    var text = new StringBuilder();
                    for (var i = 0; i < x; i++)
                    {
                        var ix = rnd.Next(0, chars.Length);
                        text.Append(chars[ix]);
                    }
                    sheet.Cells[1, x + 1].Value = text.ToString();
                    sheet.Columns[x + 1].AutoFit();
                    sheet2.Cells[1, x + 1].Value = sheet.Columns[x + 1].Width;

                    var number = new StringBuilder();
                    for (var i = 0; i < x; i++)
                    {
                        var ix = rnd.Next(0, numbers.Length);
                        number.Append(numbers[ix]);
                    }
                    sheet3.Cells[1, x + 1].Value = number.ToString();
                    sheet3.Columns[x + 1].AutoFit();
                    sheet2.Cells[2, x + 2].Value = sheet3.Columns[x + 1].Width;
                }
                if (!Directory.Exists(@"c:\Temp\FontTests")) Directory.CreateDirectory(@"c:\Temp\FontTests");
                var path = $"c:\\Temp\\FontTests\\{font}Measurements.xlsx";
                if (File.Exists(path)) File.Delete(path);
                package.SaveAs(path);
            }
        }


        [TestMethod, Ignore]
        [DataRow("Yu Gothic", 1)]
        [DataRow("Yu Mincho", 2)]
        [DataRow("Arial Rounded MT Bold", 3)]
        [DataRow("Goudy Stout",4)]
        [DataRow("Vladimir Script",5)]     
        [DataRow("Bahnschrift SemiBold SemiConden", 6)]
        [DataRow("Copperplate Gothic Bold", 7)]
        [DataRow("Gigi", 8)]
        [DataRow("Non existing font", 9)]
        public void MeasureOtherFonts(string fontFamily, int run)
        {
            var report = new ExcelPackage(@"c:\Temp\fontreport_jp.xlsx");
            var reportSheet = !report.Workbook.Worksheets.Any() ? report.Workbook.Worksheets.Add("Report") : report.Workbook.Worksheets["Report"];
            var reportColOffset = 3;
            var reportRow = (run - 1) * 5 + 2;
            var shortList = new List<string>
            {
                "新しい最新スタイルです",
                "ルの拡張サポート",
                "ピボット テー"
            };
            var mediumList = new List<string>
            {
                "数式計算エンジンの改良点とサポートされる新しい関数",
                "5435.1234556",
                "Something else"
            };
            var longList = new List<string>
            {
                "A little longer than the previous example",
                "5435.1234556",
                "ェクトが完了すると、コードを管理する開発者のライセンスのみが必要"
            };
            var reallyLongList = new List<string>
            {
                "A little longer than the previous example, 333333333333954838!!!!!!!!!!!!!!!!,,,,,",
                "5435.1234556321 - 4.32413254353",
                "EPPlusは3000万回以上ダウンロードされています。世界中の何千もの企業がスプレッドシートデータを管理するために使用しています。"
            };
            var reallyReallyLongList = new List<string>
            {
                "A little longer than the previous example, 333333333333954838!!!!!!!!!!!!!!!!,,,,,",
                "5435.1234556321 - 4.32413254353",
                "場合など)、会社は、ユーザーでもあるため、そのサービスの内部ユーザー (開発者) の数をカバーするサブスクリプションをサブスクライブする必要があります。"
            };
            var lists = new List<List<string>>
            {
                shortList,
                mediumList,
                longList,
                reallyLongList,
                reallyReallyLongList
            };
            using (var package = new ExcelPackage())
            {
                //package.Settings.TextSettings.PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
                var newFont = true;
                for (var style = FontSubFamilies.Regular; style <= FontSubFamilies.BoldItalic; style++)
                {
                    var sheet = package.Workbook.Worksheets.Add(style.ToString());
                    var range = sheet.Cells[1, 1, 5, 10];
                    range.Style.Font.Name = fontFamily;
                    range.Style.Font.Size = 24f;
                    range.Style.Font.Italic = (style == FontSubFamilies.Italic || style == FontSubFamilies.BoldItalic);
                    range.Style.Font.Bold = (style == FontSubFamilies.Bold || style == FontSubFamilies.BoldItalic);
                    var rnd = new Random();
                    for (var col = 1; col < lists.Count + 1; col++)
                    {
                        for (var row = 1; row < 2; row++)
                        {
                            var s = lists[col - 1][row - 1];
                            sheet.Cells[row, col].Value = s;
                        }
                    }
                    var sw = new Stopwatch();
                    sw.Start();
                    sheet.Columns[1, 9].AutoFit();
                    if (newFont)
                    {
                        reportSheet.Cells[reportRow, 1].Value = range.Style.Font.Name;
                        newFont = false;
                    }
                    reportSheet.Cells[reportRow, 2].Value = style.ToString();
                    for (var col = 1; col < lists.Count + 1; col++)
                    {
                        reportSheet.Cells[reportRow, col + reportColOffset].Value = sheet.Columns[col].Width;
                    }
                    reportRow++;
                    sw.Stop();
                    var ms = sw.ElapsedMilliseconds;
                }

                SaveWorkbook($"NonExistingFonts_Autofit_{fontFamily.Replace(" ", string.Empty)}.xlsx", package);
                report.Save();
                report.Dispose();
            }
        }

#if NETFULL
        [TestMethod, Ignore]
        public void AutoFitSystemDrawing()
        {
            using(var package = new ExcelPackage())
            {
                //package.Workbook.TextSettings.FallbackTextMeasurer = new OfficeOpenXml.SkiaSharp.Text.SkiaSharpTextMeasurer();
                //var sheet = package.Workbook.Worksheets.Add("Test");
                //sheet.Cells["A1"].Value = "abc 123 SDFÖLKJE !wueriopiquwejklöpasdfj";
                //sheet.Cells["A1"].Style.Font.Name = "Times New Roman";
                //sheet.Columns.AutoFit();
                //SaveWorkbook("Autofit_Candara.xlsx", package);
            }
        }
#endif
    }
}
