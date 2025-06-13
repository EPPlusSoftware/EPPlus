using FontLab1;
using FontLab1.GenericMeasurements;
using FontLab1.Tables.Hhea;
using FontLab1.Tables.Os2;
using Microsoft.Testing.Extensions.TrxReport.Abstractions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.PDF.PdfSettings;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace EPPlusTest.PDF.FontLabs1
{
    [TestClass]
    public class FontLabs1Tests : TestBase
    {
        [TestMethod]
        public void ReadFontsFromSystem()
        {
            PdfPageSettings pageSettings = new PdfPageSettings();
            //TtfFont arialData = GenericFonts.GetFontData("Arial");
            TtfFont aptosData = GenericFonts.GetFontData(pageSettings, "Aptos Narrow");
        }

        [TestMethod]
        public void ReadNormalStyleFont()
        {
            using var p1 = OpenTemplatePackage("PdfGrids\\ThemeTest.xlsx");

            var ws1 = p1.Workbook.Worksheets[0];
            
            var ns1 = p1.Workbook.Styles.GetNormalStyle();
            var nsf1 = ns1.Style.Font.Name;

            PdfPageSettings pageSettings = new PdfPageSettings();
            TtfFont f1 = GenericFonts.GetFontData(pageSettings, nsf1);

            var cap = f1.Os2Table.sCapHeight == 0 ? (f1.Os2Table.sTypoAscender - f1.Os2Table.sTypoDescender) : f1.Os2Table.sCapHeight;
            var row = ((double)cap / 100d) + 1.54d;

            var tr1 = 734d / row;
        }




        [TestMethod]
        public void ReadFontFromSystem()
        {
            PdfPageSettings pageSettings = new PdfPageSettings();
            TtfFont AptosNarrow = GenericFonts.GetFontData(pageSettings, "Aptos Narrow");
            TtfFont Calibri = GenericFonts.GetFontData(pageSettings, "Calibri");
            TtfFont TrebuchetMS = GenericFonts.GetFontData(pageSettings, "Trebuchet MS");
            TtfFont GillSansMT = GenericFonts.GetFontData(pageSettings, "Gill Sans MT");
            TtfFont TwCenMT = GenericFonts.GetFontData(pageSettings, "Tw Cen MT");
            TtfFont CenturyGothic = GenericFonts.GetFontData(pageSettings, "Century Gothic");
            TtfFont Garamond = GenericFonts.GetFontData(pageSettings, "Garamond");
            TtfFont Corbel = GenericFonts.GetFontData(pageSettings, "Corbel");
            TtfFont Rockwell = GenericFonts.GetFontData(pageSettings, "Rockwell");
            TtfFont Impact = GenericFonts.GetFontData(pageSettings, "Impact");
            TtfFont CalibriLight = GenericFonts.GetFontData(pageSettings, "Calibri Light");
            TtfFont CalistoMT = GenericFonts.GetFontData(pageSettings, "Calisto MT");
            TtfFont CenturySchoolbook = GenericFonts.GetFontData(pageSettings, "Century Schoolbook");

            //  |    Excel Data            |           HHea           |                          OS/2                                         |   Head   |
            //  |Font             Size Rows|Ascender Descender LineGap|UseTypoMetrics winAscent WinDecent Ascender Descender LineGap CapHeight|UnitsPerEm|
            //  |Aptos Narrow       11   48|    1923      -577       0|             1      2068       563     1923      -577       0      1346|      2048|
            //  |Calibri            11   50|    1536      -512     452|             0      1950       550     1536      -512     452      1294|      2048|
            //  |Trebuchet MS       11   49|    1923      -455       0|             0      1923       455     1510      -420       0      1465|      2048|
            //  |Gill Sans MT       11   44|    1903      -472       0|             0      1903       472     1415      -471     305         0|      2048|
            //  |Tw Cen MT          11   54|    1753      -477       0|             0      1753       477     1413      -387     391         0|      2048|
            //  |Century Gothic     11   51|    2060      -451       0|             0      1989       451     1536      -426     229         0|      2048|
            //  |Garamond           11   55|    1765      -539       0|             0      1765       539     1339      -539     313      1536|      2048|
            //  |Corbel             11   50|    1523      -525     425|             0      1950       550     1523      -525     425      1338|      2048|
            //  |Rockwell           11   55|    1937      -468       0|             0      1937       468     1419      -456     316         0|      2048|
            //  |Impact             11   51|    2066      -432       0|             0      2066       432     1619      -229     343      1619|      2048|
            //  |Calibri Light      11   50|    1536      -512     452|             0      1950       550     1536      -512     452      1294|      2048|
            //  |Calisto MT         11   53|    1894      -470       0|             0      1894       470     1459      -430     302         0|      2048|
            //  |Century Schoolbook 11   51|    2019      -443       0|             0      2019       442     1516      -399     276         0|      2048|

            var AptosNarrow_Cell = CalculateCellHeight2(AptosNarrow, 11);                //                               Os2 Win NoGap, Os2 Win Gap, Ymaxmin NoGap, Ymaxmin Gap,  48, 1 padding, on
            var Calibri_Cell = CalculateCellHeight2(Calibri, 11);                        //            Hhea Gap, Os2 Gap, Os2 Win NoGap,                                           50, 1 padding, on
            var TrebuchetMS_Cell = CalculateCellHeight2(TrebuchetMS, 11);                //                                                           Ymaxmin NoGap, Ymaxmin Gap,  49, 1 padding, on
            var GillSansMT_Cell = CalculateCellHeight2(GillSansMT, 11);                  //                                                                          Ymaxmin Gap,  46, 1 padding, off +2
            var TwCenMT_Cell = CalculateCellHeight2(TwCenMT, 11);                        //Hhea NoGap, Hhea gap,                                                                   56, 1 padding, off +2
            var CenturyGothic_Cell = CalculateCellHeight2(CenturyGothic, 11);            //Hhea NoGap, Hhea Gap,                                                                   50, 1 padding, off -1,
                                                                                   //****//                               Os2 Win NoGap,                                           52, 1 padding, off +1
            var Garamond_Cell = CalculateCellHeight2(Garamond, 11);                      //Hhea NoGap, Hhea Gap,          Os2 Win NoGap,                                           54, 1 padding, off -1
            var Corbel_Cell = CalculateCellHeight2(Corbel, 11);                          //                               Os2 Win NoGap,                                           50, 1 padding, on
            var Rockwell_Cell = CalculateCellHeight2(Rockwell, 11);                      //                      Os2 Gap,                                                          57, 1 padding, off +2
            var Impact_Cell = CalculateCellHeight2(Impact, 11);                          //Hhea NoGap, Hhea Gap,                                                                   50, 1 padding, off -1
            var CalibriLight_Cell = CalculateCellHeight2(CalibriLight, 11);              //            Hhea Gap, Os2 Gap, Os2 Win NoGap,              Ymaxmin Nogap,               50, 1 padding, on
            var CalistoMT_Cell = CalculateCellHeight2(CalistoMT, 11);                    //Hhea NoGap, Hhea Gap,          Os2 Win NoGap,                                           53, 1 padding, on
            var CenturySchoolbook_Cell = CalculateCellHeight2(CenturySchoolbook, 11);    //Hhea NoGap, Hhea Gap,          Os2 Win NoGap,                                           51, 1 padding, on
        }

        private List<string[]> CalculateCellHeight2(TtfFont font, double size)
        {
            List<string[]> strings = new List<string[]>();
            strings.Add(calc("HHea No Gap", font.HheaTable.ascender, font.HheaTable.descender, 0, size, font.HeadTable.UnitsPerEm));

            strings.Add(calc("HHea Gap", font.HheaTable.ascender, font.HheaTable.descender, font.HheaTable.lineGap, size, font.HeadTable.UnitsPerEm));

            strings.Add(calc("Os2 No Gap", font.Os2Table.sTypoAscender, font.Os2Table.sTypoDescender, 0, size, font.HeadTable.UnitsPerEm));

            strings.Add(calc("Os2 Gap", font.Os2Table.sTypoAscender, font.Os2Table.sTypoDescender, font.Os2Table.sTypoLineGap, size, font.HeadTable.UnitsPerEm));

            strings.Add(calc("Os2 Win No Gap", font.Os2Table.usWinAscent, font.Os2Table.usWinDescent, 0, size, font.HeadTable.UnitsPerEm));

            strings.Add(calc("Os2 Win Gap", font.Os2Table.usWinAscent, font.Os2Table.usWinDescent, font.Os2Table.sTypoLineGap, size, font.HeadTable.UnitsPerEm));

            strings.Add(calc("Ymaxmin No Gap", font.HeadTable.Ymax, font.HeadTable.Ymin, 0, size, font.HeadTable.UnitsPerEm));

            strings.Add(calc("Ymaxmin Gap", font.HeadTable.Ymax, font.HeadTable.Ymin, font.Os2Table.sTypoLineGap, size, font.HeadTable.UnitsPerEm));
            return strings;
        }

        private string[] calc(string CalcMethod, double asc,double desc, double gap, double size, double em)
        {
            var lineHeight = asc + Math.Abs( desc) + gap;
            var lineHeightPt = lineHeight * (size / em);
            var rows = 734d / lineHeightPt;
            var lineHeightPad = lineHeightPt + 1d;
            var rows2 = 734d / lineHeightPad;
            return new string[] {CalcMethod, lineHeight.ToString(), lineHeightPt.ToString(), rows.ToString(), lineHeightPad.ToString(), rows2.ToString() };
        }

        private double[] CalculateCellHeight(TtfFont font, double size)
        {
            double lineHeight = 0d;
            double calcType = 0d;
            double padding = 1d;
            if ((font.Os2Table.SelectionFlags & Os2Table.FsSelectionFlags.UseTypoMetrics) != 0)
            {
                var max = Math.Max(Math.Max(font.HheaTable.ascender, font.HeadTable.Ymax), Math.Max(font.Os2Table.usWinAscent, font.Os2Table.sTypoAscender));
                var min = Math.Min(Math.Min(font.HheaTable.descender, font.HeadTable.Ymin), Math.Min(font.Os2Table.usWinDescent, font.Os2Table.sTypoDescender));
                var gap = Math.Max(font.HheaTable.lineGap, font.Os2Table.sTypoLineGap);
                lineHeight = max - min;
                //    if (font.Os2Table.sTypoLineGap == 0)
                //    {
                //        if (font.HeadTable.Ymax > font.Os2Table.sTypoAscender)
                //        {
                //            calcType = 4;
                //            lineHeight = font.HeadTable.Ymax - font.HeadTable.Ymin;
                //        }
                //        else
                //        {
                //            calcType = 2;
                //            lineHeight = font.Os2Table.usWinAscent + font.Os2Table.usWinDescent;
                //        }
                //    }
            }
            else
            {
                var max = font.HheaTable.ascender;
                var min = font.HheaTable.descender;
                var gap = font.HheaTable.lineGap;
                padding = 0d;
                if (font.HheaTable.lineGap != 0)
                {
                    max = font.HheaTable.ascender;
                    min = font.HheaTable.descender;
                    padding = 1d;
                }
                else if (font.Os2Table.sTypoLineGap != 0)
                {
                    max = Math.Max(font.HheaTable.ascender, font.HeadTable.Ymax);
                    min = font.HeadTable.Ymin;
                    gap = font.Os2Table.sTypoLineGap;
                    padding = 1d;
                }
                lineHeight = max - min + gap;
                //    if (font.HheaTable.lineGap == 0)
                //    {
                //        if ((font.Os2Table.sTypoLineGap == 0))
                //        {
                //            calcType = 3;
                //            double max = font.HheaTable.ascender;
                //            double min = font.HheaTable.descender;
                //            if (font.HeadTable.Ymax > max)
                //            {
                //                calcType = 5;
                //                max = font.HeadTable.Ymax;
                //            }
                //            if (font.HeadTable.Ymin < min)
                //            {
                //                calcType = calcType == 5 ? 10 : 7;
                //                min = font.HeadTable.Ymin;
                //            }
                //            lineHeight = max - min;
                //        }
                //        else
                //        {
                //            calcType = 9;

                //            lineHeight = font.Os2Table.usWinAscent + font.Os2Table.usWinDescent + font.Os2Table.sTypoLineGap;
                //        }
                //    }
                //    else
                //    {
                //        calcType = 1;
                //        lineHeight = font.HheaTable.ascender - font.HheaTable.descender + font.HheaTable.lineGap;
                //    }
            }

            var lineHeightPt = lineHeight * (size / (double)font.HeadTable.UnitsPerEm);
            var cellHeight = (double)lineHeightPt + padding;
            var rows = 734d / cellHeight;
            return new double[5] { calcType , rows, lineHeight, lineHeightPt, cellHeight };
        }
    }
}
