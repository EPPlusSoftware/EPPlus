using FontLab1;
using FontLab1.GenericMeasurements;
using FontLab1.Tables.Os2;
using Microsoft.Testing.Extensions.TrxReport.Abstractions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.PDF.PdfSettings;
using System;

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

            var AptosNarrow_Cell = CalculateCellHeight(AptosNarrow, 11);
            var Calibri_Cell = CalculateCellHeight(Calibri, 11);
            var TrebuchetMS_Cell = CalculateCellHeight(TrebuchetMS, 11);
            var GillSansMT_Cell = CalculateCellHeight(GillSansMT, 11);
            var TwCenMT_Cell = CalculateCellHeight(TwCenMT, 11);
            var CenturyGothic_Cell = CalculateCellHeight(CenturyGothic, 11);
            var Garamond_Cell = CalculateCellHeight(Garamond, 11);
            var Corbel_Cell = CalculateCellHeight(Corbel, 11);
            var Rockwell_Cell = CalculateCellHeight(Rockwell, 11);
            var Impact_Cell = CalculateCellHeight(Impact, 11);
            var CalibriLight_Cell = CalculateCellHeight(CalibriLight, 11);
            var CalistoMT_Cell = CalculateCellHeight(CalistoMT, 11);
            var CenturySchoolbook_Cell = CalculateCellHeight(CenturySchoolbook, 11);
        }

        private double[] CalculateCellHeight(TtfFont font, double size)
        {
            double lineHeight = 0d;
            double calcType = 0d;
            if ((font.Os2Table.SelectionFlags & Os2Table.FsSelectionFlags.UseTypoMetrics) != 0)
            {
                if (font.Os2Table.sTypoLineGap == 0)
                {
                    if (font.HeadTable.Ymax > font.Os2Table.sTypoAscender)
                    {
                        calcType = 4;
                        lineHeight = font.HeadTable.Ymax - font.HeadTable.Ymin;
                    }
                    else
                    {
                        calcType = 2;
                        lineHeight = font.Os2Table.usWinAscent + font.Os2Table.usWinDescent;
                    }
                }
            }
            else
            {
                if (font.HheaTable.lineGap == 0)
                {
                    calcType = 3;
                    double max = font.HheaTable.ascender;
                    double min = font.HheaTable.descender;
                    if (font.HeadTable.Ymax > max)
                    {
                        calcType = 5;
                        max = font.HeadTable.Ymax;
                    }
                    if (font.HeadTable.Ymin < min)
                    {
                        calcType = calcType == 5 ? 10 : 7;
                        min = font.HeadTable.Ymin;
                    }
                    lineHeight = max - min;
                }
                else
                {
                    calcType = 1;
                    lineHeight = font.HheaTable.ascender - font.HheaTable.descender + font.Os2Table.sTypoLineGap;
                }
            }


            //}
            //else
            //{
            //    lineHeight = font.HheaTable.ascender - font.HheaTable.descender + font.HheaTable.lineGap;
            //}
            var lineHeightPt = lineHeight * (size / (double)font.HeadTable.UnitsPerEm);


            var cellHeight = (double)lineHeightPt + (size/10d);

            var rows = 734d / cellHeight;
            return new double[5] { calcType , rows, lineHeight, lineHeightPt, cellHeight };
        }
    }
}
