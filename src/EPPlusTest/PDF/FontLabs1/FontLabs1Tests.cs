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
            TtfFont Arial = GenericFonts.GetFontData(pageSettings, "Arial");
            TtfFont Candara = GenericFonts.GetFontData(pageSettings, "Candara");
            TtfFont FranklinGothicBook = GenericFonts.GetFontData(pageSettings, "Franklin Gothic Book");
            TtfFont Georgia = GenericFonts.GetFontData(pageSettings, "Georgia");
            TtfFont TimesNewRoman = GenericFonts.GetFontData(pageSettings, "Times New Roman");
            TtfFont Constantia = GenericFonts.GetFontData(pageSettings, "Constantia");
            TtfFont PalatinoLinotype = GenericFonts.GetFontData(pageSettings, "Palatino Linotype");
            TtfFont Verdana = GenericFonts.GetFontData(pageSettings, "Verdana");


            var AptosNarrow_Cell = CalculateCellHeight2(AptosNarrow, 11);                //                               Os2 Win NoGap 51.9412598044297, Os2 Win Gap, Ymaxmin NoGap, Ymaxmin Gap,  48, 1 padding, on
            var Calibri_Cell = CalculateCellHeight2(Calibri, 11);                        //            Hhea Gap, Os2 Gap, Os2 Win NoGap 54.6629818181818,                                           50, 1 padding, on
            var TrebuchetMS_Cell = CalculateCellHeight2(TrebuchetMS, 11);                //                                             57.4673904732778,              Ymaxmin NoGap, Ymaxmin Gap,  49, 1 padding, on
            var GillSansMT_Cell = CalculateCellHeight2(GillSansMT, 11);                  //                                             57.539980861244 ,                             Ymaxmin Gap,  46, 1 padding, off +2
            var TwCenMT_Cell = CalculateCellHeight2(TwCenMT, 11);                        //Hhea NoGap, Hhea gap,          Os2 Win NoGap 61.2813697513249,                                           56, 1 padding, off +2 //fit: windesc+abs(windesc-ymin)
            var CenturyGothic_Cell = CalculateCellHeight2(CenturyGothic, 11);            //Hhea NoGap, Hhea Gap,          Os2 Win NoGap 56.0071535022355,                                           50, 1 padding, off -1,
            var Garamond_Cell = CalculateCellHeight2(Garamond, 11);                      //Hhea NoGap, Hhea Gap,          Os2 Win NoGap 59.3131313131313,                                           54, 1 padding, off -1
            var Corbel_Cell = CalculateCellHeight2(Corbel, 11);                          //                               Os2 Win NoGap 54.6629818181818,                                           50, 1 padding, on
            var Rockwell_Cell = CalculateCellHeight2(Rockwell, 11);                      //                      Os2 Gap,               56.8222264222264,                                           57, 1 padding, off +2
            var Impact_Cell = CalculateCellHeight2(Impact, 11);                          //Hhea NoGap, Hhea Gap,          Os2 Win NoGap 54.7067472159546,                                           50, 1 padding, off -1 //can be rounded up for fit
            var CalibriLight_Cell = CalculateCellHeight2(CalibriLight, 11);              //            Hhea Gap, Os2 Gap, Os2 Win NoGap 54.6629818181818,              Ymaxmin Nogap,               50, 1 padding, on
            var CalistoMT_Cell = CalculateCellHeight2(CalistoMT, 11);                    //Hhea NoGap, Hhea Gap,          Os2 Win NoGap 57.8077218889402,                                           53, 1 padding, on
            var CenturySchoolbook_Cell = CalculateCellHeight2(CenturySchoolbook, 11);    //Hhea NoGap, Hhea Gap,          Os2 Win NoGap 55.5292379298881,                                           51, 1 padding, on
            var Arial_Cell = CalculateCellHeight2(Arial, 11);                            //                                             59.7279084551812,                                    
            var Candara_Cell = CalculateCellHeight2(Candara, 11);                        //                                             54.6629818181818,                            
            var FranklinGothicBook_Cell = CalculateCellHeight2(FranklinGothicBook, 11);  //                                             58.8533395975256,                                
            var Georgia_Cell = CalculateCellHeight2(Georgia, 11);                        //                                             58.726882056491 ,                                   
            var TimesNewRoman_Cell = CalculateCellHeight2(TimesNewRoman, 11);            //                                             60.2546095879429,                                                        
            var Constantia_Cell = CalculateCellHeight2(Constantia, 11);                  //                                             54.6629818181818,                                            
            var PalatinoLinotype_Cell = CalculateCellHeight2(PalatinoLinotype, 11);      //                                             49.4598098246307,                                                    
            var Verdana_Cell = CalculateCellHeight2(Verdana, 11);                        //                                             54.9045618905   ,                                         
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
            var lineHeightPad = lineHeightPt + 0.35;
            var rows2 = 734d / lineHeightPad;

            var fontDescentPoints = Math.Abs(desc) * size / em;
            var dyDescentPixels = fontDescentPoints * 96f / 72f;

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
