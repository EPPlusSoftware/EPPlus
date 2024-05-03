using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Rules;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.ConditionalFormatting
{
    [TestClass]
    public class CF_IconSetTests : TestBase
    {
        [TestMethod]
        public void WriteReadThreeIcon()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("FiveIcon");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddThreeIconSet(eExcelconditionalFormatting3IconsSetType.TrafficLights2);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.ThreeIconSet;
                    Assert.AreEqual(eExcelconditionalFormatting3IconsSetType.TrafficLights2, cf.IconSet);
                }
            }
        }

        [TestMethod]
        public void WriteReadFourIcon()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("FourIcon");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddFourIconSet(eExcelconditionalFormatting4IconsSetType.ArrowsGray);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.FourIconSet;
                    Assert.AreEqual(eExcelconditionalFormatting4IconsSetType.ArrowsGray, cf.IconSet);
                }
            }
        }
        [TestMethod]
        public void WriteReadFiveIcon()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("FiveIcon");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddFiveIconSet(eExcelconditionalFormatting5IconsSetType.Arrows);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.FiveIconSet;
                    Assert.AreEqual(eExcelconditionalFormatting5IconsSetType.Arrows, cf.IconSet);
                }
            }
        }

        [TestMethod]
        public void IconSet()
        {
            var pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("IconSet");
            var cf = ws.ConditionalFormatting.AddThreeIconSet(ws.Cells["A1:A3"], eExcelconditionalFormatting3IconsSetType.Symbols);
            ws.SetValue(1, 1, 1);
            ws.SetValue(2, 1, 2);
            ws.SetValue(3, 1, 3);

            var cf4 = ws.ConditionalFormatting.AddFourIconSet(ws.Cells["B1:B4"], eExcelconditionalFormatting4IconsSetType.Rating);
            cf4.Icon1.Type = eExcelConditionalFormattingValueObjectType.Formula;
            cf4.Icon1.Formula = "0";
            cf4.Icon2.Type = eExcelConditionalFormattingValueObjectType.Formula;
            cf4.Icon2.Formula = "1/3";
            cf4.Icon3.Type = eExcelConditionalFormattingValueObjectType.Formula;
            cf4.Icon3.Formula = "2/3";
            ws.SetValue(1, 2, 1);
            ws.SetValue(2, 2, 2);
            ws.SetValue(3, 2, 3);
            ws.SetValue(4, 2, 4);

            var cf5 = ws.ConditionalFormatting.AddFiveIconSet(ws.Cells["C1:C5"], eExcelconditionalFormatting5IconsSetType.Quarters);
            cf5.Icon1.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon1.Value = 1;
            cf5.Icon2.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon2.Value = 2;
            cf5.Icon3.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon3.Value = 3;
            cf5.Icon4.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon4.Value = 4;
            cf5.Icon5.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon5.Value = 5;
            cf5.ShowValue = false;
            cf5.Reverse = true;

            ws.SetValue(1, 3, 1);
            ws.SetValue(2, 3, 2);
            ws.SetValue(3, 3, 3);
            ws.SetValue(4, 3, 4);
            ws.SetValue(5, 3, 5);
        }

        [TestMethod]
        public void WriteReadThreeIconSameAddress()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("FiveIcon");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddThreeIconSet(eExcelconditionalFormatting3IconsSetType.TrafficLights2);
                var cf2 = ws.Cells["A1"].ConditionalFormatting.AddThreeIconSet(eExcelconditionalFormatting3IconsSetType.TrafficLights1);
                var cf3 = ws.Cells["A1"].ConditionalFormatting.AddThreeIconSet(eExcelconditionalFormatting3IconsSetType.ArrowsGray);

                p.Save();

                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.ThreeIconSet;
                    cf2 = ws.ConditionalFormatting[1].As.ThreeIconSet;
                    cf3 = ws.ConditionalFormatting[2].As.ThreeIconSet;

                    Assert.AreEqual(eExcelconditionalFormatting3IconsSetType.TrafficLights2, cf.IconSet);
                    Assert.AreEqual(eExcelconditionalFormatting3IconsSetType.TrafficLights1, cf2.IconSet);
                    Assert.AreEqual(eExcelconditionalFormatting3IconsSetType.ArrowsGray, cf3.IconSet);
                }
            }
        }

        [TestMethod]
        public void WriteReadThreeIconExtSameAddress()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("FiveIcon");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddThreeIconSet(eExcelconditionalFormatting3IconsSetType.TrafficLights2);
                var cf2 = ws.Cells["A1"].ConditionalFormatting.AddThreeIconSet(eExcelconditionalFormatting3IconsSetType.Stars);
                var cf3 = ws.Cells["A1"].ConditionalFormatting.AddThreeIconSet(eExcelconditionalFormatting3IconsSetType.ArrowsGray);
                var cf4 = ws.Cells["A1"].ConditionalFormatting.AddFiveIconSet(eExcelconditionalFormatting5IconsSetType.Boxes);

                p.Save();

                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.ThreeIconSet;
                    cf2 = ws.ConditionalFormatting[1].As.ThreeIconSet;
                    cf3 = ws.ConditionalFormatting[2].As.ThreeIconSet;
                    cf4 = ws.ConditionalFormatting[3].As.FiveIconSet;

                    Assert.AreEqual(eExcelconditionalFormatting3IconsSetType.TrafficLights2, cf.IconSet);
                    Assert.AreEqual(eExcelconditionalFormatting3IconsSetType.Stars, cf3.IconSet);
                    Assert.AreEqual(eExcelconditionalFormatting3IconsSetType.ArrowsGray, cf2.IconSet);
                    Assert.AreEqual(eExcelconditionalFormatting5IconsSetType.Boxes, cf4.IconSet);
                }
            }
        }

        [TestMethod]
        public void CustomIconsWriteRead()
        {

            using (var pck = OpenPackage("FlagTest.xlsx", true))
            {
                var wks = pck.Workbook.Worksheets.Add("FormattingTest");

                for (int i = 1; i < 21; i++)
                {
                    wks.Cells[i, 1].Value = i;
                }

                wks.Cells[1, 1, 20, 1].Copy(wks.Cells[1, 2, 20, 2]);
                wks.Cells[1, 1, 20, 1].Copy(wks.Cells[1, 3, 20, 3]);
                wks.Cells[1, 1, 20, 1].Copy(wks.Cells[1, 4, 20, 4]);

                var threeIcon = wks.ConditionalFormatting.AddThreeIconSet(new ExcelAddress("A1:A20"), eExcelconditionalFormatting3IconsSetType.Triangles);

                threeIcon.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.RedFlag;
                threeIcon.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.NoIcon;
                threeIcon.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.GrayDownInclineArrow;

                var fourIcon = wks.ConditionalFormatting.AddFourIconSet(new ExcelAddress("B1:B20"), eExcelconditionalFormatting4IconsSetType.Rating);

                fourIcon.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.PinkCircle;
                fourIcon.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder;
                fourIcon.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCircleWithBorder;
                fourIcon.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircle;

                var fiveIcon = wks.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("C1:C20"), eExcelconditionalFormatting5IconsSetType.Boxes);

                fiveIcon.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.PinkCircle;
                fiveIcon.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder;
                fiveIcon.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCircleWithBorder;
                fiveIcon.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircle;
                fiveIcon.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCircle;

                var specialCase = wks.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("D1:D20"), eExcelconditionalFormatting5IconsSetType.Boxes);

                specialCase.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithNoFilledBars;
                specialCase.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithOneFilledBar;
                specialCase.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithTwoFilledBars;
                specialCase.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithThreeFilledBars;
                specialCase.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithFourFilledBars;

                SaveAndCleanup(pck);

                ExcelPackage package2 = OpenPackage("FlagTest.xlsx");
                var threeIconRead = (ExcelConditionalFormattingThreeIconSet)package2.Workbook.Worksheets[0].ConditionalFormatting[0];

                Assert.AreEqual(threeIconRead.Icon1.CustomIcon, eExcelconditionalFormattingCustomIcon.RedFlag);
                Assert.AreEqual(threeIconRead.Icon2.CustomIcon, eExcelconditionalFormattingCustomIcon.NoIcon);
                Assert.AreEqual(threeIconRead.Icon3.CustomIcon, eExcelconditionalFormattingCustomIcon.GrayDownInclineArrow);

                SaveAndCleanup(package2);
            }
        }


        [TestMethod]
        public void IconSetsAreShowValueOnDefaultAndAfterSavingAndLoading()
        {
            using (var pck = OpenPackage("valueHide.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("valueHideWs");

                sheet.Cells["A1:A20"].Formula = "Row()";

                var cf = sheet.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("A1:A20"), eExcelconditionalFormatting5IconsSetType.Arrows);

                Assert.AreEqual(true, cf.ShowValue);

                var stream = new MemoryStream();
                pck.SaveAs(stream);

                var pckRead = new ExcelPackage(stream);

                Assert.AreEqual(true, pckRead.Workbook.Worksheets[0].ConditionalFormatting[0].As.FiveIconSet.ShowValue);

                SaveAndCleanup(pckRead);
            }
        }

        [TestMethod]
        public void EnsureIconSetAttributesReadWrite()
        {
            using (var pck = OpenPackage("IconsetAttributes.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("IconsetAttributes");

                var is3 = ws.ConditionalFormatting.AddThreeIconSet(new ExcelAddress("A1:A30"), eExcelconditionalFormatting3IconsSetType.TrafficLights1);

                var is4 = ws.ConditionalFormatting.AddFourIconSet(new ExcelAddress("B1:B30"), eExcelconditionalFormatting4IconsSetType.Rating);
                var is5 = ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("C1:C30"), eExcelconditionalFormatting5IconsSetType.Quarters);

                //Chose an extLst type
                var isExt = ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("D1:D30"), eExcelconditionalFormatting5IconsSetType.Boxes);

                //Setup for custom
                var is3Custom = ws.ConditionalFormatting.AddThreeIconSet(new ExcelAddress("E1:E30"), eExcelconditionalFormatting3IconsSetType.ArrowsGray);

                is3.ShowValue = false;
                is4.ShowValue = false;
                is5.ShowValue = false;
                isExt.ShowValue = false;

                is3.IconSetPercent = false;
                is4.IconSetPercent = false;
                is5.IconSetPercent = false;
                isExt.IconSetPercent = false;

                is3.Reverse = true;
                is4.Reverse = true;
                is5.Reverse = true;
                isExt.Reverse = true;

                is3Custom.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCrossSymbol;

                ws.Cells["A1:E30"].Formula = "Row()";

                SaveAndCleanup(pck);

                var readPck = OpenPackage("IconsetAttributes.xlsx");

                var cfs = readPck.Workbook.Worksheets[0].ConditionalFormatting;

                Assert.AreEqual(false, cfs[0].As.ThreeIconSet.ShowValue);
                Assert.AreEqual(false, cfs[1].As.FourIconSet.ShowValue);
                Assert.AreEqual(false, cfs[2].As.FiveIconSet.ShowValue);
                Assert.AreEqual(false, cfs[3].As.FiveIconSet.ShowValue);


                Assert.AreEqual(false, cfs[0].As.ThreeIconSet.IconSetPercent);
                Assert.AreEqual(false, cfs[1].As.FourIconSet.IconSetPercent);
                Assert.AreEqual(false, cfs[2].As.FiveIconSet.IconSetPercent);
                Assert.AreEqual(false, cfs[3].As.FiveIconSet.IconSetPercent);

                Assert.AreEqual(true, cfs[0].As.ThreeIconSet.Reverse);
                Assert.AreEqual(true, cfs[1].As.FourIconSet.Reverse);
                Assert.AreEqual(true, cfs[2].As.FiveIconSet.Reverse);
                Assert.AreEqual(true, cfs[3].As.FiveIconSet.Reverse);

                Assert.AreEqual(false, cfs[0].As.ThreeIconSet.Custom);
                Assert.AreEqual(false, cfs[1].As.FourIconSet.Custom);
                Assert.AreEqual(false, cfs[2].As.FiveIconSet.Custom);
                Assert.AreEqual(false, cfs[3].As.FiveIconSet.Custom);

                Assert.AreEqual(eExcelconditionalFormattingCustomIcon.RedCrossSymbol, cfs[4].As.ThreeIconSet.Icon3.CustomIcon);
                Assert.AreEqual(true, cfs[4].As.ThreeIconSet.Custom);
            }
        }

        [TestMethod]
        public void EnsureCustomIconsReturnCorrectStringsAndIndex()
        {
            using (var pck = OpenPackage("CustomIcons.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("CustomIcons");

                var c1 = ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("A1:A5"), eExcelconditionalFormatting5IconsSetType.Arrows);
                var c2 = ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("B1:B5"), eExcelconditionalFormatting5IconsSetType.Arrows);
                var c3 = ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("C1:C5"), eExcelconditionalFormatting5IconsSetType.Arrows);
                var c4 = ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("D1:D5"), eExcelconditionalFormatting5IconsSetType.Arrows);
                var c5 = ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("E1:E5"), eExcelconditionalFormatting5IconsSetType.Arrows);
                var c6 = ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("F1:F5"), eExcelconditionalFormatting5IconsSetType.Arrows);
                var c7 = ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("G1:G5"), eExcelconditionalFormatting5IconsSetType.Arrows);
                var c8 = ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("H1:H5"), eExcelconditionalFormatting5IconsSetType.Arrows);
                var c9 = ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("I1:I5"), eExcelconditionalFormatting5IconsSetType.Arrows);
                var c10 = ws.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("J1:J5"), eExcelconditionalFormatting5IconsSetType.Arrows);
                var c11 = ws.ConditionalFormatting.AddThreeIconSet(new ExcelAddress("K1:K5"), eExcelconditionalFormatting3IconsSetType.Triangles);

                ws.Cells["A1:K5"].Formula = "Row()";

                c1.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.RedDownArrow;
                c1.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowSideArrow;
                c1.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenUpArrow;
                c1.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.GrayDownArrow;
                c1.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.GraySideArrow;

                Assert.AreEqual("3Arrows", c1.Icon1.GetCustomIconStringValue());
                Assert.AreEqual(0, c1.Icon1.GetCustomIconIndex());
                Assert.AreEqual("3Arrows", c1.Icon2.GetCustomIconStringValue());
                Assert.AreEqual(1, c1.Icon2.GetCustomIconIndex());
                Assert.AreEqual("3Arrows", c1.Icon3.GetCustomIconStringValue());
                Assert.AreEqual(2, c1.Icon3.GetCustomIconIndex());

                Assert.AreEqual("3ArrowsGray", c1.Icon4.GetCustomIconStringValue());
                Assert.AreEqual(0, c1.Icon4.GetCustomIconIndex());
                Assert.AreEqual("3ArrowsGray", c1.Icon5.GetCustomIconStringValue());
                Assert.AreEqual(1, c1.Icon5.GetCustomIconIndex());

                c2.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.GrayUpArrow;
                c2.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.RedFlag;
                c2.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowFlag;
                c2.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenFlag;
                c2.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCircleWithBorder;

                Assert.AreEqual("3ArrowsGray", c2.Icon1.GetCustomIconStringValue());
                Assert.AreEqual(2, c2.Icon1.GetCustomIconIndex());
                Assert.AreEqual("3Flags", c2.Icon2.GetCustomIconStringValue());
                Assert.AreEqual(0, c2.Icon2.GetCustomIconIndex());
                Assert.AreEqual("3Flags", c2.Icon3.GetCustomIconStringValue());
                Assert.AreEqual(1, c2.Icon3.GetCustomIconIndex());
                Assert.AreEqual("3Flags", c2.Icon4.GetCustomIconStringValue());
                Assert.AreEqual(2, c2.Icon4.GetCustomIconIndex());
                Assert.AreEqual("3TrafficLights1", c2.Icon5.GetCustomIconStringValue());
                Assert.AreEqual(0, c2.Icon5.GetCustomIconIndex());

                c3.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowCircle;
                c3.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenCircle;
                c3.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.RedTrafficLight;
                c3.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowTrafficLight;
                c3.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenTrafficLight;

                Assert.AreEqual("3TrafficLights1", c3.Icon1.GetCustomIconStringValue());
                Assert.AreEqual(1, c3.Icon1.GetCustomIconIndex());
                Assert.AreEqual("3TrafficLights1", c3.Icon2.GetCustomIconStringValue());
                Assert.AreEqual(2, c3.Icon2.GetCustomIconIndex());
                Assert.AreEqual("3TrafficLights2", c3.Icon3.GetCustomIconStringValue());
                Assert.AreEqual(0, c3.Icon3.GetCustomIconIndex());
                Assert.AreEqual("3TrafficLights2", c3.Icon4.GetCustomIconStringValue());
                Assert.AreEqual(1, c3.Icon4.GetCustomIconIndex());
                Assert.AreEqual("3TrafficLights2", c3.Icon5.GetCustomIconStringValue());
                Assert.AreEqual(2, c3.Icon5.GetCustomIconIndex());

                c4.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.RedDiamond;
                c4.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowTriangle;
                c4.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCrossSymbol;
                c4.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowExclamationSymbol;
                c4.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenCheckSymbol;

                Assert.AreEqual("3Signs", c4.Icon1.GetCustomIconStringValue());
                Assert.AreEqual(0, c4.Icon1.GetCustomIconIndex());
                Assert.AreEqual("3Signs", c4.Icon2.GetCustomIconStringValue());
                Assert.AreEqual(1, c4.Icon2.GetCustomIconIndex());
                Assert.AreEqual("3Symbols", c4.Icon3.GetCustomIconStringValue());
                Assert.AreEqual(0, c4.Icon3.GetCustomIconIndex());
                Assert.AreEqual("3Symbols", c4.Icon4.GetCustomIconStringValue());
                Assert.AreEqual(1, c4.Icon4.GetCustomIconIndex());
                Assert.AreEqual("3Symbols", c4.Icon5.GetCustomIconStringValue());
                Assert.AreEqual(2, c4.Icon5.GetCustomIconIndex());

                c5.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCross;
                c5.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowExclamation;
                c5.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenCheck;
                c5.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.SilverStar;
                c5.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.HalfGoldStar;

                Assert.AreEqual("3Symbols2", c5.Icon1.GetCustomIconStringValue());
                Assert.AreEqual(0, c5.Icon1.GetCustomIconIndex());
                Assert.AreEqual("3Symbols2", c5.Icon2.GetCustomIconStringValue());
                Assert.AreEqual(1, c5.Icon2.GetCustomIconIndex());
                Assert.AreEqual("3Symbols2", c5.Icon3.GetCustomIconStringValue());
                Assert.AreEqual(2, c5.Icon3.GetCustomIconIndex());
                Assert.AreEqual("3Stars", c5.Icon4.GetCustomIconStringValue());
                Assert.AreEqual(0, c5.Icon4.GetCustomIconIndex());
                Assert.AreEqual("3Stars", c5.Icon5.GetCustomIconStringValue());
                Assert.AreEqual(1, c5.Icon5.GetCustomIconIndex());

                c6.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.GoldStar;
                c6.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.RedDownTriangle;
                c6.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDash;
                c6.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenUpTriangle;
                c6.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDownInclineArrow;

                Assert.AreEqual("3Stars", c6.Icon1.GetCustomIconStringValue());
                Assert.AreEqual(2, c6.Icon1.GetCustomIconIndex());
                Assert.AreEqual("3Triangles", c6.Icon2.GetCustomIconStringValue());
                Assert.AreEqual(0, c6.Icon2.GetCustomIconIndex());
                Assert.AreEqual("3Triangles", c6.Icon3.GetCustomIconStringValue());
                Assert.AreEqual(1, c6.Icon3.GetCustomIconIndex());
                Assert.AreEqual("3Triangles", c6.Icon4.GetCustomIconStringValue());
                Assert.AreEqual(2, c6.Icon4.GetCustomIconIndex());
                Assert.AreEqual("4Arrows", c6.Icon5.GetCustomIconStringValue());
                Assert.AreEqual(1, c6.Icon5.GetCustomIconIndex());

                c7.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowUpInclineArrow;
                c7.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.GrayDownInclineArrow;
                c7.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.GrayUpInclineArrow;
                c7.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircle;
                c7.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.GrayCircle;

                Assert.AreEqual("4Arrows", c7.Icon1.GetCustomIconStringValue());
                Assert.AreEqual(2, c7.Icon1.GetCustomIconIndex());
                Assert.AreEqual("4ArrowsGray", c7.Icon2.GetCustomIconStringValue());
                Assert.AreEqual(1, c7.Icon2.GetCustomIconIndex());
                Assert.AreEqual("4ArrowsGray", c7.Icon3.GetCustomIconStringValue());
                Assert.AreEqual(2, c7.Icon3.GetCustomIconIndex());
                Assert.AreEqual("4RedToBlack", c7.Icon4.GetCustomIconStringValue());
                Assert.AreEqual(0, c7.Icon4.GetCustomIconIndex());
                Assert.AreEqual("4RedToBlack", c7.Icon5.GetCustomIconStringValue());
                Assert.AreEqual(1, c7.Icon5.GetCustomIconIndex());

                c8.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.PinkCircle;
                c8.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCircle;
                c8.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithOneFilledBar;
                c8.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithTwoFilledBars;
                c8.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithThreeFilledBars;

                Assert.AreEqual("4RedToBlack", c8.Icon1.GetCustomIconStringValue());
                Assert.AreEqual(2, c8.Icon1.GetCustomIconIndex());
                Assert.AreEqual("4RedToBlack", c8.Icon2.GetCustomIconStringValue());
                Assert.AreEqual(3, c8.Icon2.GetCustomIconIndex());
                Assert.AreEqual("4Rating", c8.Icon3.GetCustomIconStringValue());
                Assert.AreEqual(0, c8.Icon3.GetCustomIconIndex());
                Assert.AreEqual("4Rating", c8.Icon4.GetCustomIconStringValue());
                Assert.AreEqual(1, c8.Icon4.GetCustomIconIndex());
                Assert.AreEqual("4Rating", c8.Icon5.GetCustomIconStringValue());
                Assert.AreEqual(2, c8.Icon5.GetCustomIconIndex());

                c9.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithFourFilledBars;
                c9.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder;
                c9.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithNoFilledBars;
                c9.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.WhiteCircle;
                c9.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.CircleWithThreeWhiteQuarters;

                Assert.AreEqual("4Rating", c9.Icon1.GetCustomIconStringValue());
                Assert.AreEqual(3, c9.Icon1.GetCustomIconIndex());
                Assert.AreEqual("4TrafficLights", c9.Icon2.GetCustomIconStringValue());
                Assert.AreEqual(0, c9.Icon2.GetCustomIconIndex());
                Assert.AreEqual("5Rating", c9.Icon3.GetCustomIconStringValue());
                Assert.AreEqual(0, c9.Icon3.GetCustomIconIndex());
                Assert.AreEqual("5Quarters", c9.Icon4.GetCustomIconStringValue());
                Assert.AreEqual(0, c9.Icon4.GetCustomIconIndex());
                Assert.AreEqual("5Quarters", c9.Icon5.GetCustomIconStringValue());
                Assert.AreEqual(1, c9.Icon5.GetCustomIconIndex());

                c10.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.CircleWithTwoWhiteQuarters;
                c10.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.CircleWithOneWhiteQuarter;
                c10.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.ZeroFilledBoxes;
                c10.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.OneFilledBox;
                c10.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.TwoFilledBoxes;

                Assert.AreEqual("5Quarters", c10.Icon1.GetCustomIconStringValue());
                Assert.AreEqual(2, c10.Icon1.GetCustomIconIndex());
                Assert.AreEqual("5Quarters", c10.Icon2.GetCustomIconStringValue());
                Assert.AreEqual(3, c10.Icon2.GetCustomIconIndex());
                Assert.AreEqual("5Boxes", c10.Icon3.GetCustomIconStringValue());
                Assert.AreEqual(0, c10.Icon3.GetCustomIconIndex());
                Assert.AreEqual("5Boxes", c10.Icon4.GetCustomIconStringValue());
                Assert.AreEqual(1, c10.Icon4.GetCustomIconIndex());
                Assert.AreEqual("5Boxes", c10.Icon5.GetCustomIconStringValue());
                Assert.AreEqual(2, c10.Icon5.GetCustomIconIndex());

                c11.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.ThreeFilledBoxes;
                c11.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.FourFilledBoxes;
                c11.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.NoIcon;

                Assert.AreEqual("5Boxes", c11.Icon1.GetCustomIconStringValue());
                Assert.AreEqual(3, c11.Icon1.GetCustomIconIndex());
                Assert.AreEqual("5Boxes", c11.Icon2.GetCustomIconStringValue());
                Assert.AreEqual(4, c11.Icon2.GetCustomIconIndex());
                Assert.AreEqual("NoIcons", c11.Icon3.GetCustomIconStringValue());
                Assert.AreEqual(0, c11.Icon3.GetCustomIconIndex());

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void CanReadIconsetFormulas()
        {
            using(var package = OpenTemplatePackage("s665.xlsx"))
            {
                var sheet = package.Workbook.Worksheets.GetByName("Answer Sheet");
                var iconFormula = sheet.ConditionalFormatting[0].As.ThreeIconSet.Icon2.Formula;

                SaveAndCleanup(package);
            }
        }
    }
}
