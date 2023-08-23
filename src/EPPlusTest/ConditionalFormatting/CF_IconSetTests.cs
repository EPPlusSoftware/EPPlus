﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
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
        public void IconSetsAreHideValueOnDefaultAndAfterSavingAndLoading()
        {
            using (var pck = OpenPackage("valueHide.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("valueHideWs");

                sheet.Cells["A1:A20"].Formula = "Row()";

                var cf = sheet.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("A1:A20"), eExcelconditionalFormatting5IconsSetType.Arrows);

                Assert.AreEqual(false, cf.ShowValue);

                var stream = new MemoryStream();
                pck.SaveAs(stream);

                var pckRead = new ExcelPackage(stream);

                Assert.AreEqual(false, pckRead.Workbook.Worksheets[0].ConditionalFormatting[0].As.FiveIconSet.ShowValue);

                SaveAndCleanup(pckRead);
            }
        }

        [TestMethod]
        public void EnsureCustomIconsReturnCorrectStrings()
        {
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("FormattingTest");

                var validation = (ExcelConditionalFormattingThreeIconSet)wks.ConditionalFormatting.AddThreeIconSet(new ExcelAddress("A1"), eExcelconditionalFormatting3IconsSetType.Triangles);

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowSideArrow;

                Assert.AreEqual("3Arrows", validation.Icon1.GetCustomIconStringValue());

                validation.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.GrayUpArrow;

                Assert.AreEqual("3ArrowsGray", validation.Icon2.GetCustomIconStringValue());

                validation.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowFlag;
                Assert.AreEqual("3Flags", validation.Icon3.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenCircle;

                Assert.AreEqual("3TrafficLights1", validation.Icon1.GetCustomIconStringValue());

                validation.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowTrafficLight;
                Assert.AreEqual("3TrafficLights2", validation.Icon2.GetCustomIconStringValue());

                validation.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.RedDiamond;
                Assert.AreEqual("3Signs", validation.Icon3.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowExclamationSymbol;
                Assert.AreEqual("3Symbols", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenCheck;
                Assert.AreEqual("3Symbols2", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.HalfGoldStar;
                Assert.AreEqual("3Stars", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDash;
                Assert.AreEqual("3Triangles", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDash;
                Assert.AreEqual("3Triangles", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDash;
                Assert.AreEqual("3Triangles", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDownInclineArrow;
                Assert.AreEqual("4Arrows", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.PinkCircle;
                Assert.AreEqual("4RedToBlack", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithThreeFilledBars;
                Assert.AreEqual("4Rating", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder;
                Assert.AreEqual("4TrafficLights", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithNoFilledBars;
                Assert.AreEqual("5Rating", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.CircleWithThreeWhiteQuarters;
                Assert.AreEqual("5Quarters", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.OneFilledBox;
                Assert.AreEqual("5Boxes", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.NoIcon;
                Assert.AreEqual("NoIcons", validation.Icon1.GetCustomIconStringValue());
            }
        }
    }
}
