using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.ConditionalFormatting
{
    [TestClass]
    public class CF_ColorScale : TestBase
    {
        [TestMethod]
        public void CF_ColourScaleColLocal()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("colourScale");

                var colorScale = sheet.ConditionalFormatting.AddThreeColorScale(new ExcelAddress("A1:A20"));

                for (int i = 1; i < 21; i++)
                {
                    sheet.Cells[i, 1].Value = i;
                }

                colorScale.LowValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                colorScale.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                colorScale.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Num;

                colorScale.MiddleValue.Formula = "$B$2";

                colorScale.LowValue.Formula = "Z34";

                colorScale.HighValue.Formula = "B6";

                colorScale.LowValue.ColorSettings.SetColor(eThemeSchemeColor.Accent3);
                colorScale.LowValue.ColorSettings.Tint = 0.5f;

                colorScale.MiddleValue.ColorSettings.Index = 4;
                colorScale.MiddleValue.ColorSettings.Tint = 1.0f;

                colorScale.HighValue.ColorSettings.Auto = true;

                var stream = new MemoryStream();
                pck.SaveAs(stream);

                var readPackage = new ExcelPackage(stream);

                var scale = readPackage.Workbook.Worksheets[0].ConditionalFormatting[0];

                var threeCol = scale.As.ThreeColorScale;

                Assert.AreEqual(scale.As.ThreeColorScale.MiddleValue.Formula, "$B$2");
                Assert.AreEqual(scale.As.ThreeColorScale.LowValue.Formula, "Z34");
                Assert.AreEqual(scale.As.ThreeColorScale.HighValue.Formula, "B6");

                Assert.AreEqual(scale.As.ThreeColorScale.LowValue.ColorSettings.Theme, eThemeSchemeColor.Accent3);
                Assert.AreEqual(scale.As.ThreeColorScale.LowValue.ColorSettings.Tint, 0.5f);

                Assert.AreEqual(threeCol.MiddleValue.ColorSettings.Index, 4);
                Assert.AreEqual(threeCol.MiddleValue.ColorSettings.Tint, 1.0f);

                Assert.AreEqual(threeCol.HighValue.ColorSettings.Auto, true);
            }
        }

        [TestMethod]
        public void WriteReadTwoColorScale()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TwoColorScale");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddTwoColorScale();
                cf.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
                cf.LowValue.Value = 2;
                cf.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;
                cf.HighValue.Value = 50;
                cf.PivotTable = true;

                Assert.AreEqual(2, cf.LowValue.Value);
                Assert.AreEqual(50, cf.HighValue.Value);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.TwoColorScale;
                    Assert.AreEqual(2, cf.LowValue.Value);
                    Assert.AreEqual(50, cf.HighValue.Value);
                }

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void WriteReadThreeColorScale()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ThreeColorScale");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddThreeColorScale();
                cf.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
                cf.LowValue.Value = 2;
                cf.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                cf.MiddleValue.Value = 25;
                cf.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;
                cf.HighValue.Value = 50;
                cf.PivotTable = true;

                Assert.AreEqual(2, cf.LowValue.Value);
                Assert.AreEqual(50, cf.HighValue.Value);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.ThreeColorScale;
                    Assert.AreEqual(2, cf.LowValue.Value);
                    Assert.AreEqual(25, cf.MiddleValue.Value);
                    Assert.AreEqual(50, cf.HighValue.Value);
                    Assert.AreEqual(eExcelConditionalFormattingValueObjectType.Num, cf.LowValue.Type);
                    Assert.AreEqual(eExcelConditionalFormattingValueObjectType.Percent, cf.MiddleValue.Type);
                    Assert.AreEqual(eExcelConditionalFormattingValueObjectType.Percentile, cf.HighValue.Type);
                }

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void CF_ColorScaleColExt()
        {
            using (var pck = OpenPackage("ColScaleTestExt.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("colourScale");
                var extSheet = pck.Workbook.Worksheets.Add("extSheet");

                var colorScale = sheet.ConditionalFormatting.AddThreeColorScale(new ExcelAddress("A1:A20"));

                for (int i = 1; i < 21; i++)
                {
                    sheet.Cells[i, 1].Value = i;
                }

                colorScale.LowValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                colorScale.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                colorScale.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Num;

                colorScale.MiddleValue.Formula = "extSheet!B2";

                extSheet.Cells["B2"].Value = 5;
                extSheet.Cells["B6"].Value = 70;

                colorScale.LowValue.Formula = "Z34";

                sheet.Cells["Z34"].Value = 4;

                colorScale.HighValue.Formula = "extSheet!B6";

                colorScale.LowValue.ColorSettings.SetColor(eThemeSchemeColor.Accent3);
                colorScale.LowValue.ColorSettings.Tint = 0.5f;

                colorScale.MiddleValue.ColorSettings.Index = 4;
                colorScale.MiddleValue.ColorSettings.Tint = 1.0f;

                colorScale.HighValue.ColorSettings.Auto = true;

                var stream = new MemoryStream();
                SaveAndCleanup(pck);

                var readPackage = OpenPackage("ColScaleTestExt.xlsx");

                var scale = readPackage.Workbook.Worksheets[0].ConditionalFormatting[0];

                var threeCol = scale.As.ThreeColorScale;

                Assert.AreEqual(scale.As.ThreeColorScale.MiddleValue.Formula, "extSheet!B2");
                Assert.AreEqual(scale.As.ThreeColorScale.LowValue.Formula, "Z34");
                Assert.AreEqual(scale.As.ThreeColorScale.HighValue.Formula, "extSheet!B6");

                Assert.AreEqual(scale.As.ThreeColorScale.LowValue.ColorSettings.Theme, eThemeSchemeColor.Accent3);
                Assert.AreEqual(scale.As.ThreeColorScale.LowValue.ColorSettings.Tint, 0.5f);

                Assert.AreEqual(threeCol.MiddleValue.ColorSettings.Index, 4);
                Assert.AreEqual(threeCol.MiddleValue.ColorSettings.Tint, 1.0f);

                Assert.AreEqual(threeCol.HighValue.ColorSettings.Auto, true);

                SaveAndCleanup(readPackage);
            }
        }

        [TestMethod]
        public void CF_ColorScaleColExtEmpty()
        {
            using (var pck = OpenPackage("ColScaleTestExtEmpty.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("colourScale");
                var extSheet = pck.Workbook.Worksheets.Add("extSheet");

                var colorScale = sheet.ConditionalFormatting.AddThreeColorScale(new ExcelAddress("A1:A20"));

                for (int i = 1; i < 21; i++)
                {
                    sheet.Cells[i, 1].Value = i;
                }

                colorScale.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                colorScale.MiddleValue.Formula = "extSheet!B2";

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void CF_ColorScaleDifficultFormula()
        {
            using (var pck = OpenPackage("ColScaleDifficultFormula.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("formulaColScale");

                ExcelAddress cfAddress1 = new ExcelAddress("A2:A10");
                var cfRule1 = ws.ConditionalFormatting.AddTwoColorScale(cfAddress1);

                cfRule1.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
                cfRule1.LowValue.Value = 4;
                cfRule1.LowValue.Color = Color.FromArgb(0xFF, 0xFF, 0xEB, 0x84);
                cfRule1.HighValue.Type = eExcelConditionalFormattingValueObjectType.Formula;
                cfRule1.HighValue.Formula = "IF($G$1=\"A</x:&'cfRule>\",1,5)";
                cfRule1.StopIfTrue = true;
                cfRule1.Style.Font.Bold = true;

                SaveAndCleanup(pck);

                var readPackage = OpenPackage("ColScaleDifficultFormula.xlsx");

                var cfRule2 = readPackage.Workbook.Worksheets[0].ConditionalFormatting[0].As.TwoColorScale;

                Assert.AreEqual("IF($G$1=\"A</x:&'cfRule>\",1,5)",cfRule2.HighValue.Formula);

                SaveAndCleanup(readPackage);
            }
        }
    }
}
