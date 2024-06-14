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
    /// <summary>
    /// Test the Conditional Formatting feature
    /// </summary>
    [TestClass]
    public class CF_DatabarTests : TestBase
    {
        [TestMethod]
        public void DatabarBasicTest()
        {
            var pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("Databar");
            var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
            ws.SetValue(1, 1, 1);
            ws.SetValue(2, 1, 2);
            ws.SetValue(3, 1, 3);
            ws.SetValue(4, 1, 4);
            ws.SetValue(5, 1, 5);

            pck.SaveAs(new MemoryStream());
        }


        [TestMethod]
        public void DatabarChangingAddressCorrectly()
        {
            var pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("DatabarAddressing");
            // Ensure there is at least one element that always exists below ConditionalFormatting nodes.   
            ws.HeaderFooter.AlignWithMargins = true;
            var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
            cf.Address = new ExcelAddress("C3");

            Assert.AreEqual(cf.Address, "C3");
        }

        [TestMethod]
        public void WriteReadDataBar()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("DataBar");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddDatabar(Color.Red);

                p.Save();

                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.DataBar;
                    Assert.AreEqual(Color.Red.ToArgb(), cf.Color.ToArgb());
                }
            }
        }

        //TODO: We should most likely throw a clearer exception.
        [TestMethod]
        public void CFThrowsIfDatabarValueNotSetOnSave()
        {
            ExcelPackage pck = new ExcelPackage(new MemoryStream());

            var sheet = pck.Workbook.Worksheets.Add("DatabarValueTest");

            var databar = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A5"), Color.Green);

            databar.LowValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
            databar.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percent;

            var stream = new MemoryStream();
            pck.SaveAs(stream);
        }

        [TestMethod]
        public void ConditionalFormattingOrderDatabar()
        {
            using (var pck = OpenPackage("CF_DataBarOrder.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("formulas");

                sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("B1:B5"), Color.BlueViolet);
                sheet.ConditionalFormatting.AddExpression(new ExcelAddress("B1:B5"));
                sheet.ConditionalFormatting.AddGreaterThan(new ExcelAddress("B1:B5"));

                SaveAndCleanup(pck);

                var readPackage = OpenPackage("CF_DataBarOrder.xlsx");

                var formats = readPackage.Workbook.Worksheets[0].ConditionalFormatting;

                Assert.AreEqual(eExcelConditionalFormattingRuleType.Expression, formats[0].Type);
                Assert.AreEqual(eExcelConditionalFormattingRuleType.GreaterThan, formats[1].Type);
                Assert.AreEqual(eExcelConditionalFormattingRuleType.DataBar, formats[2].Type);
            }
        }

        [TestMethod]
        public void CF_Databar_Formula()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("databars");

                var databar = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A10"), Color.BlueViolet);

                databar.LowValue.Type = eExcelConditionalFormattingValueObjectType.Formula;
                databar.LowValue.Formula = "10";

                databar.HighValue.Type = eExcelConditionalFormattingValueObjectType.Formula;
                databar.HighValue.Formula = "20";

                var stream = new MemoryStream();
                pck.SaveAs(stream);

                var readPackage = new ExcelPackage(stream);

                var ws = readPackage.Workbook.Worksheets[0];

                var readBar = ws.ConditionalFormatting[0];
                Assert.AreEqual(readBar.As.DataBar.LowValue.Formula, "10");
                Assert.AreEqual(readBar.As.DataBar.HighValue.Formula, "20");
            }
        }

        [TestMethod]
        public void CF_DataBar_ColorSettings_WriteRead()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("databar");

                var bar = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A12"), Color.Red);

                for (int i = 1; i < 11; i++)
                {
                    sheet.Cells[i, 1].Value = i - 6;
                }

                bar.LowValue.Formula = "B5";

                bar.HighValue.Formula = "Z34";

                bar.FillColor.Color = Color.Aqua;

                bar.BorderColor.Clear();
                bar.BorderColor.Theme = eThemeSchemeColor.Accent4;
                bar.BorderColor.Tint = 0.5f;

                bar.NegativeFillColor.Color = Color.Red;

                bar.NegativeBorderColor.Auto = true;
                bar.NegativeBorderColor.Tint = 0.5f;

                bar.AxisColor.Index = 2;

                MemoryStream stream = new MemoryStream();

                pck.SaveAs(stream);

                var readPck = new ExcelPackage(stream);

                var sheet2 = readPck.Workbook.Worksheets[0];

                var cf = sheet2.ConditionalFormatting[0];

                var bar2 = cf.As.DataBar;

                Assert.AreEqual(Color.FromArgb(255, Color.Aqua), bar2.FillColor.Color);
                Assert.AreEqual(eThemeSchemeColor.Accent4, bar2.BorderColor.Theme);
                Assert.AreEqual(0.5, bar2.BorderColor.Tint);
                Assert.AreEqual(Color.FromArgb(255, Color.Red), bar2.NegativeFillColor.Color);
                Assert.AreEqual(true, bar2.NegativeBorderColor.Auto);
                Assert.AreEqual(2, bar2.AxisColor.Index);
            }
        }

        [TestMethod]
        public void CF_DataBar_Sparkline_WriteRead()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("databarSparklineDataValidation");
                var sheetExt = pck.Workbook.Worksheets.Add("databarExt");

                sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A12"), Color.Red);

                sheet.SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Line, new ExcelAddress("A1:A12"), new ExcelAddress("B1:B12"));

                var listVal = sheet.DataValidations.AddListValidation("A1:A5");

                listVal.Formula.ExcelFormula = "databarExt!A1:A50";

                var stream = new MemoryStream();
                pck.SaveAs(stream);

                var readPck = new ExcelPackage(stream);

                var readSheet = readPck.Workbook.Worksheets[0];

                var cf = readSheet.ConditionalFormatting[0];
            }
        }

        [TestMethod]
        public void CF_DatabarColorReadWrite()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("basicSheet");

                var bar = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A20"), Color.Blue);

                bar.AxisColor.Theme = eThemeSchemeColor.Accent6;
                bar.BorderColor.Theme = eThemeSchemeColor.Background1;
                bar.NegativeFillColor.SetColor(Color.Red);
                bar.NegativeBorderColor.SetColor(Color.MediumPurple);
                bar.AxisPosition = eExcelDatabarAxisPosition.Middle;

                for (int i = 1; i < 21; i++)
                {
                    sheet.Cells[i, 1].Value = i - 10;
                }

                var stream = new MemoryStream();
                pck.SaveAs(stream);

                var readPck = new ExcelPackage(stream);

                var readCF = readPck.Workbook.Worksheets[0].ConditionalFormatting[0].As.DataBar;

                Assert.AreEqual(Color.Blue.ToArgb(), readCF.FillColor.Color.Value.ToArgb());
                Assert.AreEqual(eThemeSchemeColor.Accent6, readCF.AxisColor.Theme);
                Assert.AreEqual(eThemeSchemeColor.Background1, readCF.BorderColor.Theme);
                Assert.AreEqual(Color.Red.ToArgb(), readCF.NegativeFillColor.Color.Value.ToArgb());
                Assert.AreEqual(Color.MediumPurple.ToArgb(), readCF.NegativeBorderColor.Color.Value.ToArgb());
                Assert.AreEqual(eExcelDatabarAxisPosition.Middle, readCF.AxisPosition);
            }
        }

        [TestMethod]
        public void CF_DatabarDirectionTest()
        {
            using (var pck = OpenPackage("direction.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("basicSheet");

                var bar = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A20"), Color.Blue);
                var bar2 = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:B20"), Color.Red);
                var bar3 = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("B1:B20"), Color.Yellow);

                sheet.Cells["A1:B20"].Formula = "Row()";

                bar.Direction = eDatabarDirection.RightToLeft;
                bar2.Direction = eDatabarDirection.LeftToRight;
                bar3.Direction = eDatabarDirection.Context;

                SaveAndCleanup(pck);

                var pck2 = OpenPackage("direction.xlsx");

                var cf = pck2.Workbook.Worksheets[0].ConditionalFormatting;

                Assert.AreEqual(eDatabarDirection.RightToLeft, cf[0].As.DataBar.Direction);
                Assert.AreEqual(eDatabarDirection.LeftToRight, cf[1].As.DataBar.Direction);
                Assert.AreEqual(eDatabarDirection.Context, cf[2].As.DataBar.Direction);
            }
        }

        [TestMethod]
        public void CF_DatabarReadTest()
        {
            using (var pck = OpenTemplatePackage("ExtremeIconSet_Databar.xlsx"))
            {
                var sheet1 = pck.Workbook.Worksheets[0];
                var sheet2 = pck.Workbook.Worksheets[1];

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void CF_SameNrOfFormattings()
        {
            using (var pck = OpenTemplatePackage("ExtremeIconSet_Databar.xlsx"))
            {
                var sheet1 = pck.Workbook.Worksheets[0];
                var sheet2 = pck.Workbook.Worksheets[1];

                var numCFs = sheet1.ConditionalFormatting.Count;
                var numCFs2 = sheet2.ConditionalFormatting.Count;


                var stream = new MemoryStream();
                pck.SaveAs(stream);

                var pck2 = new ExcelPackage(stream);

                var readSheet1 = pck.Workbook.Worksheets[0];
                var readSheet2 = pck.Workbook.Worksheets[1];

                Assert.AreEqual(numCFs, readSheet1.ConditionalFormatting.Count);
                Assert.AreEqual(numCFs2, readSheet2.ConditionalFormatting.Count);
            }
        }

        [TestMethod]
        public void CF_SolidFillForDatabars()
        {
            using (var pck = OpenPackage("SolidFill_Databar.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("SolidFill");

                var databar = ws.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A5"), Color.Red);

                ws.Cells["A1:A5"].Formula = "Row()";

                databar.Gradient = false;

                databar.HighValue.Type = eExcelConditionalFormattingValueObjectType.AutoMax;
                databar.LowValue.Type = eExcelConditionalFormattingValueObjectType.AutoMin;

                SaveAndCleanup(pck);

                var readPck = OpenPackage("SolidFill_Databar.xlsx");
                var sheet = readPck.Workbook.Worksheets[0];

                var readDatabar = sheet.ConditionalFormatting[0].As.DataBar;

                Assert.AreEqual(false, readDatabar.Gradient);
                Assert.AreEqual(readDatabar.HighValue.Type, eExcelConditionalFormattingValueObjectType.AutoMax);
                Assert.AreEqual(readDatabar.LowValue.Type, eExcelConditionalFormattingValueObjectType.AutoMin);
            }
        }

        [TestMethod]
        public void CF_DatabarAttributesReadWrite()
        {
            using (var pck = OpenPackage("Databar_Attributes.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("attributesDatabar");

                var databar = ws.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A20"), Color.Blue);

                ws.Cells["A1:A20"].Formula = "Row()-10";

                databar.ShowValue = false;
                databar.Gradient = false;
                databar.Border = true;
                databar.Direction = eDatabarDirection.RightToLeft;

                //Do not do this. Excel will read wrong colors on other nodes in some cases.
                /*////Ensure we read and write the color even if its not currently applied
                //databar.NegativeFillColor.Color = Color.DarkBlue;*/

                databar.NegativeBarColorSameAsPositive = true;
                databar.NegativeBarBorderColorSameAsPositive = false;
                databar.AxisPosition = eExcelDatabarAxisPosition.Middle;

                SaveAndCleanup(pck);

                var readPck = OpenPackage("Databar_Attributes.xlsx");

                var bar = readPck.Workbook.Worksheets[0].ConditionalFormatting[0].As.DataBar;

                Assert.AreEqual(false, bar.ShowValue);
                Assert.AreEqual(false, bar.Gradient);
                Assert.AreEqual(true, bar.Border);
                Assert.AreEqual(eDatabarDirection.RightToLeft, bar.Direction);
                Assert.AreEqual(true, bar.NegativeBarColorSameAsPositive);
                Assert.AreEqual(false, bar.NegativeBarBorderColorSameAsPositive);
                Assert.AreEqual(eExcelDatabarAxisPosition.Middle, bar.AxisPosition);
                ////Do not do this
                //*Assert.AreEqual(Color.FromArgb(255, Color.DarkBlue), bar.NegativeFillColor.Color);*/
            }
        }

        [TestMethod]
        public void CF_DatabarPercentage()
        {
            using (var pck = OpenPackage("DataBarPercentage.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("percentageDatabars");

                ws.Cells["A1:A30"].Formula = "ROW()-10";

                var db = ws.Cells["A1:A30"].ConditionalFormatting.AddDatabar(Color.CornflowerBlue);

                ws.Calculate();

                var dbCast = (ExcelConditionalFormattingDataBar)db;

                Assert.AreEqual(0d, dbCast.GetPercentageAtCell(ws.Cells["A10"]));
                Assert.AreEqual(50d, dbCast.GetPercentageAtCell(ws.Cells["A20"]));
                Assert.AreEqual(100d, dbCast.GetPercentageAtCell(ws.Cells["A30"]));

                //Negative range should be 100% at -9 and -1 should be 100/9
                Assert.AreEqual(100d, dbCast.GetPercentageAtCell(ws.Cells["A1"]));
                Assert.AreEqual(100d/9, dbCast.GetPercentageAtCell(ws.Cells["A9"]));
                Assert.AreEqual((100d / 9) * 2, dbCast.GetPercentageAtCell(ws.Cells["A8"]));
            }
        }

        [TestMethod]
        public void CF_DatabarThemeColor()
        {
            using (var pck = OpenPackage("DatabarThemeColor.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("dataBarSheet");

                var range = sheet.Cells["A1:A30"];

                range.Formula = "ROW()-12";
                range.Calculate();

                var cf = range.ConditionalFormatting.AddDatabar(Color.AliceBlue);

                cf.FillColor.Theme = eThemeSchemeColor.Accent6;
                cf.BorderColor.Theme = eThemeSchemeColor.Background2;
                cf.AxisColor.Theme = eThemeSchemeColor.Accent2;
                cf.NegativeBorderColor.Theme = eThemeSchemeColor.Accent4;
                cf.NegativeFillColor.Theme = eThemeSchemeColor.Hyperlink;

                SaveAndCleanup(pck);
            }

            using (var pck = OpenPackage("DatabarThemeColor.xlsx"))
            {
                var ws = pck.Workbook.Worksheets[0];

                var cfs = ws.Cells["A1"].ConditionalFormatting.GetConditionalFormattings();

                var cf = cfs[0].As.DataBar;

                Assert.AreEqual(eThemeSchemeColor.Accent6, cf.FillColor.Theme);
                Assert.AreEqual(eThemeSchemeColor.Background2, cf.BorderColor.Theme);
                Assert.AreEqual(eThemeSchemeColor.Accent2, cf.AxisColor.Theme);
                Assert.AreEqual(eThemeSchemeColor.Accent4, cf.NegativeBorderColor.Theme);
                Assert.AreEqual(eThemeSchemeColor.Hyperlink, cf.NegativeFillColor.Theme);

                SaveAndCleanup(pck);
            }
        }
    }
}
