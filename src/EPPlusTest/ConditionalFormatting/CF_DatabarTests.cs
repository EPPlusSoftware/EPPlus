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

                var readBar = readPackage.Workbook.Worksheets[0].ConditionalFormatting[0];
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
    }
}
