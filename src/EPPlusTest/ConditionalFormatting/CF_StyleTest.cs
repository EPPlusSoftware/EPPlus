using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.ConditionalFormatting
{
    [TestClass]
    public class CF_StyleTest : TestBase
    {
        [TestMethod]
        public void CF_EnsureGradientFill()
        {
            using(var pck = OpenPackage("GradientCFs.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("GradientStyleWorksheet");

                var cf = sheet.ConditionalFormatting.AddBeginsWith(new ExcelAddress("A1"));

                sheet.Cells["A1"].Value = "Abcd";

                var cell = sheet.Cells["A1"];

                //cell.Style.Fill.PatternType = ExcelFillStyle.Solid;

                //cell.Style.Fill.Gradient.Color1.SetColor(Color.Red);
                //cell.Style.Fill.Gradient.Color2.SetColor(Color.Blue);

                cf.Text = "a";

                cf.Style.Fill.PatternType = ExcelFillStyle.Solid;

                cf.Style.Fill.Style = eDxfFillStyle.GradientFill;

                cf.Style.Fill.Gradient.Degree = 0;

                cf.Style.Fill.Gradient.Colors.Add(0);
                cf.Style.Fill.Gradient.Colors.Add(50);
                cf.Style.Fill.Gradient.Colors.Add(100);

                cf.Style.Fill.Gradient.Colors[0].Color.Color = Color.Red;
                cf.Style.Fill.Gradient.Colors[1].Color.Color = Color.Blue;
                cf.Style.Fill.Gradient.Colors[2].Color.Color = Color.Orange;


                SaveAndCleanup(pck);
            }
        }
    }
}
