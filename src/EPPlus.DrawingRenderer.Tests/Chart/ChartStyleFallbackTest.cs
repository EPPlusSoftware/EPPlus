using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using static OfficeOpenXml.Drawing.OleObject.Structures.OleObjectDataStructures;

namespace EPPlus.DrawingRenderer.Tests.Chart
{
    [TestClass]
    public class ChartStyleFallbackTest : TestBase
    {

        [TestMethod]
        public void EpplusGeneratedChart()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");


            using (var p = OpenPackage("StyleExamples\\epplusDefaultTest.xlsx",true))
            {
                var ws = p.Workbook.Worksheets.Add("EpplusGeneratedChart");

                //ws.Workbook.ThemeManager.GetOrCreateTheme();


                ws.Cells["A1:A3"].Formula = "ROW()+COLUMN()";

                ws.Calculate();

                var emptyLines = ws.Drawings.AddLineChart("EmptyLineChart", eLineChartType.Line);
                var generatedBar = ws.Drawings.AddBarChart("EpplusBarChart", eBarChartType.ColumnClustered);

                generatedBar.SetPosition(1, 1000);

                var defaultRect = ws.Drawings.AddShape("MyDefaultShape", OfficeOpenXml.Drawing.eShapeStyle.Round1Rect);
                var gradientRect = ws.Drawings.AddShape("GradRect", OfficeOpenXml.Drawing.eShapeStyle.Round1Rect);

                defaultRect.SetPosition(300, 1);
                gradientRect.SetPosition(300, 1000);

                defaultRect.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
                gradientRect.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.GradientFill;
                generatedBar.Series.Add(ws.Cells["A1:A3"]);

                //foreach (ExcelChart c in ws.Drawings)
                //{
                //    var borderRef = c.StyleManager.Style.ChartArea.BorderReference;
                //    var borderSetting = c.Border;
                //    var borderDirectColor = borderSetting.Fill.Color;

                //    var svg = c.ToSvg();
                //    SaveTextFileToWorkbook($"svg\\epplusDefault{ws.Name}_{c.Name}.svg", svg);
                //}
                //GetOutputFile("StyleExamples", "");
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ReadEmptyDefaultChartStyle()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");


            using (var p = OpenTemplatePackage("StyleExamples\\emptyDefault.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                foreach (ExcelChart c in ws.Drawings)
                {
                    var borderRef = c.StyleManager.Style.ChartArea.BorderReference;
                    var borderSetting = c.Border;
                    var borderDirectColor = borderSetting.Fill.Color;

                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\emptyDefaultStyle{ws.Name}_{c.Name}.svg", svg);
                }
                var fi = GetOutputFile("StyleExamples", "emptyDefault_out.xlsx");
                p.SaveAs(fi);
            }
        }
        [TestMethod]
        public void GenerateSimpleChart()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            string fileName = "EpplusSimpleChart";

            using (var p = OpenPackage($"{fileName}.xlsx",true))
            {
                var ws = p.Workbook.Worksheets.Add("s1");
                ws.Drawings.AddBarChart("simpleChart", eBarChartType.ColumnClustered);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ReadChartBorderThemeTint()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            var fileName = "ChartBorderThemeTint";

            using (var p = OpenTemplatePackage($"StyleExamples\\{fileName}.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var lChart = ws.Drawings[0].As.Chart.LineChart;

                lChart.StyleManager.Style.ChartArea.Border.Fill.SolidFill.Color.SetSchemeColor(OfficeOpenXml.Drawing.eSchemeColor.Accent1);
        
                //100 - input is what excel seems to apply
                //lChart.StyleManager.Style.ChartArea.BorderReference.Color.Transforms.AddTint(13);
                lChart.StyleManager.Style.ChartArea.Border.Fill.SolidFill.Color.Transforms.AddTint(60);
                lChart.StyleManager.Style.ChartArea.Border.Width = 10d;
                lChart.StyleManager.ApplyStyles();

                var fi = GetOutputFile("StyleExamples", $"{fileName}_Out.xlsx");
                p.SaveAs(fi);
            }
        }

        [TestMethod]
        public void RemovedStyles()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            string fileName = "emptyManuallyRemovedLnStyles";

            using (var p = OpenTemplatePackage($"StyleExamples\\{fileName}.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                foreach (var d in ws.Drawings)
                {
                    if (d is ExcelChart c)
                    {
                        //var borderSetting = c.Border;
                        //var borderDirectColor = borderSetting.Fill.Color;
                        //var theme = p.Workbook.ThemeManager.GetOrCreateTheme();

                        //var defaultColorFromTheme = theme.ColorScheme.Dark1;

                        var svg = c.ToSvg();
                        SaveTextFileToWorkbook($"svg\\{fileName}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
                var fi = GetOutputFile("StyleExamples", $"{fileName}_Out.xlsx");
                p.SaveAs(fi);
            }
        }

        [TestMethod]
        public void EditedTheme()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");


            string fileName = "ExcelThemeEdited";

            using (var p = OpenTemplatePackage($"StyleExamples\\{fileName}.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                foreach (var d in ws.Drawings)
                {
                    if (d is ExcelChart c)
                    {
                        //var borderSetting = c.Border;
                        //var borderDirectColor = borderSetting.Fill.Color;
                        //var theme = p.Workbook.ThemeManager.GetOrCreateTheme();

                        //var defaultColorFromTheme = theme.ColorScheme.Dark1;

                        var svg = c.ToSvg();
                        SaveTextFileToWorkbook($"svg\\{fileName}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
                var fi = GetOutputFile("StyleExamples", $"{fileName}_Out.xlsx");
                p.SaveAs(fi);
            }
        }


        [TestMethod]
        public void ManualSystemText()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            string fileName = "ExcelThemeManualSystemText";

            using (var p = OpenTemplatePackage($"StyleExamples\\{fileName}.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                foreach (var d in ws.Drawings)
                {
                    if (d is ExcelChart c)
                    {
                        var borderSetting = c.Border;
                        var borderDirectColor = borderSetting.Fill.Color;
                        var theme = p.Workbook.ThemeManager.GetOrCreateTheme();

                        var defaultColorFromTheme = theme.ColorScheme.Dark1;

                        var svg = c.ToSvg();
                        SaveTextFileToWorkbook($"svg\\{fileName}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
                var fi = GetOutputFile("StyleExamples", $"{fileName}_Out.xlsx");
                p.SaveAs(fi);
            }
        }


        [TestMethod]
        public void ExcelThemeLnDeleted()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            string fileName = "ExcelThemeLnDeleted";

            using (var p = OpenTemplatePackage($"StyleExamples\\{fileName}.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                foreach (var d in ws.Drawings)
                {
                    if (d is ExcelChart c)
                    {
                        var borderSetting = c.Border;
                        var borderDirectColor = borderSetting.Fill.Color;
                        var theme = p.Workbook.ThemeManager.GetOrCreateTheme();

                        var defaultColorFromTheme = theme.ColorScheme.Dark1;

                        var svg = c.ToSvg();
                        SaveTextFileToWorkbook($"svg\\{fileName}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
                var fi = GetOutputFile("StyleExamples", $"{fileName}_Out.xlsx");
                p.SaveAs(fi);
            }
        }


        [TestMethod]
        public void PureExcelTheme()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            string fileName = "PureExcelTheme";

            using (var p = OpenTemplatePackage($"StyleExamples\\{fileName}.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                foreach (var d in ws.Drawings)
                {
                    if (d is ExcelChart c)
                    {
                        var borderSetting = c.Border;
                        var borderDirectColor = borderSetting.Fill.Color;
                        var theme = p.Workbook.ThemeManager.GetOrCreateTheme();

                        var defaultColorFromTheme = theme.ColorScheme.Dark1;

                        var svg = c.ToSvg();
                        SaveTextFileToWorkbook($"svg\\{fileName}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
                var fi = GetOutputFile("StyleExamples", $"{fileName}_Out.xlsx");
                p.SaveAs(fi);
            }
        }
    }
}
