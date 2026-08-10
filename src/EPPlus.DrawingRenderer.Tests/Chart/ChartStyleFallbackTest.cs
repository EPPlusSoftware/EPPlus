using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.DrawingRenderer.Tests.Chart
{
    [TestClass]
    public class ChartStyleFallbackTest : TestBase
    {
        [TestMethod]
        public void ReadEmptyDefaultChartStyle()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");


            using (var p = OpenTemplatePackage("StyleExamples\\emptyDefault.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                foreach (ExcelChart c in ws.Drawings)
                {
                    var borderRef = c.StyleManager.Style.Wall.BorderReference;
                    var borderSetting = c.Border;
                    var borderDirectColor = borderSetting.Fill.Color;

                    var svg = c.ToSvg();
                    SaveTextFileToWorkbook($"svg\\emptyDefaultStyle{ws.Name}_{c.Name}.svg", svg);
                }
                GetOutputFile("StyleExamples", "");
                SaveAndCleanup(p);
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
                        var borderSetting = c.Border;
                        var borderDirectColor = borderSetting.Fill.Color;
                        var theme = p.Workbook.ThemeManager.GetOrCreateTheme();

                        var defaultColorFromTheme = theme.ColorScheme.Dark1;

                        var svg = c.ToSvg();
                        SaveTextFileToWorkbook($"svg\\{fileName}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
                GetOutputFile("StyleExamples", "");
                SaveAndCleanup(p);

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
                GetOutputFile("StyleExamples", "");
                SaveAndCleanup(p);

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
                GetOutputFile("StyleExamples", "");
                SaveAndCleanup(p);
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
                GetOutputFile("StyleExamples", "");
                SaveAndCleanup(p);
            }
        }
    }
}
