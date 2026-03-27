using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.ImageRenderer.Tests
{
    [TestClass]
    public class StyleTests : TestBase
    {
        [TestMethod]
        public void ExtractThemeStyleWorks()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            using (var p = OpenTemplatePackage("MyLineIonTheme.xlsx"))
            {
                var wbStyles = p.Workbook.Styles;
                var simpleChart = p.Workbook.Worksheets[0].Drawings[0].As.Chart.LineChart;

                simpleChart.Title.Text = "Hello";

                p.Workbook.ThemeManager.CurrentTheme.ColorScheme.Accent1.SetPresetColor(Color.DarkGoldenrod);

                var chartDefaultStyle = simpleChart.StyleManager.Style;
                //chartDefaultStyle.DataPointLine.FillReference.Color

                //Highest order of styling if datapoint does not exist
                //simpleChart.StyleManager.Style.DeleteAllNode("cs:dataPointLine");
                simpleChart.StyleManager.Style.DataPointLine.BorderReference.Color.SetPresetColor(Color.Green);

                //simpleChart.StyleManager.Style.DataPointLine.Border.Fill.DeleteNode(path);
                //simpleChart.StyleManager.Style.DataPointLine.Border.Fill.Color = Color.Red;
                //simpleChart.StyleManager.Style.DataPointLine.Fill.Color = Color.Green;

                //simpleChart.StyleManager.Style.DataPoint.Fill.Color = Color.Yellow;
                //simpleChart.StyleManager.Style.DataPoint.Border.Fill.Color = Color.Magenta;

                simpleChart.StyleManager.ApplyStyles();

                //var serLine = chartDefaultStyle.SeriesLine;
                //var lineElement = serLine.Border.LineElement;
                //var serLineFillRef = serLine.Border.Fill;
                //var myFillRef = chartDefaultStyle.Title.FillReference;
                //var myFill = chartDefaultStyle.Title.Fill;
                //var myLine = chartDefaultStyle.Title.Border;

                SaveAndCleanup(p);
            }
        }
    }
}

