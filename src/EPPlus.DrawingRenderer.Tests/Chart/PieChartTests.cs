using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Tests.Chart
{
    [TestClass]
    public class PieChartTests : TestBase
    {

        [TestMethod]
        public void ReadAndCreateSvgsAll()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PieChartSvgALL.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                for (int i = 4; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        if(c.Name == "ManySlices_Rot")
                        {
                            var svg = c.ToSvg();
                            SaveTextFileToWorkbook($"svg\\PieChartSvgALL\\s{i}_{ws.Name}_{c.Name}.svg", svg);
                        }
                    }
                }
            }
        }


        [TestMethod]
        public void SimpleBestFit()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("SimpleBestFit.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                for (int i = 0; i < p.Workbook.Worksheets.Count; i++)
                {
                    ws = p.Workbook.Worksheets[i];
                    foreach (ExcelChart c in ws.Drawings)
                    {
                        var svg = c.ToSvg();
                        SaveTextFileToWorkbook($"svg\\s{i}_{ws.Name}_{c.Name}.svg", svg);
                    }
                }
            }
        }
    }
}
