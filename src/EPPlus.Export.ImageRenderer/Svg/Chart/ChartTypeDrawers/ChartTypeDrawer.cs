using EPPlus.Export.ImageRenderer.Utils;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal abstract class ChartTypeDrawer : SvgChartObject
    {
        protected SvgChart _svgChart;
        protected ExcelChart _chartType;
        internal ChartTypeDrawer(SvgChart svgChart,  ExcelChart chartType) : base(svgChart)
        {
            _svgChart = svgChart;
            _chartType = chartType;
        }
        protected List<object> LoadSeriesValues(string serieAddress, double[] numLiterals, string[] strLiterals)
        {
            List<object> values=new List<object>();
            if (numLiterals != null)
            {
                values.AddRange(numLiterals.Select(x => (object)x));
            }
            else if (strLiterals != null)
            {
                values.AddRange(strLiterals.Select(x => (object)x));
            }
            else
            {
                if(string.IsNullOrEmpty(serieAddress))
                {
                    return null;
                }
                var address = new ExcelAddressBase(serieAddress);
                var wsName = address.WorkSheetName;
                if (string.IsNullOrEmpty(wsName))
                {
                    wsName = Chart.WorkSheet.Name;
                }
                if (Chart.WorkSheet.Workbook.Worksheets[wsName] != null)
                {
                    for (int r = address.Start.Row; r <= address.End.Row; r++)
                    {
                        for (int c = address.Start.Column; c <= address.End.Column; c++)
                        {
                            values.Add(Chart.WorkSheet.Workbook.Worksheets[wsName].Cells[r, c].Value);
                        }
                    }
                }
            }
            return values; 
        }
        public List<RenderItem> RenderItems { get; } = new List<RenderItem>();
        internal static List<ChartTypeDrawer> Create(SvgChart svgChart)
        {
            var drawers = new List<ChartTypeDrawer>();
            foreach (var ct in svgChart.Chart.PlotArea.ChartTypes)
            {
                switch (ct.ChartType)
                {
                    case eChartType.Line:
                    case eChartType.LineMarkers:
                    case eChartType.LineStacked:
                    case eChartType.LineStacked100:
                    case eChartType.LineMarkersStacked:
                    case eChartType.LineMarkersStacked100:
                        drawers.Add(new LineChartTypeDrawer(svgChart, ct));
                        break;
                    default:
                        throw new NotImplementedException($"No Svg support for Chart type {ct} is implemented.");
                }
            }
            return drawers;
        }
        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.AddRange(RenderItems);
        }
    }

}
