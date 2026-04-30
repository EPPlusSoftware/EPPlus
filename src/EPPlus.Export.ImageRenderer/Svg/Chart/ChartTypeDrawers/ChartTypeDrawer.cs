using EPPlus.Export.ImageRenderer.Svg.Chart.ChartTypeDrawers;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal abstract class ChartTypeDrawer : SvgChartObject
    {
        protected SvgChart _svgChart;
        protected ExcelChart _chartType;
        internal virtual bool SupportsTrendlines { get { return false; } }
        internal virtual bool SupportsErrorBars { get { return false; } }
        internal virtual bool SupportsUpDownBars { get { return false; } }
        internal virtual bool SupportsDataTable { get { return false; } }
        internal ChartTypeDrawer(SvgChart svgChart,  ExcelChart chartType) : base(svgChart)
        {
            _svgChart = svgChart;
            _chartType = chartType;
        }

        protected List<object> LoadSeriesValues(string serieAddressInput, double[] numLiterals, string[] strLiterals)
        {
            string serieAddress = serieAddressInput;

            //Some addresses are split and within parenthesis
            if (serieAddressInput.StartsWith("("))
            {
                serieAddress = serieAddressInput.Trim('(', ')');
            }

            List<object> values = new List<object>();
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
                if (string.IsNullOrEmpty(serieAddress))
                {
                    return null;
                }
                var address = new ExcelAddressBase(serieAddress);

                if (address.Addresses != null && address.Addresses.Count > 1)
                {
                    foreach (var splitAddress in address.Addresses)
                    {
                        FillValuesFromAddress(splitAddress, ref values);
                    }
                }
                else
                {
                    FillValuesFromAddress(address, ref values);
                }
            }
            return values;
        }

        protected void FillValuesFromAddress(ExcelAddressBase address, ref List<object> values)
        {
            if (address.IsExternal)
            {
                var wb = Chart.WorkSheet.Workbook;
                var extWb = wb.ExternalLinks[address.ExternalReferenceIndex - 1] as ExcelExternalWorkbook;
                if (extWb != null)
                {
                    var wsName = address.WorkSheetName;
                    if (extWb.Package == null)
                    {
                        var extWs = extWb.CachedWorksheets[wsName];
                        FillExternalValues(extWs, address, ref values);
                    }
                    else
                    {
                        var ws = extWb.Package.Workbook.Worksheets[wsName];
                        FillInternalValues(ws, address, ref values);
                    }
                }
            }
            else
            {
                var wsName = address.WorkSheetName;

                if (string.IsNullOrEmpty(wsName))
                {
                    wsName = Chart.WorkSheet.Name;
                }

                var ws = Chart.WorkSheet.Workbook.Worksheets[wsName];
                FillInternalValues(ws, address, ref values);
            }
        }

        protected void FillExternalValues(ExcelExternalWorksheet extWs, ExcelAddressBase address, ref List<object> values)
        {
            if (extWs != null)
            {
                for (int r = address.Start.Row; r <= address.End.Row; r++)
                {
                    for (int c = address.Start.Column; c <= address.End.Column; c++)
                    {
                        values.Add(extWs.CellValues[r, c].Value);
                    }
                }
            }
        }

        private void FillInternalValues(ExcelWorksheet ws, ExcelAddressBase address, ref List<object> values)
        {

            if (ws != null)
            {
                for (int r = address.Start.Row; r <= address.End.Row; r++)
                {
                    for (int c = address.Start.Column; c <= address.End.Column; c++)
                    {
                        values.Add(ws.Cells[r, c].Value);
                    }
                }
            }
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
                        drawers.Add(new LineChartTypeDrawer(svgChart, (ExcelLineChart)ct));
                        break;
                    case eChartType.ColumnClustered:
                    case eChartType.ColumnStacked:
                    case eChartType.ColumnStacked100:
                    case eChartType.BarClustered:
                    case eChartType.BarStacked:
                    case eChartType.BarStacked100:
                        drawers.Add(new BarColumnChartTypeDrawer(svgChart, (ExcelBarChart)ct));
                        break;
                    case eChartType.Pie:
                    case eChartType.PieExploded:
                        drawers.Add(new PieChartTypeDrawer(svgChart, (ExcelPieChart)ct));
                        break;
                    default:
                        throw new NotImplementedException($"No Svg support for Chart type {ct} is implemented.");
                }
            }
            return drawers;
        }
        internal void SumSeries(List<List<object>> series)
        {
            for (var i = 1; i < series.Count; i++)
            {
                for (var j = 0; j < series[i].Count; j++)
                {
                    series[i][j] = ConvertUtil.GetValueDouble(series[i][j]) + ConvertUtil.GetValueDouble(series[i - 1][j]);
                }
            }
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.AddRange(RenderItems);
        }
    }

}
