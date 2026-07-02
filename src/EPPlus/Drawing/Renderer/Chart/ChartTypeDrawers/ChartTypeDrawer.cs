using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.Svg.Chart.ChartTypeDrawers;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Renderer.Chart.ChartTypeDrawers;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal abstract class ChartTypeDrawer : ChartDrawingObject
    {
        internal protected ExcelChart _chartType;
        internal List<ChartTrendlineRenderer> Trendlines { get; } = new List<ChartTrendlineRenderer>();
        internal ChartErrorBarRenderer ErrorBars { get; private set; }
        internal virtual bool SupportsTrendlines { get { return false; } }
        internal virtual bool SupportsErrorBars { get { return false; } }
        internal virtual bool SupportsUpDownBars { get { return false; } }
        internal virtual bool SupportsDataTable { get { return false; } }
        internal ChartTypeDrawer(ChartRenderer svgChart,  ExcelChart chartType) : base(svgChart)
        {
            _chartType = chartType;
            //Avoid null-refs
            Rectangle = new RectRenderItem(svgChart.Bounds);
        }
        internal abstract void DrawSeries();



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

        public List<RenderItem> SeriesRenderItems { get; } = new List<RenderItem>();
        public List<RenderItem> ChartAreaRenderItems { get; } = new List<RenderItem>();
        internal static List<ChartTypeDrawer> Create(ChartRenderer svgChart)
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
        protected void CreateTrendlines(ExcelChart chartType, List<List<object>> xValues, List<List<object>> yValues)
        {
            var serieIndex = 0;
            foreach (ExcelChartSerie serie in chartType.Series)
            {
                if (serie.TrendLines.Count > 0)
                {
                    var xSerie = xValues[serieIndex];
                    var ySerie = yValues[serieIndex];
                    foreach (var trendline in serie.TrendLines)
                    {
                        var tr = new ChartTrendlineRenderer(ChartRenderer, trendline, xSerie, ySerie, _chartType, serieIndex);
                        Trendlines.Add(tr);
                    }
                }
                serieIndex++;
            }
        }
        protected void CreateErrorBars(ExcelChart chartType, List<List<object>> xValues, List<List<object>> yValues)
        {
            if (!(chartType.IsTypeLine() || chartType.IsTypeColumn() || chartType.IsTypeBar())) return;

            var serieIndex = 0;
            foreach (ExcelChartSerieWithErrorBars serie in chartType.Series)
            {
                if (serie.HasErrorBars())
                {
                    var xSerie = xValues[serieIndex];
                    var ySerie = yValues[serieIndex];
                    ErrorBars = new ChartErrorBarRenderer(ChartRenderer, serie.ErrorBars, xSerie, ySerie, _chartType, serieIndex); 
                }
                serieIndex++;
            }
        }

        internal bool IsOnAxis(ExcelChartAxisStandard ax)
        {
            return _chartType.YAxis==ax || _chartType.XAxis==ax;
        }
        internal static void SetFillDataPoint(ExcelChart chart, ExcelBarChartSerie cStandardSerie, int index, RectRenderItem item, ExcelChartDataPoint dp, ExcelChartStyleEntry entry)
        {
            var theme = chart.WorkSheet.Workbook.ThemeManager.GetOrCreateTheme();
            var color = GetVaryColor(theme, chart.StyleManager?.ColorsManager, index);

            item.SetDrawingPropertiesFill(theme, dp.Fill.IsEmpty ? cStandardSerie.Fill : dp.Fill, entry?.FillReference.Color, false, color);
            item.SetDrawingPropertiesBorder(theme, dp.Border.IsEmpty ? cStandardSerie.Border : dp.Border, entry?.BorderReference.Color, dp.Border.Fill.Style != eFillStyle.NoFill, null, 0.75);
        }

        internal static void SetFillSerie(ExcelChart chart, ExcelChart ct, ExcelChartStandardSerie cStandardSerie, int serieIndex, int index, RenderItem item)
        {
            var theme = chart.WorkSheet.Workbook.ThemeManager.GetOrCreateTheme();
            if (ct.VaryColors)
            {
                //Get the color based on the index, if no style is set. Accent1, Accent2, Accent3...
                var color = GetVaryColor(theme, chart.StyleManager.ColorsManager, index);
                item.SetDrawingPropertiesFill(theme, cStandardSerie.Fill, chart.StyleManager.Style?.SeriesLine.FillReference.Color, false, color);
            }
            else
            {
                var color = GetVaryColor(theme, chart.StyleManager?.ColorsManager, serieIndex);
                item.SetDrawingPropertiesFill(theme, cStandardSerie.Fill, chart.StyleManager.Style?.SeriesLine.FillReference.Color, false, color);
            }
            item.SetDrawingPropertiesBorder(theme, cStandardSerie.Border, chart.StyleManager.Style?.SeriesLine.BorderReference.Color, cStandardSerie.Border.Fill.Style != eFillStyle.NoFill, null, 0.75);
        }

        private static Color? GetVaryColor(ExcelTheme theme, ExcelChartColorsManager colorsManager, int index)
        {
            Color baseColor;
            var baseColorIndex = index % 6;
            if (colorsManager == null || baseColorIndex >= colorsManager.Colors.Count)
            {
                baseColor = theme.ColorScheme.GetColorByEnum(eSchemeColor.Accent1 + baseColorIndex).GetColor();
            }
            else
            {
                baseColor = colorsManager.Colors[baseColorIndex].GetColor();
            }

            var variationIndex = index / 6;
            if (variationIndex == 0)
            {
                return baseColor;
            }
            else
            {
                ExcelColorTransformCollection variation;
                if (colorsManager == null)
                {
                    var variations = ExcelColorTransformCollection.GetDefault();
                    variation = variations[variationIndex % variations.Count];
                }
                else
                {
                    if (colorsManager.Variations.Count == 0)
                    {
                        return baseColor;
                    }
                    variation = colorsManager.Colors[variationIndex % colorsManager.Variations.Count].Transforms;
                }
                return OfficeOpenXml.Utils.TypeConversion.ColorConverter.ApplyTransforms(baseColor, variation);
            }
        }


    }

}
