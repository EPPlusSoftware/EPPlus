using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.Svg.Chart.ChartTypeDrawers;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml;
using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Renderer.Chart.ChartTypeDrawers;
using OfficeOpenXml.Drawing.Renderer.Chart.Defaults;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal abstract class ChartTypeDrawer : ChartDrawingObjectWithBackground
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
        /// <summary>
        /// Returns the serie count for the chart type and previouse chart types. This is used to determin the color of the series.
        /// </summary>
        /// <param name="startSerie"></param>
        /// <returns></returns>
        internal abstract int DrawSeries(int startSerie);



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

        public List<Transform> SeriesRenderItems { get; } = new List<Transform>();
        public List<Transform> ChartAreaRenderItems { get; } = new List<Transform>();
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
                        throw new NotImplementedException($"No Svg support for Chart type {ct.ChartType} is implemented.");
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

        internal int GetVariationCount(ExcelChartStandardSerie cStandardSerie)
        {
            var valValues = LoadSeriesValues(cStandardSerie.Series, cStandardSerie.NumberLiteralsY, cStandardSerie.StringLiteralsY);
            var variationCount = valValues.Count / 6;
            //Compare these two and take the higher count later? We should standardize loading series values more
            //var catValues = LoadSeriesValues(cStandardSerie.XSeries, cStandardSerie.NumberLiteralsX, cStandardSerie.StringLiteralsX);

            return variationCount;
        }

        internal static void SetFillDataPoint(ExcelChart chart, ExcelChartStandardSerie cStandardSerie, int index, RenderItemStyle style, ExcelChartDataPoint dp, ExcelChartStyleEntry entry, int numDataPoints, UserSpaceSettings spaceSettings = UserSpaceSettings.UserSpaceOnUse_Object)
        {
            var legacyVariationCount = numDataPoints / 6;
            var theme = chart.WorkSheet.Workbook.ThemeManager.GetOrCreateTheme();
            var color = GetVaryColor(theme, chart.StyleManager, index, legacyVariationCount);

            style.SetDrawingPropertiesFill(theme, dp.Fill.IsEmpty ? cStandardSerie.Fill : dp.Fill, entry?.FillReference.Color, spaceSettings, color);
            style.SetDrawingPropertiesBorder(theme, dp.Border.IsEmpty ? cStandardSerie.Border : dp.Border, entry?.BorderReference.Color, dp.Border.Fill.Style != eFillStyle.NoFill, () => Color.Transparent, 0.75);
        }

        internal static void SetSingleFillSerieLegacy(ExcelChart chart, ExcelChart ct, ExcelChartStandardSerie cStandardSerie, int valueCount, int index, RenderItemStyle style)
        {
            var theme = chart.WorkSheet.Workbook.ThemeManager.GetOrCreateTheme();
            var baseColor = GetBaseColor(theme, chart.StyleManager, index);

            var currentIterationIndex = index / 6;
            var totalNumberOfVariations = (valueCount / 6);

            var variation = DefaultDataPointPatterns.GetSingleLegacyColorVariation(currentIterationIndex, totalNumberOfVariations, baseColor);
            style.SetDrawingPropertiesFill(theme, cStandardSerie.Fill, chart.StyleManager.Style?.SeriesLine.FillReference.Color, UserSpaceSettings.UserSpaceOnUse_Object, variation);

            var styleFormat = DefaultChartBorder.GetDefaulDataPoint2d(theme, chart, currentIterationIndex);
            style.SetDrawingPropertiesBorder(theme, cStandardSerie.Border, chart.StyleManager.Style?.SeriesLine.BorderReference.Color, cStandardSerie.Border.Fill.Style != eFillStyle.NoFill || styleFormat.width > 0, () => styleFormat.color, 0.75);
        }

        internal static void SetFillSerie(ExcelChart chart, ExcelChart ct, ExcelChartStandardSerie cStandardSerie, int serieIndex, int index, RenderItemStyle style, int numValuesInSeries)
        {
            var theme = chart.WorkSheet.Workbook.ThemeManager.GetOrCreateTheme();
            var legacyVariationCount = numValuesInSeries / 6;
            if (ct.VaryColors)
            {
                var chartStyle =(int)chart.Style;
                //Get the color based on the index, if no style is set. Accent1, Accent2, Accent3...
                var color = GetVaryColor(theme, chart.StyleManager, index, legacyVariationCount);
                style.SetDrawingPropertiesFill(theme, cStandardSerie.Fill, chart.StyleManager.Style?.SeriesLine.FillReference.Color, UserSpaceSettings.UserSpaceOnUse_Object, color);
            }
            else
            {
                var color = GetVaryColor(theme, chart.StyleManager, serieIndex, legacyVariationCount);
                style.SetDrawingPropertiesFill(theme, cStandardSerie.Fill, chart.StyleManager.Style?.SeriesLine.FillReference.Color, UserSpaceSettings.UserSpaceOnUse_Object, color);
            }
            var styleFormat = DefaultChartBorder.GetDefaulDataPoint2d(theme, chart, serieIndex);
            style.SetDrawingPropertiesBorder(theme, cStandardSerie.Border, chart.StyleManager.Style?.SeriesLine.BorderReference.Color, cStandardSerie.Border.Fill.Style != eFillStyle.NoFill || styleFormat.width>0, () => styleFormat.color, 0.75);
        }

        internal static Color GetBaseColor(ExcelTheme theme, ExcelChartStyleManager? styleManager, int index)
        {
            Color baseColor;
            var baseColorIndex = index % 6;

            ExcelChartColorsManager colorsManager = styleManager?.ColorsManager;
            if ((colorsManager == null || baseColorIndex >= colorsManager.Colors.Count))
            {
                if(styleManager.Style == null)
                {
                    baseColor = DefaultDataPointPatterns.FallBackColors[baseColorIndex];
                }
                else
                {
                    //TODO: Check Chart.Style here. Only when THAT is set to None should the below line happen
                    baseColor = theme.ColorScheme.GetColorByEnum(eSchemeColor.Accent1 + baseColorIndex).GetColor();
                }
            }
            else
            {
                baseColor = colorsManager.Colors[baseColorIndex].GetColor();
            }
            return baseColor;
        }

        internal static Color? GetVaryColor(ExcelTheme theme, ExcelChartStyleManager? styleManager, int index, int legacyVariationCount)
        {
            Color baseColor;
            var baseColorIndex = index % 6;
            
            ExcelChartColorsManager colorsManager = styleManager?.ColorsManager;
            if ((colorsManager == null || baseColorIndex >= colorsManager.Colors.Count))
            {
                if (styleManager.Style != null)
                {
                    //TODO: Check Chart.Style here. Only when THAT is set to None should the below line happen
                    baseColor = theme.ColorScheme.GetColorByEnum(eSchemeColor.Accent1 + baseColorIndex).GetColor();
                }
                else
                {
                    baseColor = DefaultDataPointPatterns.FallBackColors[baseColorIndex];
                }
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
                    if(styleManager.Style == null)
                    {
                        var variations = ExcelColorTransformCollection.GetLegacyDefaultPattern2(legacyVariationCount);
                        variation = variations[variationIndex % variations.Count];
                    }
                    else
                    {
                        var variations = ExcelColorTransformCollection.GetDefault();
                        variation = variations[variationIndex % variations.Count];
                    }
                }
                else
                {
                    if (colorsManager.Variations.Count == 0)
                    {
                        return baseColor;
                    }

                    int colorCount = -1;

                    if (colorsManager.Variations.Count > colorsManager.Colors.Count)
                    {
                        colorCount = colorsManager.Colors.Count;
                    }
                    else
                    {
                        colorCount = colorsManager.Variations.Count;
                    }

                    variation = colorsManager.Colors[variationIndex % colorCount].Transforms;
                }
                return OfficeOpenXml.Utils.TypeConversion.ColorConverter.ApplyTransforms(baseColor, variation);
            }
        }
        protected double GetAxisBaseY(ChartAxisRenderer catAx, ChartAxisRenderer valAx)
        {

            if (catAx.Axis.Crosses == eCrosses.AutoZero)
            {
                if(catAx.Axis.CrossesAt.HasValue)
                {
                    var cv = catAx.Axis.CrossesAt.Value;
                    if (cv < valAx.Min) cv = valAx.Min;
                    if (cv > valAx.Max) cv = valAx.Max;
                    
                    return valAx.GetPositionInPlotarea(cv, true);
                }
                else
                {
                    return valAx.GetPositionInPlotarea(valAx.Min <= 0 ? 0D : valAx.Min, true);
                }
            }
            else if (catAx.Axis.Crosses == eCrosses.Min)
            {
                return valAx.GetPositionInPlotarea(valAx.Min, true);
            }
            else
            {
                return valAx.GetPositionInPlotarea(valAx.Max, true);
            }
        }
    }
}
