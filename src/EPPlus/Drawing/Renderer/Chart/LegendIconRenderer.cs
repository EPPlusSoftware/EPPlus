using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.Svg.Chart;
using EPPlus.Export.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Drawing.Renderer.Chart
{
    internal static class LegendIconRenderer
    {
        const float MarginIconText = 1.5f;
        const float MarginHeight = 7.5f;
        const float LineLength = 21;
        const float MinBarLength = 4;
        const float MinPieLength = 5.25f;

        internal static void SetBarLegendSingle<T>(ChartRenderer chart, T parent, double entryWidth, double entryHeight, double maxIconLength) where T : ChartDrawingObject, ILegendKeyContainer
        {
            var ct = chart.Chart.PlotArea.ChartTypes[0];
            var s = ct.Series[0];
            var series = ct.Series[0];
            var catSeries = series.XSeries;
            var catValues = DrawingExtensions.LoadSeriesValues(ct, catSeries, series.NumberLiteralsX, series.StringLiteralsX);

            var valValues = DrawingExtensions.LoadSeriesValues(ct, s.Series, s.NumberLiteralsY, s.StringLiteralsY);

            if (catValues == null || catValues.Count == 0)
            {
                //Blank cat series. Add blank cat
                catValues = new List<object>();
                catValues.Add("");
                parent.Rectangle.Height = entryHeight + MarginHeight;
            }

            var index = 0;
            DrawingLegendSerie pSls = null;

            foreach (var cv in catValues)
            {
                var sls = new DrawingLegendSerie();
                var bs = (ExcelBarChartSerie)s;
                var tm = parent.SeriesHeadersMeasure[index];
                var si = LegendIconRenderer.GetBarSeriesIcon(chart, ct, parent, bs, pSls, entryWidth, entryHeight, 0, index);
                sls.SeriesIcon = si;

                var tbLeft = si.Left + maxIconLength + MarginIconText;
                var tbTop = si.Top - (entryHeight - si.Height) / 2;
                double tbWidth;

                tbWidth = parent.Rectangle.Bounds.Width - tbLeft;

                var tbHeight = tm.Height;
                sls.Textbox = new DrawingTextBody(parent.RenderContext, chart.Chart, parent.Rectangle.Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);
                //sls.Textbox.Bounds.Left = si.Bottom + MarginIconText;

                var headerText = cv.ToString();

                if (parent is ChartLegendRenderer)
                {
                    var entry = chart.Chart.Legend.Entries.FirstOrDefault(x => x.Index == index);
                    if (entry == null || entry.Font.IsEmpty)
                    {
                        //sls.Textbox.AddText(s.GetHeaderText(), sc.Chart.Legend.Font);
                        sls.Textbox.ImportParagraph(chart.Chart.Legend.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
                    }
                    else
                    {
                        //sls.Textbox.AddText(s.GetHeaderText(), entry.Font);
                        sls.Textbox.ImportParagraph(entry.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
                    }
                }
                else if (parent is ChartDataTableRenderer)
                {
                    if(chart.Chart.PlotArea.DataTable.TextBody.Paragraphs.Any())
                    {
                        sls.Textbox.ImportParagraph(chart.Chart.PlotArea.DataTable.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
                    }
                    else
                    {
                        sls.Textbox.AddParagraph(headerText);
                    }
                }
                parent.SeriesIcon.Add(sls);
                pSls = sls;
                index++;
            }
        }

        internal static void SetBarLegend<T>(ChartRenderer chart, T parent, ExcelChart ct, int index, DrawingLegendSerie pSls, ExcelChartSerie s, DrawingLegendSerie sls, double entryWidth, double entryHeight, double maxIconLength) where T : ChartDrawingObject, ILegendKeyContainer
        {
            var bs = (ExcelBarChartSerie)s;
            var tm = parent.SeriesHeadersMeasure[index];
            var si = GetBarSeriesIcon(chart, ct, parent, bs, pSls, entryWidth, entryHeight, index, -1);
            sls.SeriesIcon = si;

            var tbLeft = si.Left + maxIconLength + MarginIconText;
            var tbTop = si.Top - (entryHeight - si.Height) / 2;
            double tbWidth;

            tbWidth = parent.Rectangle.Bounds.Width - tbLeft;

            var tbHeight = tm.Height;
            sls.Textbox = new DrawingTextBody(parent.RenderContext, chart.Chart, parent.Rectangle.Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);
            //sls.Textbox.Bounds.Left = si.Bottom + MarginIconText;

            var entry = parent.Chart.Legend.Entries.FirstOrDefault(x => x.Index == index);
            var headerText = s.GetHeaderText(index);
            if (entry == null || entry.Font.IsEmpty)
            {
                //sls.Textbox.AddText(s.GetHeaderText(), sc.Chart.Legend.Font);
                sls.Textbox.ImportParagraph(chart.Chart.Legend.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            }
            else
            {
                //sls.Textbox.AddText(s.GetHeaderText(), entry.Font);
                sls.Textbox.ImportParagraph(entry.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            }
        }

        internal static void SetLineLegend<T>(ChartRenderer chart, T parent, ExcelChart ct, int index, DrawingLegendSerie pSls, ExcelChartSerie s, DrawingLegendSerie sls, double entryWidth, double entryHeight, double maxIconLength) where T : ChartDrawingObject, ILegendKeyContainer
        {
            var ls = (ExcelLineChartSerie)s;

            var si = GetLineSeriesIcon(chart, parent, ls, pSls, entryWidth, entryHeight);
            sls.SeriesIcon = si;

            var tbLeft = si.X1 + maxIconLength + MarginIconText;
            var tbTop = si.Y2 - entryHeight * 0.5;
            var tbWidth = parent.Rectangle.Bounds.Width - tbLeft;

            var tbHeight = entryHeight;
            sls.Textbox = new DrawingTextBody(parent.RenderContext, chart.Chart, parent.Rectangle.Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);

            var headerText = s.GetHeaderText(index);
            var entry = chart.Chart.Legend.Entries.FirstOrDefault(x => x.Index == index);
            if (entry == null || entry.Font.IsEmpty)
            {
                sls.Textbox.ImportParagraph(chart.Chart.Legend.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            }
            else
            {
                //sls.Textbox.AddText(s.GetHeaderText(), entry.Font);
                sls.Textbox.ImportParagraph(entry.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            }

            if (ls.HasMarker() && ls.Marker.Style != eMarkerStyle.None)
            {
                var l = sls.SeriesIcon as LineRenderItem;
                var x = l.X1 + (l.X2 - l.X1) / 2;
                var y = l.Y1;
                sls.MarkerIcon = LineMarkerHelper.GetMarkerItem(chart, ls, ls.Marker, x, y, true);
                if ((ls.Marker.Style == eMarkerStyle.Plus || ls.Marker.Style == eMarkerStyle.X || ls.Marker.Style == eMarkerStyle.Star) &&
                    ls.Marker.Fill.IsEmpty == false)
                {
                    sls.MarkerBackground = LineMarkerHelper.GetMarkerBackground(chart, ls, x, y, true);
                }
                else
                {
                    sls.MarkerBackground = null;
                }
            }
        }
        internal static void SetPieLegend<T>(ChartRenderer chart, T parent, ExcelChart ct, int index, DrawingLegendSerie pSls, ExcelChartSerie s, DrawingLegendSerie sls, double entryWidth, double entryHeight, double maxIconLength) where T : ChartDrawingObject, ILegendKeyContainer
        {
            var ps = (ExcelPieChartSerie)s;
            pSls = null;

            //Pie chart only cares about series 0
            var series = ct.Series[0];
            var catSeries = series.XSeries;
            var catValues = DrawingExtensions.LoadSeriesValues(ct, catSeries, series.NumberLiteralsX, series.StringLiteralsX);

            //Excel fallsback to index + 1 if no literals and no series 
            if (catValues == null)
            {
                catValues = new List<Object>();
                foreach (var dp in ps.DataPoints)
                {
                    catValues.Add($"{dp.Index + 1}");
                }
            }

            chart.Chart.Legend.TextBody.GetInsetsOrDefaults(out double lDefMargin, out double tDefMarg, out double rDefMargin, out double bDefMarg);

            var widestEntry = entryWidth;

            double lastWidth = 0d;
            double totalWidth = 0d;

            double firstIconWidth = 0d;

            for (int i = 0; i < catValues.Count; i++)
            {
                var tm = parent.SeriesHeadersMeasure[index + i];
                //Step 1: Retrieve Icon

                var si = GetPieSeriesIcon(chart, ct, parent, ps, pSls, lastWidth, entryHeight, i);
                //The si-width is used as left margin for each entry seemingly
                //si.Left += (si.BorderWidth ?? 0d)*2d;
                if (si.Left < 4.5d)
                {
                    si.Left = 4.5d;
                }

                if (i == 0)
                {
                    firstIconWidth = si.Width;
                }
                sls = new DrawingLegendSerie();
                var tbLeft = si.Left + si.Width + MarginIconText + (si.BorderWidth ?? 0d);
                var tbTop = si.Top - ((entryHeight + MarginIconText) / 2);

                double tbWidth;

                if (i != catValues.Count - 1)
                {
                    tbWidth = parent.Rectangle.Bounds.Width - tbLeft;
                }
                else
                {
                    tbWidth = parent.Rectangle.Bounds.Width;
                }

                var tbHeight = tm.Height;
                sls.Textbox = new DrawingTextBody(parent.RenderContext, chart.Chart, parent.Rectangle.Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);
                sls.Textbox.ImportParagraph(chart.Chart.Legend.TextBody.Paragraphs.FirstOrDefault(), 0, catValues[i].ToString());

                //si.Left += sls.Textbox.LeftMargin;

                sls.SeriesIcon = si;
                sls.Textbox.RecalculateParagraphs();

                tbWidth = sls.Textbox.Width + rDefMargin; /*+ lDefMargin + rDefMargin;*/

                lastWidth = tbWidth + si.Width - (si.BorderWidth ?? 0d);

                totalWidth += tbWidth + si.Width + (si.BorderWidth ?? 0d) + MarginIconText;

                if (i >= 0 && ps.DataPoints.ContainsKey(i))
                {
                    var dp = ps.DataPoints[i];
                    ChartTypeDrawer.SetFillDataPoint(chart.Chart, ps, i, sls.SeriesIcon, dp, chart.Chart.StyleManager.Style?.SeriesLine);
                }
                else
                {
                    ChartTypeDrawer.SetFillSerie(chart.Chart, ct, ps, 0, i, sls.SeriesIcon);
                }

                parent.SeriesIcon.Add(sls);
                pSls = sls;
            }

            foreach (var icon in parent.SeriesIcon)
            {
                icon.SeriesIcon.Bounds.Top = icon.SeriesIcon.Bounds.Top - ((entryHeight) / 4);
            }
            var position = chart.Chart.Legend.Position;
            if (position == eLegendPosition.Top || position == eLegendPosition.Bottom)
            {
                //Rectangle.Bounds.Width = totalWidth;
                parent.Rectangle.Bounds.Width = parent.SeriesIcon.Last().Textbox.Bounds.GetGlobalBoundingbox().Right - parent.SeriesIcon[0].SeriesIcon.Bounds.GlobalLeft + 4d + firstIconWidth * 2;
                parent.Rectangle.Bounds.Left = ((chart.Bounds.Width) / 2d) - (totalWidth / 2d) + 1.5d;

                if (parent.Rectangle.Bounds.Width < parent.MaxWidth)
                {
                    parent.Rectangle.Bounds.Height = entryHeight + parent.TopMargin + parent.BottomMargin;
                    parent.Rectangle.Bounds.Top = chart.ChartArea.Rectangle.Height - parent.Rectangle.Height - parent.BottomMargin - parent.TopMargin;
                }
            }
            pSls = null;
            sls = null;
        }

        internal static LineRenderItem GetLineSeriesIcon<T>(ChartRenderer chart, T parent, ExcelChartStandardSerie cStandardSerie, DrawingLegendSerie pSls, double entryWidth, double entryHeight) where T : ChartDrawingObject, ILegendKeyContainer
        {
            var line = new LineRenderItem(parent.Rectangle.Bounds);
            //Default style is NoLine NoFill
            line.SetDrawingPropertiesBorder(chart.Theme, cStandardSerie.Border, chart.Chart.StyleManager.Style?.SeriesLine.BorderReference.Color, cStandardSerie.Border.IsEmpty || cStandardSerie.Border.Fill.Style != eFillStyle.NoFill, () => Color.Empty, 3);
            double iconTop = 0, iconLeft = 0;
            pSls?.GetIconTopLeft(out iconTop, out iconLeft);

            GetItemPosition(chart, parent, pSls, entryWidth, entryHeight, iconLeft, iconTop, out double x, out double y);

            line.X1 = x;
            line.X2 = x + LineLength;
            line.Y1 = y;
            line.Y2 = y;
            line.LineCap = LineCap.Round;

            return line;
        }
        internal static RectRenderItem GetBarSeriesIcon<T>(ChartRenderer chart, ExcelChart ct, T parent, ExcelBarChartSerie chartSerie, DrawingLegendSerie pSls, double entryWidth, double entryHeight, int serieIndex, int index) where T : ChartDrawingObject, ILegendKeyContainer
        {
            var item = new RectRenderItem(parent.Rectangle.Bounds);
            var iconHeight = GetIconLength(ct, entryHeight);
            //var icon = pSls?.SeriesIcon as RectRenderItem;
            double iconTop = 0, iconLeft = 0;
            pSls?.GetIconTopLeft(out iconTop, out iconLeft);

            GetItemPosition(chart, parent, pSls, entryWidth, entryHeight, iconLeft, iconTop + (iconHeight / 2), out double x, out double y);

            item.LineCap = LineCap.Round;
            item.Left = x;
            if (pSls != null && (chart.Chart.Legend.Position == eLegendPosition.Left || chart.Chart.Legend.Position == eLegendPosition.Right))
            {
                item.Top = y - iconHeight / 2d;
            }
            else
            {
                item.Top = y - iconHeight / 2d;
            }
            //item.Top = y;
            item.Width = iconHeight;
            item.Height = iconHeight;

            if (index >= 0 && chartSerie.DataPoints.ContainsKey(index))
            {
                var dp = chartSerie.DataPoints[index];
                ChartTypeDrawer.SetFillDataPoint(chart.Chart, chartSerie, index, item, dp, chart.Chart.StyleManager.Style?.SeriesLine);
            }
            else
            {
                ChartTypeDrawer.SetFillSerie(chart.Chart, ct, chartSerie, serieIndex, index, item);
            }

            return item;
        }
        internal static LineRenderItem GetTrendLineSeriesIcon<T>(ChartRenderer chart, ExcelChart ct, T parent, ExcelChartTrendline tl, DrawingLegendSerie pSls, double entryWidth, double entryHeight) where T : ChartDrawingObject, ILegendKeyContainer
        {
            var line = new LineRenderItem(parent.Rectangle.Bounds);
            line.SetDrawingPropertiesFill(chart.Theme, tl.Fill, chart.Chart.StyleManager.Style?.Trendline.FillReference.Color, UserSpaceSettings.UserSpaceOnUse_Global, parent.DefaultFillColor);
            //Default is actually NoLine
            line.SetDrawingPropertiesBorder(chart.Theme, tl.Border, chart.Chart.StyleManager.Style?.Trendline.BorderReference.Color, tl.Border.Fill.Style != eFillStyle.NoFill, () => parent.DefaultBorderColor, 0.75);
            double iconTop = 0, iconLeft = 0;
            pSls?.GetIconTopLeft(out iconTop, out iconLeft);

            GetItemPosition(chart, parent, pSls, entryWidth, entryHeight, iconLeft, iconTop, out double x, out double y);

            line.X1 = x;
            line.Y1 = y;
            line.X2 = x + LineLength;
            line.Y2 = y;
            line.LineCap = LineCap.Round;

            return line;
        }

        internal static RectRenderItem GetPieSeriesIcon<T>(ChartRenderer chart, ExcelChart ct, T parent, ExcelPieChartSerie pcS, DrawingLegendSerie pSls, double entryWidth, double entryHeight, int i) where T : ChartDrawingObject, ILegendKeyContainer
        {
            var item = new RectRenderItem(parent.Rectangle.Bounds);

            var iconHeight = GetIconLength(ct, entryHeight);
            var icon = pSls?.SeriesIcon as RectRenderItem;

            GetItemPosition(chart, parent, pSls, entryWidth, entryHeight, icon?.Left ?? 0D, icon?.Top ?? 0D, out double x, out double y);

            item.LineCap = LineCap.Round;
            item.Left = x;
            if (pSls != null && (chart.Chart.Legend.Position == eLegendPosition.Left || chart.Chart.Legend.Position == eLegendPosition.Right))
            {
                item.Top = y - (iconHeight / 2d);
            }
            else
            {
                item.Top = y;
            }

            double borderWidth = pcS.Border.Width;

            if (pcS.DataPoints != null && pcS.DataPoints.Count != 0 && i < pcS.DataPoints.Count)
            {
                if (borderWidth < pcS.DataPoints[i].Border.Width)
                {
                    borderWidth = pcS.DataPoints[i].Border.Width;
                }
            }

            //item.Top = y;
            item.Width = iconHeight;
            item.Height = iconHeight;

            item.SetDrawingPropertiesFill(chart.Theme, pcS.Fill, chart.Chart.StyleManager.Style?.SeriesLine.FillReference.Color);
            item.SetDrawingPropertiesBorder(chart.Theme, pcS.Border, chart.Chart.StyleManager.Style?.SeriesLine.BorderReference.Color, pcS.Border.Fill.Style != eFillStyle.NoFill, () => parent.DefaultBorderColor, 1.5d);

            return item;
        }

        internal static double GetIconLength(ExcelChart c, double highestText)
        {
            return c.IsTypeLine() ? LineLength : Math.Max(MinBarLength, highestText * 0.4);
        }

        private static double GetItemPosition<T>(ChartRenderer chart, T parent, DrawingLegendSerie pSls, double entryWidth, double entryHeight, double iconLeft, double iconCenter, out double x, out double y) where T : ChartDrawingObject, ILegendKeyContainer
        {
            var topOffset = 0D;
            if (chart.Chart.Legend.Position == eLegendPosition.Top ||
               chart.Chart.Legend.Position == eLegendPosition.Bottom)
            {
                if (pSls != null && iconLeft + entryWidth * 2 + parent.MarginItemsWidth + parent.RightMargin > parent.MaxWidth)
                {
                    topOffset += entryHeight * 1.25;
                    x = parent.LeftMargin;
                }
                else
                {
                    if (pSls == null)
                    {
                        x = (float)parent.LeftMargin;
                    }
                    else
                    {
                        x = iconLeft + entryWidth + parent.MarginItemsWidth;
                    }
                }

                if (pSls == null)
                {
                    y = parent.TopMargin + entryHeight / 2;
                }
                else
                {
                    y = iconCenter + topOffset;
                }


            }
            else
            {
                if (pSls == null)
                {
                    y = parent.TopMargin + entryHeight / 2;
                }
                else
                {
                    y = iconCenter + entryHeight * 1.5;
                }
                x = parent.LeftMargin;

            }

            return topOffset;
        }

        internal static RenderItem GetSeriesIcon(ChartRenderer chartRenderer, BoundingBox parentItem, ExcelChartStandardSerie s, int index, double entryHeight)
        {
            const float MarginExtra = 1.5f;
            const float DefaultStrokeWidth = 0.75f;

            var theme = s._chart.WorkSheet.Workbook.ThemeManager.GetOrCreateTheme();

            if(s._chart.IsTypeLine())
            {
                var item = new LineRenderItem(parentItem);
                item.SetDrawingPropertiesFill(theme, s.Fill, chartRenderer.Chart.StyleManager.Style.SeriesLine.FillReference.Color, UserSpaceSettings.ObjectBoundingBox);
                item.SetDrawingPropertiesBorder(theme, s.Border, chartRenderer.Chart.StyleManager.Style.SeriesLine.BorderReference.Color, s.Border.Fill.Style != eFillStyle.NoFill, null, DefaultStrokeWidth, UserSpaceSettings.ObjectBoundingBox);

                float y = (float)parentItem.Top + MarginExtra;
                float x = 0;
                item.X1 = x;
                item.Y1 = y;
                item.X2 = x + (LineLength - (float)item.BorderWidth);
                item.Y2 = y;
                item.LineCap = LineCap.Round;
                return item;
            }
            else
            {
                var item = new RectRenderItem(parentItem);
                var iconHeight = GetIconLength(chartRenderer.Chart, entryHeight);
                item.Width = iconHeight;
                item.Height = iconHeight;
                var chart = chartRenderer.Chart;
                var bs = (ExcelBarChartSerie)s;
                if (index >= 0 && bs.DataPoints.ContainsKey(index))
                {
                    var dp = bs.DataPoints[index];
                    ChartTypeDrawer.SetFillDataPoint(chart, s, index, item, dp, chart.StyleManager.Style?.SeriesLine);
                }
                else
                {
                    ChartTypeDrawer.SetFillSerie(chart, s._chart, s, index, index, item);
                }

                return item;
            }


        }

    }
}
