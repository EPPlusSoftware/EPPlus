/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.Svg.Chart;
using EPPlus.Export.Utils;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Renderer.Chart;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
namespace EPPlusImageRenderer.Svg
{
    internal interface ILegendKeyContainer
    {
        float MarginItemsWidth { get; }
        double MaxWidth { get; }
        double MaxHeight { get; }
        List<TextMeasurement> SeriesHeadersMeasure { get; }
        List<DrawingLegendSerie> SeriesIcon { get; }
    }
    internal class ChartLegendRenderer : ChartDrawingObject, ILegendKeyContainer
    {
        
        List<TextMeasurement> _seriesHeadersMeasure = new List<TextMeasurement>();
        ITextMeasurer _ttMeasurer;
        const float MarginIconText = 1.5f;
        const float MarginHeight = 7.5f;
        const float LineLength = 21;
        const float MinBarLength = 4;
        float MinPieLength = 5.25f;
        float _marginItemsWidth;
        double _maxWidth, _maxHeight;
        internal ChartLegendRenderer(ChartRenderer sc) : base(sc)
        {

            var mf = Chart.Font.GetMeasureFont();
            var shaper = RenderContext.FontEngine.GetShaperForFont(mf);
            var _ttMeasurer = new OpenTypeFontTextMeasurer(shaper);

            if (sc.Chart.HasLegend == false && sc.Chart.Series.Count == 0)
            {
                return;
            }

            var l = ((ExcelChartStandard)sc.Chart).Legend;

            LeftMargin = RightMargin = 3; //4px
            TopMargin = BottomMargin = 3; //4px
            _marginItemsWidth = mf.Size / 2; //We use half the size of the font as margin between items.

            switch (l.Position)
            {
                case eLegendPosition.Top:
                case eLegendPosition.Bottom:
                    if(sc.Chart.IsTypePie())
                    {
                        _maxWidth = sc.ChartArea.Rectangle.Width * 0.95d;
                    }
                    else
                    {
                        _maxWidth = sc.ChartArea.Rectangle.Width * 0.85d;
                    }
                    _maxHeight = sc.ChartArea.Rectangle.Height * 0.6d;
                    break;
                default:
                    _maxWidth = sc.ChartArea.Rectangle.Width * 0.6d;
                    _maxHeight = sc.ChartArea.Rectangle.Height * 0.85d;
                    break;
            }
            double entryWidth, entryHeight;

            Rectangle = GetLegendRectangleAndEntrySize(l, out entryWidth, out entryHeight);

            if (l.Layout.HasLayout) //Manual layout will override the position and size of legend, but not the entry size which is used for calculating the position of legend entries.
            {
                Rectangle = GetRectFromManualLayout(sc, l.Layout);
            }

            Rectangle.SetDrawingPropertiesFill(sc.Theme, l.Fill, sc.Chart.StyleManager.Style?.Title.FillReference.Color, UserSpaceSettings.UserSpaceOnUse_Global, DefaultFillColor);
            Rectangle.SetDrawingPropertiesBorder(sc.Theme, l.Border, sc.Chart.StyleManager.Style?.Legend.BorderReference.Color, l.Border.Fill.Style != eFillStyle.NoFill, () => DefaultBorderColor, 0.75);
            
            var pSls = SetLegendSeries(entryWidth, entryHeight);
            SetLegendTrendlines(entryWidth, entryHeight, pSls);
        }

        private RectRenderItem GetLegendRectangleAndEntrySize(ExcelChartLegend l, out double entryWidth, out double entryHeight)
        {
            //var rect = new RectRenderItem(RectanBounds);
            var rect = Rectangle = new RectRenderItem(ChartRenderer.Bounds);
            var widest = 0d;
            var highest = 0d;
            var index = 0;

            //Find the widest and hightest legend entry, and calculate the total width and hight of the legend based on the orientation. 
            foreach (var ct in Chart.PlotArea.ChartTypes)
            {
                if (ct.GetType() == typeof(ExcelPieChart))
                {
                    //Pie chart cares only about first series
                    if (ct.Series[0].GetType() == typeof(ExcelPieChartSerie))
                    {
                        var ps = (ExcelPieChartSerie)ct.Series[0];
                        var catValues = DrawingExtensions.LoadSeriesValues(ct, ps.XSeries, ps.NumberLiteralsX, ps.StringLiteralsX);

                        //Excel fallsback to index + 1 if no literals and no series 
                        if (catValues == null)
                        {
                            catValues = new List<Object>();
                            foreach (var dp in ps.DataPoints)
                            {
                                catValues.Add($"{dp.Index + 1}");
                            }
                        }
                        for (int i = 0; i < catValues.Count; i++)
                        {
                            var text = catValues[i].ToString();
                            GetSerieSize(l, index, text, ref widest, ref highest);
                            index++;
                        }
                        Chart.Legend.TextBody.GetInsetsOrDefaults(out double lDefMargin, out double tDefMarg, out double rDefMargin, out double bDefMarg);

                        //In Excel VBA the margins for legend.TextFrame2 appear to always be 7.2:
                        //RightMargin = rDefMargin;
                        LeftMargin = lDefMargin;

                        //widest += rDefMargin;

                        //if (MinPieLength < highest * 0.5d)
                        //{
                        //    MinPieLength = (float)(highest * 0.5d);
                        //}
                    }
                    //Skip the rest
                    break;
                }
                //Single series bar/column chart, the legend entries are the categories, not the series. //J
                //---
                //That's not true if Both Series and XSeries exists on e.g a clustered column chart.
                //Uncertain of what scenario this is helpful for. Most bar and column charts don't appear to work this way even with a single series
                //It IS the correct way for Pie charts and for a series with only a single point in it and no xSSeries so perhaps that?
                //See EpplusTest.Export.SvgExport tests especially DtPt tests. //O
                else if ((ct.IsTypeColumn() || ct.IsTypeBar()) && Chart.PlotArea.ChartTypes.Count == 1 && ct.Series.Count == 1)
                {
                    var s = ct.Series[0];
                    var catSeries = s.XSeries;

                    var catValues = DrawingExtensions.LoadSeriesValues(ct, catSeries, s.NumberLiteralsX, s.StringLiteralsX);
                    var valValues = DrawingExtensions.LoadSeriesValues(ct, s.Series, s.NumberLiteralsY, s.StringLiteralsY);

                    //If there is no xSeries but there ARE SerieValues we MUST use the serie values
                    //Same vice-versa
                    if (catValues == null)
                    {
                        if(valValues != null)
                        {
                            catValues = valValues;
                        }
                    }

                    if(catValues != null)
                    {
                        for (int i = 0; i < catValues.Count; i++)
                        {
                            var text = catValues[i].ToString();
                            GetSerieSize(l, index, text, ref widest, ref highest);
                            index++;
                        }
                    }
                }
                else
                {
                    foreach (var s in ct.Series)
                    {
                        var text = s.GetHeaderText(index);
                        GetSerieSize(l, index, text, ref widest, ref highest);

                        if(highest == 0)
                        {
                            highest = MinBarLength;
                        }
                        if(widest == 0)
                        {
                            widest = MinBarLength;
                        }

                        index++;
                    }
                }
            }

            //Trendlines also get legend entries, but they should appear after the series name.
            var trIndex = 0;
            foreach (var ct in Chart.PlotArea.ChartTypes)
            {
                foreach (var s in ct.Series)
                {
                    foreach (var tl in s.TrendLines)
                    {
                        var text = tl.GetName(index);
                        GetSerieSize(l, trIndex, text, ref widest, ref highest);
                        trIndex++;
                    }
                }
            }

            index += trIndex;

            var maxIconLength = GetMaxIconLength(Chart, highest);
            entryWidth = maxIconLength + MarginIconText + widest;
            entryHeight = highest;

            switch (l.Position)
            {
                case eLegendPosition.Top:
                case eLegendPosition.Bottom:
                    var fullLength = LeftMargin + entryWidth * index + _marginItemsWidth * (index - 1) + RightMargin;
                    if(fullLength > _maxWidth)
                    {
                        var height = entryHeight * 0.25;
                        var widestLine = 0D;
                        var width = LeftMargin + entryWidth;
                        
                        for(int i = 0; i < index; i++)
                        {
                            if (width + entryWidth + RightMargin > _maxWidth)
                            {
                                height += entryHeight * 1.25;
                                if (width + RightMargin > widestLine)
                                {
                                    widestLine = width + RightMargin;
                                }
                                width = RightMargin + entryWidth;
                            }
                            else
                            {
                                width += entryWidth + _marginItemsWidth;
                            }
                        }

                        //height+= BottomMargin;
                        rect.Width = Math.Max(widestLine, width);
                        rect.Height = height + entryHeight * 1.25; 
                    }
                    else
                    {
                        rect.Width = fullLength;
                        rect.Height = entryHeight * 1.5;
                    }
                    rect.Left = (ChartRenderer.ChartArea.Rectangle.Width - rect.Width) / 2;
                    if (l.Position == eLegendPosition.Top)
                    {                        
                        rect.Top = ChartRenderer.Title.Rectangle.Bottom + MarginHeight;
                    }
                    else
                    {
                        rect.Top = ChartRenderer.ChartArea.Rectangle.Height - rect.Height - BottomMargin - TopMargin;
                    }
                    break;
                case eLegendPosition.Right:
                case eLegendPosition.TopRight:
                case eLegendPosition.Left:
                    rect.Width = LeftMargin + entryWidth + RightMargin;
                    rect.Height = TopMargin + (entryHeight * index) + ((index - 1) * entryHeight * 0.5) + BottomMargin; //use margin as 50% of the entry height and to the top and the bottom.;

                    if (rect.Height > _maxHeight)
                    {
                        rect.Height = _maxHeight;
                    }

                    if (l.Position == eLegendPosition.Right ||
                        l.Position == eLegendPosition.TopRight)
                    {
                        rect.Left = ChartRenderer.ChartArea.Rectangle.Width - rect.Width - LeftMargin;
                    }
                    else
                    {
                        rect.Left = LeftMargin + 2;
                    }
                    if (l.Position == eLegendPosition.Left ||
                        l.Position == eLegendPosition.Right)
                    {
                        //Will be set when the plotarea width is calculated.
                        //rect.Top = sc.ChartArea.Rectangle.Height / 2 - rect.Height / 2;
                    }
                    else
                    {
                        if (ChartRenderer.Title == null)
                        {
                            rect.Top = 8 + 8;
                        }
                        else
                        {
                            rect.Top = ChartRenderer.Title.Rectangle.Height + 8 + 8; //Height + Margin Top and Bottom Title
                        }
                    }
                    break;
            }

            return rect;
        }

        private Color GetAccentBasedOnPos(int pos)
        {
            var mod6 = pos % 6;
            //TODO: Only works for base-case. Add support for patterns 1,3 and 4 instead of just 2 as basecase
            return ChartRenderer.Theme.ColorScheme.GetColorByEnum(eSchemeColor.Accent1 + mod6).GetColor();
        }

        private void GetSerieSize(ExcelChartLegend l, int index, string text, ref double widest, ref double highest)
        {
            var entry = l.Entries.FirstOrDefault(x => x.Index == index);
            ExcelTextFont font;
            MeasurementFont mf;
            if (entry == null || entry.Font.IsEmpty)
            {
                font = l.Font;
                mf = l.Font.GetMeasureFont();
            }
            else
            {
                font = entry.Font;
                mf = entry.Font.GetMeasureFont();
            }

            if (_ttMeasurer == null)
            {
                _ttMeasurer = new OpenTypeFontTextMeasurer(RenderContext.FontEngine.GetShaperForFont(mf));
            }

            TextMeasurement tm;

            if(text == "")
            {
                //We need at least SOME sort of height or width based on font
                tm = _ttMeasurer.MeasureText("|", mf);
            }
            else
            {
                tm = _ttMeasurer.MeasureText(text, mf);
            }

            _seriesHeadersMeasure.Add(tm);

            if (tm.Width > widest)
            {
                widest = tm.Width;
            }

            if (tm.Height > highest)
            {
                highest = tm.Height;
            }
        }

        private double GetMaxIconLength(ExcelChart ct, double highestText)
        {
            var maxIconLength = 0D;
            foreach(var c in ct.PlotArea.ChartTypes)
            {
                var il = LegendIconRenderer.GetIconLength(c, highestText);
                if (il > maxIconLength)
                {
                    maxIconLength = il;
                }
                if (c.ChartType == eChartType.Pie && maxIconLength < MinPieLength)
                {
                    maxIconLength = MinPieLength;
                }
            }
            return maxIconLength;
        }


        internal DrawingLegendSerie SetLegendSeries(double entryWidth, double entryHeight)
        {
            int index = 0;
            DrawingLegendSerie pSls = null;
            var pos = Chart.Legend.Position;
            var maxIconLength = GetMaxIconLength(Chart, entryHeight);
            if ((Chart.IsTypeBar() || Chart.IsTypeColumn()) && Chart.PlotArea.ChartTypes.Count == 1 && Chart.PlotArea.ChartTypes[0].Series.Count == 1)
            {
                LegendIconRenderer.SetBarLegendSingle(ChartRenderer, this, entryWidth, entryHeight, maxIconLength);
            }
            else
            {
                foreach (var ct in Chart.PlotArea.ChartTypes)
                {
                    int ix, end;
                    if (ct.IsTypeBar())
                    {
                        ix = ct.Series.Count - 1;
                        end = -1;
                    }
                    else
                    {
                        ix = 0;
                        end = ct.Series.Count;
                    }

                    while (ix != end)
                    {
                        var s = ct.Series[ix];
                        var sls = new DrawingLegendSerie();

                        switch (ct.ChartType)
                        {
                            case eChartType.Line:
                            case eChartType.LineMarkers:
                            case eChartType.LineMarkersStacked:
                            case eChartType.LineMarkersStacked100:
                            case eChartType.LineStacked:
                            case eChartType.LineStacked100:
                                LegendIconRenderer.SetLineLegend(ChartRenderer,this, ct, index, pSls, s, sls, entryWidth, entryHeight, maxIconLength);
                                break;
                            case eChartType.ColumnClustered:
                            case eChartType.ColumnStacked:
                            case eChartType.ColumnStacked100:
                            case eChartType.BarClustered:
                            case eChartType.BarStacked:
                            case eChartType.BarStacked100:
                                LegendIconRenderer.SetBarLegend(ChartRenderer, this, ct, index, pSls, s, sls, entryWidth, entryHeight, maxIconLength);
                                break;
                            case eChartType.Pie:
                            case eChartType.PieExploded:
                                if (ix == 0)
                                {
                                    LegendIconRenderer.SetPieLegend(ChartRenderer, this, ct, index, pSls, s, sls, entryWidth, entryHeight, maxIconLength);
                                    pSls = null;
                                    sls = null;
                                }
                                break;
                            default:
                                break;
                        }
                        if (Chart.Legend.Position == eLegendPosition.Top ||
                            Chart.Legend.Position == eLegendPosition.Bottom)
                        {
                            //if (sls.Textbox.Bounds.Bottom > Rectangle.Bottom)
                            //{
                            //    break;
                            //}
                        }
                        else
                        {
                            if (sls != null && sls.Textbox.Bounds.Bottom > Rectangle.Height)
                            {
                                break;
                            }
                        }
                        if (sls != null)
                        {
                            SeriesIcon.Add(sls);
                        }
                        //SeriesIcon.Add(sls);
                        pSls = sls;
                        //else
                        //{
                        //    pSls = null;
                        //}
                        index++;
                        if (ix < end)
                        {
                            ix++;
                        }
                        else
                        {
                            ix--;
                        }
                    }
                }
            }
            return pSls;
        }

        private void SetLegendTrendlines(double entryWidth, double entryHeight, DrawingLegendSerie pSls)
        {
            int index = SeriesIcon.Count;
            var pos = Chart.Legend.Position;
            foreach (var ct in Chart.PlotArea.ChartTypes)
            {
                int ix, end;
                if (ct.IsTypeBar())
                {
                    ix = ct.Series.Count - 1;
                    end = -1;
                }
                else
                {
                    ix = 0;
                    end = ct.Series.Count;
                }

                while (ix != end)
                {
                    var s = ct.Series[ix];
                    foreach (var tl in s.TrendLines)
                    {
                        var sls = new DrawingLegendSerie();

                        SetTrendlineLegend(ct, ix, index, pSls, pos, tl, sls, entryWidth, entryHeight);

                        if (sls.Textbox.Bounds.Bottom > Rectangle.Height)
                        {
                            return;
                        }

                        SeriesIcon.Add(sls);
                        pSls = sls;
                        index++;

                    }
                    if (ix < end)
                    {
                        ix++;
                    }
                    else
                    {
                        ix--;
                    }
                }
            }
        }
        private void SetTrendlineLegend(ExcelChart ct, int serieIndex, int entryIndex, DrawingLegendSerie pSls, eLegendPosition pos, ExcelChartTrendline tl, DrawingLegendSerie sls, double entryWidth, double entryHeight)
        {

            var si = LegendIconRenderer.GetTrendLineSeriesIcon(ChartRenderer, ct, this, tl, pSls, entryWidth, entryHeight);
            sls.SeriesIcon = si;

            var tbLeft = si.X1 + LineLength + MarginIconText;
            var tbTop = si.Y2 - entryHeight * 0.5;    //TODO:Should probably be font ascent 
            double tbWidth;
            tbWidth = Rectangle.Bounds.Width - tbLeft;

            var tbHeight = entryHeight;
            sls.Textbox = new DrawingTextBody(RenderContext, Chart, Rectangle.Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);

            var entry = Chart.Legend.Entries.FirstOrDefault(x => x.Index == entryIndex);
            var headerText = tl.GetName(serieIndex);
            if (entry == null || entry.Font.IsEmpty)
            {
                sls.Textbox.ImportParagraph(Chart.Legend.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            }
            else
            {
                sls.Textbox.ImportParagraph(entry.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            }
        }
        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var groupItem = new GroupRenderItem(ChartRenderer.Bounds);
            groupItem.Top = Rectangle.Bounds.Top;
            groupItem.Left = Rectangle.Bounds.Left;
            renderItems.Add(groupItem);

            //The rectangle is position using the group transform, so we need to set the rectangle position to 0,0
            Rectangle.Bounds.Top = 0;
            Rectangle.Bounds.Left = 0;

            groupItem.RenderItems.Add(Rectangle);
            foreach(var s in SeriesIcon)
            {
                if(s.SeriesIcon != null) groupItem.RenderItems.Add(s.SeriesIcon);
                if(s.MarkerBackground != null) groupItem.RenderItems.Add(s.MarkerBackground);
                if (s.MarkerIcon != null) groupItem.RenderItems.Add(s.MarkerIcon);
                //renderItems.Add(s.Textbox);
                if(s.Textbox != null) s.Textbox.AppendRenderItems(groupItem.RenderItems);
            }
        }
        internal override Color? DefaultFillColor => Color.Transparent;

        internal override Color? DefaultBorderColor => Color.Transparent;

        public float MarginItemsWidth => _marginItemsWidth;

        public double MaxWidth => _maxWidth;

        public double MaxHeight => _maxHeight;

        public List<TextMeasurement> SeriesHeadersMeasure => _seriesHeadersMeasure;

        public List<DrawingLegendSerie> SeriesIcon { get; } = new List<DrawingLegendSerie>();

    }
}