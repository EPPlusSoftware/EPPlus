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
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.Svg;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace EPPlusImageRenderer.Svg
{
    internal class ChartLegendRenderer : ChartDrawingObject
    {
        
        List<TextMeasurement> _seriesHeadersMeasure = new List<TextMeasurement>();
        ITextMeasurer _ttMeasurer;
        const float MarginIconText = 1.5f;
        const float MarginHeight = 7.5f;
        const float LineLength = 21;
        const float MinBarLength = 4;
        float _marginItemsWidth;

        double _maxWidth, _maxHeight;
        internal ChartLegendRenderer(ChartRenderer sc, bool isDataLabelLegend = false) : base(sc)
        {
            var mf = Chart.Font.GetMeasureFont();
            var shaper = OpenTypeFonts.GetShaperForFont(mf);
            var _ttMeasurer = new OpenTypeFontTextMeasurer(shaper);

            if (sc.Chart.HasLegend == false && isDataLabelLegend == false || sc.Chart.Series.Count == 0)
            {
                return;
            }

            var l = ((ExcelChartStandard)sc.Chart).Legend;

            LeftMargin = RightMargin = 3; //4px
            TopMargin = BottomMargin = 3; //4px
            _marginItemsWidth = mf.Size; //We use the size of the font as margin between items.
            switch (l.Position)
            {
                case eLegendPosition.Top:
                case eLegendPosition.Bottom:
                    _maxWidth = sc.ChartArea.Rectangle.Width * 0.8;
                    _maxHeight = sc.ChartArea.Rectangle.Height * 0.6;
                    break;
                default:
                    _maxWidth = sc.ChartArea.Rectangle.Width * 0.6;
                    _maxHeight = sc.ChartArea.Rectangle.Height * 0.8;
                    break;
            }
            double entryWidth, entryHeight;

            Rectangle = GetLegendRectangleAndEntrySize(l, out entryWidth, out entryHeight);

            if (l.Layout.HasLayout) //Manual layout will override the position and size of legend, but not the entry size which is used for calculating the position of legend entries.
            {
                Rectangle = GetRectFromManualLayout(sc, l.Layout);
            }

            //Bounds.Left = Rectangle.Left;
            //Bounds.Top = Rectangle.Top;
            //Bounds.Width = Rectangle.Width;
            //Bounds.Height = Rectangle.Height;
            //Rectangle.Bounds.Left = Rectangle.Bounds.Top = 0;

            Rectangle.SetDrawingPropertiesFill(sc.Theme, l.Fill, sc.Chart.StyleManager.Style.Title.FillReference.Color);
            Rectangle.SetDrawingPropertiesBorder(sc.Theme, l.Border, sc.Chart.StyleManager.Style.Title.BorderReference.Color, l.Border.Fill.Style != eFillStyle.NoFill, 0.75);

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
                foreach (var s in ct.Series)
                {
                    var text = s.GetHeaderText(index);
                    GetSerieSize(l, index, text, ref widest, ref highest);
                    index++;
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

            var maxIconLength = GetMaxIconLenght(Chart, highest);
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
                                width = RightMargin + widest;
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
                        rect.Top = ChartRenderer.ChartArea.Rectangle.Height - rect.Height - BottomMargin;
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
                        rect.Left = ChartRenderer.ChartArea.Rectangle.Width - rect.Width - TopMargin;
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
                _ttMeasurer = new OpenTypeFontTextMeasurer(OpenTypeFonts.GetShaperForFont(mf));
            }

            var tm = _ttMeasurer.MeasureText(text, mf);
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

        private double GetMaxIconLenght(ExcelChart ct, double heighestText)
        {
            var maxIconLength = 0D;
            foreach(var c in ct.PlotArea.ChartTypes)
            {
                var il = GetIconLenght(c, heighestText);
                if (il > maxIconLength)
                {
                    maxIconLength = il;
                }
            }
            return maxIconLength;
        }
        private double GetIconLenght(ExcelChart c, double heighestText)
        {
            return c.IsTypeLine() ? LineLength : Math.Max(MinBarLength, heighestText * 0.4);
        }


        internal DrawingLegendSerie SetLegendSeries(double entryWidth, double entryHeight)
        {
            int index = 0;
            DrawingLegendSerie pSls=null;
            var pos = Chart.Legend.Position;
            var maxIconLength = GetMaxIconLenght(Chart, entryHeight);
            foreach (var ct in Chart.PlotArea.ChartTypes)
            {
                int ix, end;
                if(ct.IsTypeBar())
                {
                    ix = ct.Series.Count-1;
                    end = -1;
                }
                else
                {
                    ix = 0;
                    end = ct.Series.Count;
                }

                while(ix != end)
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
                            SetLineLegend(ct, index, pSls, pos, s, sls, entryWidth, entryHeight, maxIconLength);
                            break;
                        case eChartType.ColumnClustered:
                        case eChartType.ColumnStacked:
                        case eChartType.ColumnStacked100:
                        case eChartType.BarClustered:
                        case eChartType.BarStacked:
                        case eChartType.BarStacked100:
                            SetBarLegend(ct, index, pSls, pos, s, sls, entryWidth, entryHeight, maxIconLength);
                            break;
                        case eChartType.Pie:
                        case eChartType.PieExploded:
                            SetPieLegend(ct, index, pSls, pos, s, sls, entryWidth, entryHeight, maxIconLength);
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
                        if (sls.Textbox.Bounds.Bottom > Rectangle.Height)
                        {
                            break;
                        }
                    }
                    SeriesIcon.Add(sls);
                    pSls = sls;
                    index++;
                    if(ix<end)
                    {
                        ix++;
                    }
                    else 
                    {
                        ix--;
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

            var si = GetTrendLineSeriesIcon(ct, tl, pSls, entryWidth, entryHeight);
            sls.SeriesIcon = si;

            var tbLeft = si.X1 + LineLength + MarginIconText;
            var tbTop = si.Y2 - entryHeight * 0.5;    //TODO:Should probably be font ascent 
            double tbWidth;
            tbWidth = Rectangle.Bounds.Width - tbLeft;

            var tbHeight = entryHeight;
            sls.Textbox = new DrawingTextbody(Chart, Rectangle.Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);

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


        private void SetPieLegend(ExcelChart ct, int index, DrawingLegendSerie pSls, eLegendPosition pos, ExcelChartSerie s, DrawingLegendSerie sls, double entryWidth, double entryHeight, double maxIconLength)
        {
            var ps = (ExcelPieChartSerie)s;

            var si = GetLineSeriesIcon(ct, ps, pSls, entryWidth, entryHeight);
            sls.SeriesIcon = si;

            var tbLeft = si.X1 + maxIconLength + MarginIconText;
            var tbTop = si.Y2 - entryHeight * 0.5;
            var tbWidth = Rectangle.Bounds.Width - tbLeft;

            var tbHeight = entryHeight;
            sls.Textbox = new DrawingTextbody(Chart, Rectangle.Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);

            //Cat values are the header text
            //They create a rect marker for each slice

            //var entry = Chart.Legend.Entries.FirstOrDefault(x => x.Index == index);
            //var catValues = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);
            //for (int i = 0; i< ps.NumberOfItems; i++)
            //{
               
            //    var headerText = ps.XSeries
            //}
            //var headerText = s.GetHeaderText(index);
            //if (entry == null || entry.Font.IsEmpty)
            //{
            //    sls.Textbox.ImportParagraph(sc.Chart.Legend.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            //}
            //else
            //{
            //    //sls.Textbox.AddText(s.GetHeaderText(), entry.Font);
            //    sls.Textbox.ImportParagraph(entry.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            //}

            //if (ps.DataPoints != null && ps.DataPoints.Count != null /*&& ps.Marker.Style != eMarkerStyle.None*/)
            //{
            //    var l = sls.SeriesIcon as SvgRenderLineItem;
            //    var x = l.X1 + (l.X2 - l.X1) / 2;
            //    var y = l.Y1;

            //    //sls.MarkerIcon = LineMarkerHelper.GetMarkerItem(sc, ps, x, y, true);
            //    if ((ps.Marker.Style == eMarkerStyle.Plus || ps.Marker.Style == eMarkerStyle.X || ps.Marker.Style == eMarkerStyle.Star) &&
            //        ps.Marker.Fill.IsEmpty == false)
            //    {
            //        sls.MarkerBackground = LineMarkerHelper.GetMarkerBackground(sc, ps, x, y, true);
            //    }
            //    else
            //    {
            //        sls.MarkerBackground = null;
            //    }
            //}
        }

        private void SetLineLegend(ExcelChart ct, int index, DrawingLegendSerie pSls, eLegendPosition pos, ExcelChartSerie s, DrawingLegendSerie sls, double entryWidth, double entryHeight, double maxIconLength)
        {
            var ls = (ExcelLineChartSerie)s;

            var si = GetLineSeriesIcon(ct, ls, pSls, entryWidth, entryHeight);
            sls.SeriesIcon = si;

            var tbLeft = si.X1 + maxIconLength + MarginIconText;
            var tbTop = si.Y2 - entryHeight * 0.5;
            var tbWidth = Rectangle.Bounds.Width - tbLeft;

            var tbHeight = entryHeight;
            sls.Textbox = new DrawingTextbody(Chart, Rectangle.Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);

            var entry = Chart.Legend.Entries.FirstOrDefault(x => x.Index == index);
            var headerText = s.GetHeaderText(index);
            if (entry == null || entry.Font.IsEmpty)
            {
                sls.Textbox.ImportParagraph(Chart.Legend.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
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
                sls.MarkerIcon = LineMarkerHelper.GetMarkerItem(ChartRenderer, ls, x, y, true);
                if ((ls.Marker.Style == eMarkerStyle.Plus || ls.Marker.Style == eMarkerStyle.X || ls.Marker.Style == eMarkerStyle.Star) &&
                    ls.Marker.Fill.IsEmpty == false)
                {
                    sls.MarkerBackground = LineMarkerHelper.GetMarkerBackground(ChartRenderer, ls, x, y, true);
                }
                else
                {
                    sls.MarkerBackground = null;
                }
            }
        }

        private void SetBarLegend(ExcelChart ct, int index, DrawingLegendSerie pSls, eLegendPosition pos, ExcelChartSerie s, DrawingLegendSerie sls, double entryWidth, double entryHeight, double maxIconLength)
        {
            var bs = (ExcelBarChartSerie)s;
            var tm = _seriesHeadersMeasure[index];
            var si = GetBarSeriesIcon(ct, bs, pSls, entryWidth, entryHeight);
            sls.SeriesIcon = si;

            var tbLeft = si.Left + maxIconLength + MarginIconText;
            var tbTop = si.Top - (entryHeight - si.Height) / 2; 
            double tbWidth;

            tbWidth = Rectangle.Bounds.Width - tbLeft;

            var tbHeight = tm.Height;
            sls.Textbox = new DrawingTextbody(Chart, Rectangle.Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);
            //sls.Textbox.Bounds.Left = si.Bottom + MarginIconText;

            var entry = Chart.Legend.Entries.FirstOrDefault(x => x.Index == index);
            var headerText = s.GetHeaderText(index);
            if (entry == null || entry.Font.IsEmpty)
            {
                //sls.Textbox.AddText(s.GetHeaderText(), sc.Chart.Legend.Font);
                sls.Textbox.ImportParagraph(Chart.Legend.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            }
            else
            {
                //sls.Textbox.AddText(s.GetHeaderText(), entry.Font);
                sls.Textbox.ImportParagraph(entry.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            }
        }

        private LineRenderItem GetLineSeriesIcon(ExcelChart ct, ExcelChartStandardSerie cStandardSerie, DrawingLegendSerie pSls, double entryWidth, double entryHeight)
        {
            var line = new LineRenderItem(Rectangle.Bounds);
            line.SetDrawingPropertiesFill(ChartRenderer.Theme, cStandardSerie.Fill, Chart.StyleManager.Style.SeriesLine.FillReference.Color);
            line.SetDrawingPropertiesBorder(ChartRenderer.Theme, cStandardSerie.Border, Chart.StyleManager.Style.SeriesLine.BorderReference.Color, cStandardSerie.Border.Fill.Style != eFillStyle.NoFill, 0.75);
            double iconTop = 0, iconLeft = 0;
            pSls?.GetIconTopLeft(out iconTop, out iconLeft);

            GetItemPosition(pSls, entryWidth, entryHeight, iconLeft, iconTop, out double x, out double y);

            line.X1 = x;
            line.X2 = x + LineLength;
            line.Y1 = y;
            line.Y2 = y;
            line.LineCap = LineCap.Round;

            return line;
        }
        private LineRenderItem GetTrendLineSeriesIcon(ExcelChart ct, ExcelChartTrendline tl, DrawingLegendSerie pSls, double entryWidth, double entryHeight)
        {
            var line = new LineRenderItem(Rectangle.Bounds);
            line.SetDrawingPropertiesFill(ChartRenderer.Theme, tl.Fill, Chart.StyleManager.Style.Trendline.FillReference.Color);
            line.SetDrawingPropertiesBorder(ChartRenderer.Theme, tl.Border, Chart.StyleManager.Style.Trendline.BorderReference.Color, tl.Border.Fill.Style != eFillStyle.NoFill, 0.75);
            double iconTop = 0, iconLeft = 0;
            pSls?.GetIconTopLeft(out iconTop, out iconLeft);

            GetItemPosition(pSls, entryWidth, entryHeight, iconLeft, iconTop, out double x, out double y);

            line.X1 = x;
            line.Y1 = y;
            line.X2 = x + LineLength;
            line.Y2 = y;
            line.LineCap = LineCap.Round;

            return line;
        }

        private RectRenderItem GetBarSeriesIcon(ExcelChart ct, ExcelChartStandardSerie cStandardSerie, DrawingLegendSerie pSls, double entryWidth, double entryHeight)
        {            
            var item = new RectRenderItem(Rectangle.Bounds);
            var iconHeight = GetIconLenght(ct, entryHeight);
            //var icon = pSls?.SeriesIcon as RectRenderItem;
            double iconTop = 0, iconLeft = 0;
            pSls?.GetIconTopLeft(out iconTop, out iconLeft);

            GetItemPosition(pSls, entryWidth, entryHeight, iconTop, iconTop + (iconHeight / 2), out double x, out double y);

            item.LineCap = LineCap.Round;
            item.Left = x;
            if(pSls !=null && (Chart.Legend.Position == eLegendPosition.Left || Chart.Legend.Position == eLegendPosition.Right))
            {
                item.Top = y - iconHeight / 2;
            }
            else
            {
                item.Top = y - iconHeight / 2;
            }
            //item.Top = y;
            item.Width = iconHeight;
            item.Height = iconHeight;

            item.SetDrawingPropertiesFill(ChartRenderer.Theme, cStandardSerie.Fill, Chart.StyleManager.Style.SeriesLine.FillReference.Color);
            item.SetDrawingPropertiesBorder(ChartRenderer.Theme, cStandardSerie.Border, Chart.StyleManager.Style.SeriesLine.BorderReference.Color, cStandardSerie.Border.Fill.Style != eFillStyle.NoFill, 0.75);

            return item;
        }

        private double GetItemPosition(DrawingLegendSerie pSls, double entryWidth, double entryHeight, double iconLeft, double iconCenter, out double x, out double y)
        {
            var topOffset = 0D;
            if (Chart.Legend.Position == eLegendPosition.Top ||
               Chart.Legend.Position == eLegendPosition.Bottom)
            {
                if (pSls != null && pSls.Textbox.Bounds.Right + entryWidth + RightMargin > _maxWidth)
                {
                    topOffset += entryHeight * 1.25;
                    x = LeftMargin;
                }
                else
                {
                    if (pSls == null)
                    {
                        x = (float)LeftMargin;
                    }
                    else
                    {
                        x = iconLeft + entryWidth + _marginItemsWidth;
                    }
                }

                if (pSls == null)
                {
                    y = TopMargin + entryHeight / 2;
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
                    y = TopMargin + entryHeight / 2;
                }
                else
                {
                    y = iconCenter + entryHeight * 1.5;
                }
                x = LeftMargin;

            }

            return topOffset;
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

        public List<DrawingLegendSerie> SeriesIcon { get; } = new List<DrawingLegendSerie>();

    }
}