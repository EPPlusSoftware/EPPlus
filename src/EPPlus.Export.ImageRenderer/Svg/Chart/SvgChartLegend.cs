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
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Svg;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartLegend : SvgChartObject
    {
        
        List<TextMeasurement> _seriesHeadersMeasure = new List<TextMeasurement>();
        ITextMeasurer _ttMeasurer;
        const float MarginIconText = 1.5f;
        const float MarginItemsWidth = 3f;
        const float MarginHeight = 7.5f;
        const float LineLength = 21;
        const float MinBarLength = 4;
        double _maxWidth, _maxHeight;
        internal SvgChartLegend(SvgChart sc, bool isDataLabelLegend = false) : base(sc)
        {
            var mf = sc.Chart.Font.GetMeasureFont();
            var shaper = OpenTypeFonts.GetShaperForFont(mf);
            var _ttMeasurer = new OpenTypeFontTextMeasurer(shaper);

            if (sc.Chart.HasLegend == false && isDataLabelLegend == false || sc.Chart.Series.Count == 0)
            {
                return;
            }
            var l = ((ExcelChartStandard)sc.Chart).Legend;

            LeftMargin = RightMargin = 3; //4px
            TopMargin = BottomMargin = 3; //4px
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

            Rectangle = GetLegendRectangleAndEntrySize(sc, l, out entryWidth, out entryHeight);

            if (l.Layout.HasLayout) //Manual layout will override the position and size of legend, but not the entry size which is used for calculating the position of legend entries.
            {
                Rectangle = GetRectFromManualLayout(sc, l.Layout);
            }

            Bounds.Left = Rectangle.Left;
            Bounds.Top = Rectangle.Top;
            Bounds.Width = Rectangle.Width;
            Bounds.Height = Rectangle.Height;
            Rectangle.Bounds.Left = Rectangle.Bounds.Top = 0;

            Rectangle.SetDrawingPropertiesFill(l.Fill, sc.Chart.StyleManager.Style.Title.FillReference.Color);
            Rectangle.SetDrawingPropertiesBorder(l.Border, sc.Chart.StyleManager.Style.Title.BorderReference.Color, l.Border.Fill.Style != eFillStyle.NoFill, 0.75);

            SetLegend(sc, entryWidth, entryHeight);
        }

        private SvgRenderRectItem GetLegendRectangleAndEntrySize(SvgChart sc, ExcelChartLegend l, out double entryWidth, out double entryHeight)
        {
            var rect = new SvgRenderRectItem(sc, sc.Bounds);
            
            var widest = 0d;
            var highest = 0d;
            var index = 0;
            //Find the widest and hightest legend entry, and calculate the total width and hight of the legend based on the orientation. 
            foreach (var ct in sc.Chart.PlotArea.ChartTypes)
            {
                foreach (var s in ct.Series)
                {
                    var text = s.GetHeaderText(index);
                    var entry = l.Entries.FirstOrDefault(x => x.Index == index);
                    ExcelTextFont font;
                    MeasurementFont mf;
                    if(entry==null || entry.Font.IsEmpty)
                    {
                        font = l.Font;
                        mf = l.Font.GetMeasureFont();
                    }
                    else
                    {
                        font = entry.Font;
                        mf = entry.Font.GetMeasureFont();
                    }

                    if(_ttMeasurer == null)
                    {
                        _ttMeasurer = new OpenTypeFontTextMeasurer(OpenTypeFonts.GetShaperForFont(mf));
                    }

                    var tm = _ttMeasurer.MeasureText(text, mf);
                    _seriesHeadersMeasure.Add(tm);

                    if(tm.Width > widest)
                    {
                        widest = tm.Width;
                    }

                    if(tm.Height> highest)
                    {
                        highest = tm.Height;
                    }

                    index++;
                }
            }
            var maxIconLength = GetMaxIconLenght(sc.Chart, highest);
            entryWidth = maxIconLength + MarginIconText + widest;
            entryHeight = highest;
            //hight += BottomMargin;     //remove last margin and add bottom margin
            switch (l.Position)
            {
                case eLegendPosition.Top:
                case eLegendPosition.Bottom:
                    var fullLength = LeftMargin + entryWidth * index + MarginItemsWidth * (index - 1) + RightMargin;
                    if(fullLength > _maxWidth)
                    {
                        var height = entryHeight*0.25;
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
                                width += entryWidth + MarginItemsWidth;
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
                    rect.Left = (sc.ChartArea.Rectangle.Width - rect.Width) / 2;
                    if (l.Position == eLegendPosition.Top)
                    {                        
                        rect.Top = sc.Title.Rectangle.Bottom + MarginHeight;
                    }
                    else 
                    {
                        rect.Top = sc.ChartArea.Rectangle.Height - rect.Height - BottomMargin;
                    }
                    break;
                case eLegendPosition.Right:
                case eLegendPosition.TopRight:
                case eLegendPosition.Left:
                    rect.Width = LeftMargin + entryWidth + RightMargin; 
                    rect.Height = (entryHeight * index) + ((index - 1) * MarginItemsWidth) + entryHeight;

                    if (rect.Height > _maxHeight)
                    {
                        rect.Height = _maxHeight;
                    }

                    if (l.Position == eLegendPosition.Right ||
                        l.Position == eLegendPosition.TopRight)
                    {
                        rect.Left = sc.ChartArea.Rectangle.Width - rect.Width - TopMargin;
                    }
                    else
                    {
                        rect.Left = LeftMargin + 2;
                    }
                    if (l.Position == eLegendPosition.Left ||
                        l.Position == eLegendPosition.Right)
                    {
                        rect.Top = sc.ChartArea.Rectangle.Height / 2 - rect.Height / 2;
                    }
                    else
                    {
                        if (sc.Title == null)
                        {
                            rect.Top = 8 + 8;
                        }
                        else
                        {
                            rect.Top = sc.Title.Rectangle.Height + 8 + 8; //Height + Margin Top and Bottom Title
                        }
                    }
                    break;
            }

            return rect;
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


        internal void SetLegend(SvgChart sc, double entryWidth, double entryHeight)
        {
            int index = 0;
            SvgLegendSerie pSls=null;
            var pos = Chart.Legend.Position;
            var maxIconLength = GetMaxIconLenght(sc.Chart, entryHeight);
            foreach (var ct in sc.Chart.PlotArea.ChartTypes)
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
                    var sls = new SvgLegendSerie();
                    switch (ct.ChartType)
                    {
                        case eChartType.Line:
                        case eChartType.LineMarkers:
                        case eChartType.LineMarkersStacked:
                        case eChartType.LineMarkersStacked100:
                        case eChartType.LineStacked:
                        case eChartType.LineStacked100:
                            SetLineLegend(sc, ct, index, pSls, pos, s, sls, entryWidth, entryHeight, maxIconLength);
                            break;
                        case eChartType.ColumnClustered:
                        case eChartType.ColumnStacked:
                        case eChartType.ColumnStacked100:
                        case eChartType.BarClustered:
                        case eChartType.BarStacked:
                        case eChartType.BarStacked100:
                            SetBarLegend(sc, ct, index, pSls, pos, s, sls, entryWidth, entryHeight, maxIconLength);
                            break;
                        default:
                            break;
                    }
                    if (sc.Chart.Legend.Position == eLegendPosition.Top ||
                        sc.Chart.Legend.Position == eLegendPosition.Bottom)
                    {
                        //if (sls.Textbox.Bounds.Bottom > Rectangle.Bottom)
                        //{
                        //    break;
                        //}
                    }
                    else
                    {
                        if (sls.Textbox.Bounds.Bottom > Rectangle.Bottom)
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
        }

        private void SetLineLegend(SvgChart sc, ExcelChart ct, int index, SvgLegendSerie pSls, eLegendPosition pos, ExcelChartSerie s, SvgLegendSerie sls, double entryWidth, double entryHeight, double maxIconLength)
        {
            var ls = (ExcelLineChartSerie)s;

            var si = GetLineSeriesIcon(sc, ct, ls, pSls, entryWidth, entryHeight);
            sls.SeriesIcon = si;

            var tbLeft = si.X1 + maxIconLength + MarginIconText;
            var tbTop = si.Y2 - entryHeight * 0.5;    //TODO:Should probably be font ascent 
            double tbWidth;
            if (pos == eLegendPosition.Left || pos == eLegendPosition.Right)
            {
                tbWidth = Bounds.Width - tbLeft - RightMargin;
            }
            else
            {
                tbWidth = Bounds.Width - tbLeft - RightMargin;
            }

            var tbHeight = entryHeight;
            sls.Textbox = new SvgTextBodyItem(ChartRenderer, Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);

            var entry = Chart.Legend.Entries.FirstOrDefault(x => x.Index == index);
            var headerText = s.GetHeaderText(index);
            if (entry == null || entry.Font.IsEmpty)
            {
                sls.Textbox.ImportParagraph(sc.Chart.Legend.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            }
            else
            {
                //sls.Textbox.AddText(s.GetHeaderText(), entry.Font);
                sls.Textbox.ImportParagraph(entry.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            }

            if (ls.HasMarker() && ls.Marker.Style != eMarkerStyle.None)
            {
                var l = sls.SeriesIcon as SvgRenderLineItem;
                var x = l.X1 + (l.X2 - l.X1) / 2;
                var y = l.Y1;
                sls.MarkerIcon = LineMarkerHelper.GetMarkerItem(sc, ls, x, y, true);
                if ((ls.Marker.Style == eMarkerStyle.Plus || ls.Marker.Style == eMarkerStyle.X || ls.Marker.Style == eMarkerStyle.Star) &&
                    ls.Marker.Fill.IsEmpty == false)
                {
                    sls.MarkerBackground = LineMarkerHelper.GetMarkerBackground(sc, ls, x, y, true);
                }
                else
                {
                    sls.MarkerBackground = null;
                }
            }
        }

        private void SetBarLegend(SvgChart sc, ExcelChart ct, int index, SvgLegendSerie pSls, eLegendPosition pos, ExcelChartSerie s, SvgLegendSerie sls, double entryWidth, double entryHeight, double maxIconLength)
        {
            var bs = (ExcelBarChartSerie)s;
            var tm = _seriesHeadersMeasure[index];
            var si = GetBarSeriesIcon(sc, ct, bs, pSls, entryWidth, entryHeight);
            sls.SeriesIcon = si;

            var tbLeft = si.Left + maxIconLength + MarginIconText;
            var tbTop = si.Top - (entryHeight - si.Height) / 2; 
            double tbWidth;

            tbWidth = Bounds.Width - tbLeft - RightMargin;

            var tbHeight = tm.Height;
            sls.Textbox = new SvgTextBodyItem(ChartRenderer, Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);
            //sls.Textbox.Bounds.Left = si.Bottom + MarginIconText;

            var entry = Chart.Legend.Entries.FirstOrDefault(x => x.Index == index);
            var headerText = s.GetHeaderText(index);
            if (entry == null || entry.Font.IsEmpty)
            {
                //sls.Textbox.AddText(s.GetHeaderText(), sc.Chart.Legend.Font);
                sls.Textbox.ImportParagraph(sc.Chart.Legend.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            }
            else
            {
                //sls.Textbox.AddText(s.GetHeaderText(), entry.Font);
                sls.Textbox.ImportParagraph(entry.TextBody.Paragraphs.FirstOrDefault(), 0, headerText);
            }
        }

        private SvgRenderLineItem GetLineSeriesIcon(SvgChart sc, ExcelChart ct, ExcelChartStandardSerie cStandardSerie, SvgLegendSerie pSls, double entryWidth, double entryHeight)
        {
            var line = new SvgRenderLineItem(sc, Rectangle.Bounds);
            line.SetDrawingPropertiesFill(cStandardSerie.Fill, sc.Chart.StyleManager.Style.SeriesLine.FillReference.Color);
            line.SetDrawingPropertiesBorder(cStandardSerie.Border, sc.Chart.StyleManager.Style.SeriesLine.BorderReference.Color, cStandardSerie.Border.Fill.Style != eFillStyle.NoFill, 0.75);
            var icon = pSls?.SeriesIcon as SvgRenderLineItem;

            GetItemPosition(sc, pSls, entryWidth, entryHeight, icon?.X1 ?? 0D, icon?.Y1 ?? 0D, out double x, out double y);

            line.X1 = x;
            line.Y1 = y;
            line.X2 = x + LineLength;
            line.Y2 = y;
            line.LineCap = eLineCap.Round;

            return line;
        }

        private SvgRenderRectItem GetBarSeriesIcon(SvgChart sc, ExcelChart ct, ExcelChartStandardSerie cStandardSerie, SvgLegendSerie pSls, double entryWidth, double entryHeight)
        {            
            var item = new SvgRenderRectItem(sc, Rectangle.Bounds);
            var iconHeight = GetIconLenght(ct, entryHeight);
            var icon = pSls?.SeriesIcon as SvgRenderRectItem;

            GetItemPosition(sc, pSls, entryWidth, entryHeight, icon?.Left ?? 0D, icon?.Top ?? 0D, out double x, out double y);

            item.LineCap = eLineCap.Round;
            item.Left = x;
            item.Top = y;
            item.Width = iconHeight;
            item.Height = iconHeight;

            item.SetDrawingPropertiesFill(cStandardSerie.Fill, sc.Chart.StyleManager.Style.SeriesLine.FillReference.Color);
            item.SetDrawingPropertiesBorder(cStandardSerie.Border, sc.Chart.StyleManager.Style.SeriesLine.BorderReference.Color, cStandardSerie.Border.Fill.Style != eFillStyle.NoFill, 0.75);

            return item;
        }

        private double GetItemPosition(SvgChart sc, SvgLegendSerie pSls, double entryWidth, double entryHeight, double iconLeft, double iconTop, out double x, out double y)
        {
            var topOffset = 0D;
            if (sc.Chart.Legend.Position == eLegendPosition.Top ||
               sc.Chart.Legend.Position == eLegendPosition.Bottom)
            {
                if (pSls != null && pSls.Textbox.Bounds.Right + entryWidth + RightMargin > _maxWidth)
                {
                    topOffset += entryHeight * 1.25;
                    x = Rectangle.Left + LeftMargin;
                }
                else
                {
                    if (pSls == null)
                    {
                        x = Rectangle.Left + (float)LeftMargin;
                    }
                    else
                    {
                        x = iconLeft + entryWidth + MarginItemsWidth;
                    }
                }
                if (pSls == null)
                {
                    y = Rectangle.Top + entryHeight / 2;
                }
                else
                {
                    y = iconTop + topOffset;
                }


            }
            else
            {
                if (pSls == null)
                {
                    y = entryHeight / 2;
                }
                else
                {
                    y = pSls.Textbox.Bounds.Top + entryHeight * 1.25;
                }
                x = LeftMargin;

            }

            return topOffset;
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var groupItem = new SvgGroupItem(ChartRenderer, Bounds);
            renderItems.Add(groupItem);
            renderItems.Add(Rectangle);
            foreach(var s in SeriesIcon)
            {
                if(s.SeriesIcon != null) renderItems.Add(s.SeriesIcon);
                if(s.MarkerBackground != null) renderItems.Add(s.MarkerBackground);
                if (s.MarkerIcon != null) renderItems.Add(s.MarkerIcon);
                //renderItems.Add(s.Textbox);
                if(s.Textbox != null) s.Textbox.AppendRenderItems(renderItems);
            }
            renderItems.Add(new SvgEndGroupItem(ChartRenderer, null));
        }

        public List<SvgLegendSerie> SeriesIcon { get; } = new List<SvgLegendSerie>();

    }
}