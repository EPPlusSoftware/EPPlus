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

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartLegend : SvgChartObject
    {
        
        List<TextMeasurement> _seriesHeadersMeasure = new List<TextMeasurement>();
        ITextMeasurer _ttMeasurer;
        const float MarginExtra = 1.5f;
        const float MiddleMargin = 7.5f;
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
            var iconLengh = GetIconLenght(sc, highest);
            entryWidth = iconLengh + MarginExtra + widest;
            entryHeight = highest;
            //hight += BottomMargin;     //remove last margin and add bottom margin
            switch (l.Position)
            {
                case eLegendPosition.Top:
                case eLegendPosition.Bottom:
                    var fullLength = LeftMargin + entryWidth * index + MarginExtra*(index-1) + RightMargin;
                    if(fullLength > _maxWidth)
                    {
                        var height = TopMargin + highest;
                        var widestLine = 0D;
                        var width = LeftMargin + entryWidth;
                        
                        for(int i = 0; i < index; i++)
                        {
                            if (width + entryWidth + RightMargin > _maxWidth)
                            {
                                height = height + highest;
                                if (width + RightMargin > widestLine)
                                {
                                    widestLine = width + RightMargin;
                                }
                                width = RightMargin + widest;
                            }
                            else
                            {
                                width += entryWidth + MarginExtra;
                            }
                        }

                        height+= BottomMargin;
                        rect.Width = Math.Max(widestLine, width);
                        rect.Height = height + BottomMargin;
                    }
                    else
                    {
                        rect.Width = fullLength;
                        rect.Height = TopMargin + highest + BottomMargin;
                    }
                    rect.Left = (sc.ChartArea.Rectangle.Width - rect.Width) / 2;
                    if (l.Position == eLegendPosition.Top)
                    {                        
                        rect.Top = sc.Title.Rectangle.Bottom + MiddleMargin;
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
                    rect.Height = TopMargin + (highest * index) + ((index - 1) * MarginExtra) + BottomMargin;

                    if (rect.Height > _maxHeight)
                    {
                        rect.Height = _maxHeight;
                    }

                    if (l.Position == eLegendPosition.Right ||
                        l.Position == eLegendPosition.TopRight)
                    {
                        rect.Left = sc.ChartArea.Rectangle.Width - rect.Width - TopMargin;
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

        private double GetIconLenght(SvgChart sc, double heighestText)
        {
            return sc.Chart.IsTypeLine() ? LineLength : Math.Max(MinBarLength, heighestText * 0.4);
        }

        internal void SetLegend(SvgChart sc, double entryWidth, double entryHeight)
        {
            int index = 0;
            SvgLegendSerie pSls=null;
            var pos = Chart.Legend.Position;
            foreach (var ct in sc.Chart.PlotArea.ChartTypes)
            {
                foreach (var s in ct.Series)
                {
                    var sls = new SvgLegendSerie();
                    switch (ct.ChartType)
                    {
                        case eChartType.Line:
                        case eChartType.LineMarkers:
                        case eChartType.LineMarkersStacked:
                        case eChartType.LineMarkersStacked100:
                        case eChartType.LineStacked:
                        case eChartType.LineStacked100:
                            SetLineLegend(sc, index, pSls, pos, s, sls, entryWidth, entryHeight);
                            break;
                        case eChartType.ColumnClustered:
                        case eChartType.ColumnStacked:
                        case eChartType.ColumnStacked100:
                        case eChartType.BarClustered:
                        case eChartType.BarStacked:
                        case eChartType.BarStacked100:
                            SetBarLegend(sc, index, pSls, pos, s, sls, entryWidth, entryHeight);
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
                }
            }
        }

        private void SetLineLegend(SvgChart sc, int index, SvgLegendSerie pSls, eLegendPosition pos, ExcelChartSerie s, SvgLegendSerie sls, double entryWidth, double entryHeight)
        {
            var ls = (ExcelLineChartSerie)s;
            var tm = _seriesHeadersMeasure[index];
            TextMeasurement prevTm = tm;
            if (pSls != null)
            {
                prevTm = _seriesHeadersMeasure[index - 1];
            }

            var si = GetLineSeriesIcon(sc, ls, prevTm, tm, pSls, entryWidth, entryHeight);
            sls.SeriesIcon = si;

            var tbLeft = si.X2 + MarginExtra;
            var tbTop = si.Y2 - tm.Height * 0.5;    //TODO:Should probably be font ascent 
            double tbWidth;
            if (pos == eLegendPosition.Left || pos == eLegendPosition.Right)
            {
                tbWidth = Bounds.Width - tbLeft - RightMargin;
            }
            else
            {
                tbWidth = Bounds.Width - tbLeft - RightMargin;
            }

            var tbHeight = tm.Height;
            sls.Textbox = new SvgTextBodyItem(ChartRenderer, Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);
            //sls.Textbox.Bounds.Left = si.X2 + MarginExtra;

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

        private void SetBarLegend(SvgChart sc, int index, SvgLegendSerie pSls, eLegendPosition pos, ExcelChartSerie s, SvgLegendSerie sls, double entryWidth, double entryHeight)
        {
            var bs = (ExcelBarChartSerie)s;
            var tm = _seriesHeadersMeasure[index];
            TextMeasurement prevTm = tm;
            if (pSls != null)
            {
                prevTm = _seriesHeadersMeasure[index - 1];
            }
            var si = GetBarSeriesIcon(sc, bs, prevTm, tm, pSls, entryWidth, entryHeight);
            sls.SeriesIcon = si;

            var tbLeft = si.Right + MarginExtra;
            var tbTop = si.Top - (tm.Height - si.Height) / 2; 
            double tbWidth;

            tbWidth = Bounds.Width - tbLeft - RightMargin;

            var tbHeight = tm.Height;
            sls.Textbox = new SvgTextBodyItem(ChartRenderer, Bounds, tbLeft, tbTop, tbWidth, tbHeight, false, true);
            //sls.Textbox.Bounds.Left = si.Bottom + MarginExtra;

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

        private SvgRenderLineItem GetLineSeriesIcon(SvgChart sc, ExcelChartStandardSerie cStandardSerie, TextMeasurement pTm, TextMeasurement tm, SvgLegendSerie pSls, double entryWidth, double entryHeight)
        {
            var line = new SvgRenderLineItem(sc, Rectangle.Bounds);
            line.SetDrawingPropertiesFill(cStandardSerie.Fill, sc.Chart.StyleManager.Style.SeriesLine.FillReference.Color);
            line.SetDrawingPropertiesBorder(cStandardSerie.Border, sc.Chart.StyleManager.Style.SeriesLine.BorderReference.Color, cStandardSerie.Border.Fill.Style != eFillStyle.NoFill, 0.75);
            var icon = pSls?.SeriesIcon as SvgRenderLineItem;

            GetItemPosition(sc, pTm, tm, pSls, entryWidth, entryHeight, icon?.X1 ?? 0D, icon?.Y1 ?? 0D, out double x, out double y);

            line.X1 = x;
            line.Y1 = y;
            line.X2 = x + LineLength;
            line.Y2 = y;
            line.LineCap = eLineCap.Round;

            return line;
        }

        private SvgRenderRectItem GetBarSeriesIcon(SvgChart sc, ExcelChartStandardSerie cStandardSerie, TextMeasurement pTm, TextMeasurement tm, SvgLegendSerie pSls, double entryWidth, double entryHeight)
        {
            var item = new SvgRenderRectItem(sc, Rectangle.Bounds);
            item.SetDrawingPropertiesFill(cStandardSerie.Fill, sc.Chart.StyleManager.Style.SeriesLine.FillReference.Color);
            item.SetDrawingPropertiesBorder(cStandardSerie.Border, sc.Chart.StyleManager.Style.SeriesLine.BorderReference.Color, cStandardSerie.Border.Fill.Style != eFillStyle.NoFill, 0.75);
            var iconHeight = GetIconLenght(sc, entryHeight);
            var icon = pSls?.SeriesIcon as SvgRenderRectItem;

            GetItemPosition(sc, pTm, tm, pSls, entryWidth, entryHeight, icon?.Left ?? 0D, icon?.Top ?? 0D, out double x, out double y);

            item.LineCap = eLineCap.Round;
            item.Left = x;
            item.Top = y;
            item.Width = iconHeight;
            item.Height = iconHeight;

            return item;
        }

        private double GetItemPosition(SvgChart sc, TextMeasurement pTm, TextMeasurement tm, SvgLegendSerie pSls, double entryWidth, double entryHeight, double iconLeft, double iconTop, out double x, out double y)
        {
            var topOffset = 0D;
            if (sc.Chart.Legend.Position == eLegendPosition.Top ||
               sc.Chart.Legend.Position == eLegendPosition.Bottom)
            {
                if (pSls != null && pSls.Textbox.Bounds.Right + entryWidth + RightMargin > _maxWidth)
                {
                    topOffset += entryHeight + MarginExtra;
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
                        x = iconLeft + entryWidth + MarginExtra;
                    }
                }
                if (pSls == null)
                {
                    y = Rectangle.Top + TopMargin + tm.Height / 2;
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
                    y = TopMargin;
                }
                else
                {
                    y = pSls.Textbox.Bounds.Bottom + MiddleMargin;
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