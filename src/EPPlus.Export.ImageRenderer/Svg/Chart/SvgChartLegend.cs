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
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
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
        internal SvgChartLegend(SvgChart sc, bool isDataLabelLegend = false) : base(sc)
        {
            _ttMeasurer = sc.Chart.WorkSheet._package.Settings.TextSettings.GenericTextMeasurerTrueType;
            if (sc.Chart.HasLegend == false && isDataLabelLegend == false || sc.Chart.Series.Count == 0)
            {
                return;
            }
            var l = ((ExcelChartStandard)sc.Chart).Legend;

            LeftMargin = RightMargin = 3; //4px
            TopMargin = BottomMargin = 3; //4px

            if (l.Layout.HasLayout)
            {
                Rectangle = GetRectFromManualLayout(sc, l.Layout);
            }
            else
            {
                Rectangle = GetLegendRectangle(sc, l);
            }

            Bounds.Left = Rectangle.Left;
            Bounds.Top = Rectangle.Top;
            Bounds.Width = Rectangle.Width;
            Bounds.Height = Rectangle.Height;
            Rectangle.Bounds.Left = Rectangle.Bounds.Top = 0;

            Rectangle.SetDrawingPropertiesFill(l.Fill, sc.Chart.StyleManager.Style.Title.FillReference.Color);
            Rectangle.SetDrawingPropertiesBorder(l.Border, sc.Chart.StyleManager.Style.Title.BorderReference.Color, l.Border.Fill.Style != eFillStyle.NoFill, 0.75);

            SetLegend(sc);
        }

        private SvgRenderRectItem GetLegendRectangle(SvgChart sc, ExcelChartLegend l)
        {
            var rect = new SvgRenderRectItem(sc, sc.Bounds);
            bool isVertical;
            switch (l.Position)
            {
                case eLegendPosition.Top:
                case eLegendPosition.Bottom:
                    isVertical = false;
                    break;
                default:
                    isVertical = true; 
                    break;
            }
            
            var widest = 0d;
            var highest = 0d;
            var textWidth = 0d;
            var height = TopMargin;
            var index = 0;
            foreach (var ct in sc.Chart.PlotArea.ChartTypes)
            {
                foreach (var s in ct.Series)
                {
                    var text = s.GetHeaderText(index);
                    var entry = l.Entries.FirstOrDefault(x => x.Index == index);
                    ExcelTextFont font;
                    if(entry==null || entry.Font.IsEmpty)
                    {
                        font = l.Font;
                    }
                    else
                    {
                        font = entry.Font;
                    }
                    var tm = _ttMeasurer.MeasureText(text, font.GetMeasureFont());
                    _seriesHeadersMeasure.Add(tm);
                    if(tm.Width > widest)
                    {
                        widest = tm.Width;
                    }
                    if (tm.Height > height)
                    {
                        highest = tm.Height;
                    }
                    textWidth += tm.Width;
                    height += tm.Height + MiddleMargin;
                    index++;
                }
            }
            height = height - MiddleMargin + BottomMargin; //remove last margin and add bottom margin
            switch (l.Position)
            {
                case eLegendPosition.Top:
                case eLegendPosition.Bottom:
                    rect.Width = textWidth + LeftMargin + RightMargin + ((LineLength + MarginExtra) * index + (MiddleMargin*Math.Max(index-1,0))) + 2 ; // 28 is for the line length + 2px between line and text
                    rect.Height = TopMargin + BottomMargin + highest + MarginExtra;
                    rect.Left = (sc.ChartArea.Rectangle.Width - rect.Width) / 2;
                    if (l.Position == eLegendPosition.Top)
                    {                        
                        rect.Top = sc.Title.Rectangle.Top + sc.Title.Rectangle.Height + MiddleMargin;
                    }
                    else 
                    {
                        rect.Top = sc.ChartArea.Rectangle.Height - rect.Height - BottomMargin;
                    }
                    break;
                case eLegendPosition.Right:
                case eLegendPosition.TopRight:
                case eLegendPosition.Left:
                    rect.Width = widest + LeftMargin + RightMargin + LineLength + 2; // 28 is for the line length + 2px between line and text
                    rect.Height = height + BottomMargin;
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
                        rect.Top = sc.ChartArea.Rectangle.Height / 2 + TopMargin + 2;
                    }
                    else
                    {
                        if (sc.Title == null)
                        {
                            rect.Top = 8 + 8;
                        }
                        else
                        {
                            rect.Top = sc.Title.Rectangle.Height + 8 + 8; //Height+Margin Top and Bottom Title
                        }
                    }
                    break;
            }
            if (isVertical)
            {                

                //var top = sc.Title.GetRectangle.Height+8+10;
                //var width = margin;
            }
            return rect;
        }

        internal void SetLegend(SvgChart sc)
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
                            var ls=(ExcelLineChartSerie)s;
                            var tm = _seriesHeadersMeasure[index];
                            var prevTm = _seriesHeadersMeasure[index - 1];
                            var si = GetSeriesIcon(sc, ls, prevTm, tm, pSls);
                            sls.SeriesIcon = si;

                            var tbLeft = si.X2 + MarginExtra;
                            var tbTop = si.Y2 - tm.Height * 0.5; //TODO:Should probably be font ascent 
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
                            sls.Textbox.Bounds.Left = si.X2 + MarginExtra;

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
                                var x= l.X1 + (l.X2 - l.X1) / 2;
                                var y = l.Y1;
                                sls.MarkerIcon = LineMarkerHelper.GetMarkerItem(sc, ls, x, y, true);
                                if((ls.Marker.Style == eMarkerStyle.Plus || ls.Marker.Style == eMarkerStyle.X || ls.Marker.Style == eMarkerStyle.Star) &&
                                    ls.Marker.Fill.IsEmpty == false)
                                {
                                    sls.MarkerBackground = LineMarkerHelper.GetMarkerBackground(sc, ls, x, y, true);
                                }
                                else
                                {
                                    sls.MarkerBackground = null;
                                }
                            }
                            break;
                        default:
                            break;
                    }
                    SeriesIcon.Add(sls);
                    pSls = sls;
                    index++;
                }
            }
        }

        private SvgRenderLineItem GetSeriesIcon(SvgChart sc, ExcelChartStandardSerie cStandardSerie, TextMeasurement pTm, TextMeasurement tm, SvgLegendSerie pSls)
        {
            var item = new SvgRenderLineItem(sc, Rectangle.Bounds);
            item.SetDrawingPropertiesFill(cStandardSerie.Fill, sc.Chart.StyleManager.Style.SeriesLine.FillReference.Color);
            item.SetDrawingPropertiesBorder(cStandardSerie.Border, sc.Chart.StyleManager.Style.SeriesLine.BorderReference.Color, cStandardSerie.Border.Fill.Style!=eFillStyle.NoFill, 0.75);

            if (sc.Chart.Legend.Position == eLegendPosition.Top ||
               sc.Chart.Legend.Position == eLegendPosition.Bottom)
            {
                float y = (float)Rectangle.Top + (float)TopMargin + tm.Height / 2 + MarginExtra;
                float x = 0;                
                if (pSls == null)
                {
                    x = (float)Rectangle.Left + (float)LeftMargin;// + MarginExtra;
                }
                else
                {
                    x = (float)pSls.Textbox.Bounds.Right + MiddleMargin;
                }

                item.X1 = x;
                item.Y1 = y;
                item.X2 = x + LineLength;
                item.Y2 = y;
                item.LineCap = eLineCap.Round;
            }
            else
            {
                double y;
                if (pSls == null)
                {
                    y = TopMargin + tm.Height / 2 + MarginExtra;
                }
                else
                {
                    y = ((SvgRenderLineItem)pSls.SeriesIcon).Y1 + pTm.Height / 2 + tm.Height / 2 + MiddleMargin;
                }

                item.X1 = (float)LeftMargin; //4
                item.Y1 = y;
                item.X2 = (float)LineLength;
                item.Y2 = y;
                item.LineCap = eLineCap.Round;
            }

            return item;
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var groupItem = new SvgGroupItem(ChartRenderer, Bounds);
            renderItems.Add(groupItem);
            renderItems.Add(Rectangle);
            foreach(var s in SeriesIcon)
            {
                renderItems.Add(s.SeriesIcon);
                if(s.MarkerBackground != null) renderItems.Add(s.MarkerBackground);
                if (s.MarkerIcon != null) renderItems.Add(s.MarkerIcon);
                //renderItems.Add(s.Textbox);
                s.Textbox.AppendRenderItems(renderItems);
            }
            renderItems.Add(new SvgEndGroupItem(ChartRenderer, null));
        }

        public List<SvgLegendSerie> SeriesIcon { get; } = new List<SvgLegendSerie>();

    }
}