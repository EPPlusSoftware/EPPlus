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
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Text;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartLegend : SvgChartObject
    {
        
        List<TextMeasurement> _seriesHeadersMeasure =new List<TextMeasurement>();
        ITextMeasurer _ttMeasurer;
        int _middleMargin = 10;
        internal SvgChartLegend(SvgChart sc) : base(sc.Chart)
        {
            _ttMeasurer = sc.Chart.WorkSheet._package.Settings.TextSettings.GenericTextMeasurerTrueType;
            if (sc.Chart.HasLegend == false || sc.Chart.Series.Count == 0)
            {
                return;
            }
            var l = ((ExcelChartStandard)sc.Chart).Legend;

            LeftMargin = RightMargin = 4;
            TopMargin = BottomMargin = 4;

            if (l.Layout.HasLayout)
            {
                Rectangle = GetRectFromManualLayout(sc, l.Layout);
            }
            else
            {
                Rectangle = GetLegendRectangle(sc, l);
            }

            Rectangle.SetDrawingPropertiesFill(l.Fill, sc.Chart.StyleManager.Style.Title.FillReference.Color);
            Rectangle.SetDrawingPropertiesBorder(l.Border, sc.Chart.StyleManager.Style.Title.BorderReference.Color, l.Border.Fill.Style != eFillStyle.NoFill, 0.75);

            SetLegend(sc);
        }

        private SvgRenderRectItem GetLegendRectangle(SvgChart sc, ExcelChartLegend l)
        {
            var rect = new SvgRenderRectItem(Chart);
            bool isVertical;
            const int LineLength = 28;
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
            var height = TopMargin;
            var index = 0;
            foreach (var ct in sc.Chart.PlotArea.ChartTypes)
            {
                foreach (var s in ct.Series)
                {
                    var text = s.GetHeaderText();
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
                    height += tm.Height + _middleMargin;
                    index++;
                }
            }
            height = height - _middleMargin + BottomMargin; //remove last margin and add bottom margin
            switch (l.Position)
            {
                case eLegendPosition.Top:
                    break;
                case eLegendPosition.Bottom:
                    break;
                case eLegendPosition.Right:
                case eLegendPosition.TopRight:
                case eLegendPosition.Left:
                    rect.Width = widest + LeftMargin + RightMargin + LineLength + 2; // 28 is for the line length + 2px between line and text
                    rect.Height = height + BottomMargin;
                    if (l.Position == eLegendPosition.Right ||
                        l.Position == eLegendPosition.TopRight)
                    {
                        rect.X = sc.ChartArea.Width - rect.Width - TopMargin;
                    }
                    else
                    {
                        rect.X = LeftMargin + 2;
                    }
                    if (l.Position == eLegendPosition.Left ||
                        l.Position == eLegendPosition.Right)
                    {
                        rect.Y = sc.ChartArea.Height / 2 + TopMargin + 2;
                    }
                    else
                    {
                        if (sc.Title == null)
                        {
                            rect.Y = 8 + 8;
                        }
                        else
                        {
                            rect.Y = sc.Title.Rectangle.Height + 8 + 8; //Height+Margin Top and Bottom Title
                        }
                    }
                    break;
            }
            if (isVertical)
            {                

                //var top = sc.Title.Rectangle.Height+8+10;
                //var width = margin;
            }
            return rect;
        }

        internal void SetLegend(SvgChart sc)
        {
            int index = 0;
            SvgLegendSerie pSls=null;
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
                            var si = GetSeriesIcon(sc, ls, index, tm, pSls);
                            sls.SeriesIcon = si;
                            sls.Textbox = new TextBox(sc.Chart, si.X2+4, (si.Y1 - (tm.Height * 0.75)), tm.Width, tm.Height);
                            var entry = Chart.Legend.Entries.FirstOrDefault(x => x.Index == index);
                            if (entry == null || entry.Font.IsEmpty)
                            {
                                sls.Textbox.AddText(s.GetHeaderText(), sc.Chart.Legend.Font);
                            }
                            else
                            {
                                sls.Textbox.AddText(s.GetHeaderText(), entry.Font);
                            }
                            if (ls.HasMarker() && ls.Marker.Style != eMarkerStyle.None)
                            {
                                sls.MarkerIcon = GetMarkerItem(sc, ls, sls, index);
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

        private SvgRenderLineItem GetSeriesIcon(SvgChart sc, ExcelLineChartSerie ls, int index, TextMeasurement tm, SvgLegendSerie pSls)
        {
            var item = new SvgRenderLineItem(sc.Chart);
            item.SetDrawingPropertiesFill(ls.Fill, sc.Chart.StyleManager.Style.SeriesLine.FillReference.Color);
            item.SetDrawingPropertiesBorder(ls.Border, sc.Chart.StyleManager.Style.SeriesLine.BorderReference.Color, ls.Border.Fill.Style!=eFillStyle.NoFill, 0.75);
            float y;
            if(pSls==null)
            {
                y = (float)Rectangle.Y + (float)TopMargin + tm.Height / 2 + 2;
            }
            else
            {
                var pTm = _seriesHeadersMeasure[index - 1];
                y = ((SvgRenderLineItem)pSls.SeriesIcon).Y1 + pTm.Height / 2 + tm.Height / 2 + _middleMargin;
            }
            item.X1 = (float)Rectangle.X + 4;
            item.Y1 = y;
            item.X2 = (float)Rectangle.X + 32;
            item.Y2 = y;
            item.LineCap = eLineCap.Round;
            return item;
        }


        private RenderItem GetMarkerItem(SvgChart sc, ExcelLineChartSerie ls, SvgLegendSerie sls, int index)
        {
            SvgRenderItem item;
            var m = ls.Marker;
            var size = m.Size > 7 ? 7:m.Size;
            var line = sls.SeriesIcon as SvgRenderLineItem;
            switch (m.Style)
            {
                case eMarkerStyle.Circle:
                    item = new SvgRenderEllipseItem(sc.Drawing)
                    {
                        Rx = size,
                        Ry = size,
                        Cx = line.X1 + (line.X2 - line.X1) / 2,
                        Cy = line.Y1
                    };
                    break;
                case eMarkerStyle.Triangle:
                    item = new SvgRenderPathItem(sc.Drawing)
                    {
                        Commands = new List<PathCommands>()
                    };
                    var cmd = new PathCommands(PathCommandType.Move, item, new double[] { 0, size, size / 2, 0, size, size });
                    ((SvgRenderPathItem)item).Commands.Add(cmd);
                    break;
                case eMarkerStyle.Diamond:
                    item = new SvgRenderPathItem(sc.Drawing)
                    {
                        Commands = new List<PathCommands>()
                    };
                    var hs = 7 / 2;
                    cmd = new PathCommands(PathCommandType.Move, item, new double[] { hs, hs, hs, 0, size, hs, hs, size });
                    ((SvgRenderPathItem)item).Commands.Add(cmd);
                    break;
                case eMarkerStyle.Dot:
                    item = new SvgRenderLineItem(sc.Drawing);
                    break;
                case eMarkerStyle.Dash:
                    item = new SvgRenderLineItem(sc.Drawing);
                    break;
                case eMarkerStyle.Plus:
                case eMarkerStyle.Star:
                case eMarkerStyle.X:
                case eMarkerStyle.Square:
                    item = new SvgRenderRectItem(sc.Drawing)
                    {
                        X = line.X1 + (line.X2 - line.X1 - size) / 2,
                        Y = line.Y1 - ((size+1) / 2),
                        Width = size,
                        Height = size
                    };
                    break;
                default:
                    item = null;
                    break;
            }
            if (ls.Fill.IsEmpty)
            {
                item?.SetDrawingPropertiesFill(ls.Border.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color);
            }
            else
            {
                item?.SetDrawingPropertiesFill(ls.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color);
            }
            item?.SetDrawingPropertiesBorder(ls.Border, sc.Chart.StyleManager.Style.DataPointMarker.BorderReference.Color, ls.Border.Fill.Style == eFillStyle.NoFill, 0.75);
            return item;
        }

        public override void Render(StringBuilder sb)
        {
            Rectangle.Render(sb);
            foreach(var s in SeriesIcon)
            {
                s.SeriesIcon.Render(sb);
                s.MarkerIcon?.Render(sb);
                s.Textbox.Render(sb);
            }
        }

        public List<SvgLegendSerie> SeriesIcon { get; } = new List<SvgLegendSerie>();
    }
    internal class SvgLegendSerie
    {
        internal RenderItem SeriesIcon { get; set; }
        internal RenderItem MarkerIcon { get; set; }
        internal TextBox Textbox { get; set;}
    }
}