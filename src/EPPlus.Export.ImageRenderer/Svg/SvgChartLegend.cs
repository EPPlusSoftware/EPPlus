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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Collections.Generic;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartLegend : SvgChartObject
    {
        internal SvgChartLegend(SvgChart sc) : base(sc.Chart)
        {
            if (sc.Chart.HasLegend == false || sc.Chart.Series.Count == 0)
            {
                return;
            }
            float textHeight, textWidth;
            var l = ((ExcelChartStandard)sc.Chart).Legend;


            SetMargins(l.TextBody);
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
        }

        private SvgRenderRectItem GetLegendRectangle(SvgChart sc, ExcelChartLegend l)
        {
            var rect = new SvgRenderRectItem(Chart);
            bool isVertical;
            switch(l.Position)
            {
                case eLegendPosition.Top:
                case eLegendPosition.Bottom:
                    isVertical = false;
                    break;
                default:
                    isVertical = true; 
                    break;
            }

            if (isVertical)
            {
                var width = 30D;
                var tm = TextMeasurer.MeasureText("Aq", l.Font.GetMeasureFont()); 
                var height = TopMargin + tm.Height + BottomMargin;

                //foreach (var ct in sc.Chart.PlotArea.ChartTypes)
                //{
                //    foreach(var s in ct.Series)
                //    {
                //        foreach(var t in s)
                //        { 
                            
                //        }
                //    }
                //}
            }
            return rect;
        }

        internal void SetLegened(SvgChart sc)
        {
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
                            if(ls.HasMarker() && ls.Marker.Style != eMarkerStyle.None)
                            {
                                sls.SeriesIcon = GetSeriesIcon(sc, ls);
                                sls.MarkerIcon = GetMarkerItem(sc, ls, sls);
                            }
                            break;
                        default:
                            break;
                    } 
                }
            }
        }

        private RenderItem GetSeriesIcon(SvgChart sc, ExcelLineChartSerie ls)
        {
            var item = new SvgRenderLineItem(sc.Chart);
            item.SetDrawingPropertiesFill(ls.Fill, sc.Chart.StyleManager.Style.SeriesLine.FillReference.Color);
            item.SetDrawingPropertiesBorder(ls.Border, sc.Chart.StyleManager.Style.SeriesLine.BorderReference.Color, ls.Border.Fill.Style!=eFillStyle.NoFill, 0.75);
            return item;
        }


        private RenderItem GetMarkerItem(SvgChart sc, ExcelLineChartSerie ls, SvgLegendSerie sls)
        {
            SvgRenderItem item;
            var m = ls.Marker;
            switch (m.Style)
            {
                case eMarkerStyle.Circle:
                    item = new SvgRenderEllipseItem(sc.Drawing)
                    {
                        Rx = m.Size,
                        Ry = m.Size
                    };
                    break;
                case eMarkerStyle.Triangle:
                    item = new SvgRenderPathItem(sc.Drawing)
                    {
                        Commands = new List<PathCommands>()
                    };
                    var cmd = new PathCommands(PathCommandType.Move, item, new double[] { 0, m.Size, m.Size / 2, 0, m.Size, m.Size });
                    ((SvgRenderPathItem)item).Commands.Add(cmd);
                    break;
                case eMarkerStyle.Diamond:
                    item = new SvgRenderPathItem(sc.Drawing)
                    {
                        Commands = new List<PathCommands>()
                    };
                    var hs = m.Size / 2;
                    cmd = new PathCommands(PathCommandType.Move, item, new double[] { hs, hs, hs, 0, m.Size, hs, hs, m.Size });
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
                        Width = m.Size,
                        Height = m.Size
                    };
                    break;
                default:
                    item = null;
                    break;
            }
            item?.SetDrawingPropertiesFill(ls.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color);
            item?.SetDrawingPropertiesBorder(ls.Border, sc.Chart.StyleManager.Style.DataPointMarker.BorderReference.Color, ls.Border.Fill.Style == eFillStyle.NoFill, 0.75);
            return item;
        }

        public override void Render(StringBuilder sb)
        {
            Rectangle.Render(sb);
            foreach(var s in SeriesIcon)
            {
                s.SeriesIcon.Render(sb);
                s.MarkerIcon.Render(sb);
            }
        }

        public List<SvgLegendSerie> SeriesIcon { get; } = new List<SvgLegendSerie>();
    }
    internal class SvgLegendSerie
    {
        internal RenderItem SeriesIcon { get; set; }
        internal RenderItem MarkerIcon { get; set; }
    }
}