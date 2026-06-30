using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace OfficeOpenXml.Drawing.Renderer.Chart
{
    internal class LineMarkerHelper
    {

        internal static RenderItem GetMarkerItem(ChartRenderer sc, ExcelLineChartSerie ls, double x, double y, bool isLegend)
        {
            RenderItem item;
            var m = ls.Marker;
            float maxSize = isLegend ? 7f : float.MaxValue;
            var size = m.Size > maxSize ? maxSize : m.Size;
            var halfSize = size / 2;
            var xPath = x;
            var yPath = y;

            //var halfY = halfSize / sc.ChartArea.Rectangle.Height;
            //var halfX = halfSize / sc.ChartArea.Rectangle.Width;
            switch (m.Style)
            {
                case eMarkerStyle.Circle:
                    item = new EllipseRenderItem(sc.Bounds)
                    { 
                        Rx = halfSize,
                        Ry = halfSize,
                        Cx = x,
                        Cy = y
                    };
                    break;
                case eMarkerStyle.Triangle:
                    item = new PathRenderItem(sc.Bounds);
                    var cmd = new PathCommands(PathCommandType.Move, new double[] { xPath + halfSize, yPath + halfSize, xPath, yPath - halfSize, xPath - halfSize, yPath + halfSize });
                    ((PathRenderItem)item).Commands.Add(cmd);
                    ((PathRenderItem)item).Commands.Add(new PathCommands(PathCommandType.End));
                    break;
                case eMarkerStyle.Diamond:
                    item = new PathRenderItem(sc.ChartArea.Rectangle.Bounds);
                    cmd = new PathCommands(PathCommandType.Move, new double[] { (xPath - halfSize), yPath, xPath, yPath + halfSize, xPath + halfSize, yPath, xPath, yPath - halfSize });
                    ((PathRenderItem)item).Commands.Add(cmd);
                    ((PathRenderItem)item).Commands.Add(new PathCommands(PathCommandType.End));
                    break;
                case eMarkerStyle.Dot:
                case eMarkerStyle.Dash:
                    if (isLegend)
                    {
                        item = null;
                    }
                    else
                    {
                        if(m.Style == eMarkerStyle.Dot)
                        {
                            item = new RectRenderItem(sc.Bounds)
                            {
                                Left = x,
                                Top = y - size / 8,
                                Width = size / 2,
                                Height = size / 4
                            };
                        }
                        else //Dash
                        {
                            item = new RectRenderItem(sc.Bounds)
                            {
                                Left = x - size / 2,
                                Top = y - size / 8,
                                Width = size,
                                Height = size / 4
                            };
                        }
                    }
                    break;
                case eMarkerStyle.Square:
                    item = new RectRenderItem(sc.Bounds)
                    {
                        Left = x - size / 2,
                        Top = y - size / 2,
                        Width = size,
                        Height = size
                    };
                    break;
                case eMarkerStyle.Plus:
                case eMarkerStyle.Star:
                case eMarkerStyle.X:
                    var pathItem = new PathRenderItem(sc.Bounds);
                    if (m.Style == eMarkerStyle.Star)
                    {
                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, new double[] { xPath - halfSize, yPath - halfSize, xPath + halfSize, yPath + halfSize }));

                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, new double[] { xPath, yPath + halfSize, xPath, yPath - halfSize }));

                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, new double[] { xPath + halfSize, yPath - halfSize, xPath - halfSize, yPath + halfSize }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End));

                    }
                    else if (m.Style == eMarkerStyle.X)
                    {
                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, new double[] { xPath - halfSize, yPath - halfSize, xPath + halfSize, yPath + halfSize }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End));

                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, new double[] { xPath - halfSize, yPath + halfSize, xPath + halfSize, yPath - halfSize }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End));
                    }
                    else
                    {
                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, new double[] { xPath, yPath - halfSize, xPath, yPath + halfSize }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End));

                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, new double[] { xPath - halfSize, yPath, xPath + halfSize, yPath }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End));
                    }
                    item = pathItem;
                    break;
                default:
                    item = null;
                    break;
            }
            if (ls.Marker.Fill.IsEmpty == false)
            {
                item?.SetDrawingPropertiesFill(sc.Theme, ls.Marker.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color, false);
            }
            else if (ls.Fill.IsEmpty)
            {
                item?.SetDrawingPropertiesFillBasic(sc.Theme, ls.Border.Fill, sc.Chart.StyleManager.Style?.DataPointMarker.FillReference.Color, false, Color.Empty);
            }
            else
            {
                item?.SetDrawingPropertiesFill(sc.Theme, ls.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color, false);
            }

            if (ls.Marker.Border.Width > 0)
            {
                if (ls.Marker.Border.Fill.IsEmpty)
                {
                    item?.SetDrawingPropertiesBorder(sc.Theme, ls.Border, sc.Chart.StyleManager.Style.DataPointMarker.BorderReference.Color, ls.Border.Fill.Style != eFillStyle.NoFill, sc.Theme.FormatScheme.BorderStyle[0].Fill.Color, 0.75d);
                }
                else
                {
                    item?.SetDrawingPropertiesBorder(sc.Theme, ls.Marker.Border, sc.Chart.StyleManager.Style.DataPointMarker.BorderReference.Color, ls.Marker.Border.Fill.Style != eFillStyle.NoFill, sc.Theme.FormatScheme.BorderStyle[0].Fill.Color, 0.75d);
                }
            }
            return item;
        }
        internal static RenderItem GetMarkerBackground(ChartRenderer sc, ExcelLineChartSerie ls,  double x, double y, bool isLegend)
        {
            RenderItem item;
            var m = ls.Marker;
            float maxSize = isLegend ? 7f : float.MaxValue;
            var size = m.Size > maxSize ? maxSize : m.Size;
            item = new RectRenderItem(sc.Bounds)
            {
                Left = x - (size / 2),
                Top = y - (size / 2),
                Width = size,
                Height = size
            };
            item?.SetDrawingPropertiesFill(sc.Theme, ls.Marker.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color, false);
            return item;
        }

    }
}
