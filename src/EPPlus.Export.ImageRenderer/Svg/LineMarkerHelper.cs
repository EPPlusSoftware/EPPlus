using EPPlus.Export.ImageRenderer.Utils;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.Svg
{
    internal class LineMarkerHelper
    {
        internal static RenderItem GetMarkerItem(SvgChart sc, ExcelChartMarker m, ExcelChartStandardSerie serie, double x, double y, bool isLegend)
        {
            SvgRenderItem item;
            float maxSize = isLegend ? 7f : float.MaxValue;
            var size = m.Size > maxSize ? maxSize : m.Size;
            var halfSize = size / 2;
            var xPath = x / sc.ChartArea.Rectangle.Width;
            var yPath = y / sc.ChartArea.Rectangle.Height;
            var halfY = halfSize / sc.ChartArea.Rectangle.Height;
            var halfX = halfSize / sc.ChartArea.Rectangle.Width;
            switch (m.Style)
            {
                case eMarkerStyle.Circle:
                    item = new SvgRenderEllipseItem(sc, sc.Bounds)
                    {
                        Rx = halfSize,
                        Ry = halfSize,
                        Cx = x,
                        Cy = y
                    };
                    break;
                case eMarkerStyle.Triangle:
                    item = new SvgRenderPathItem(sc, sc.Bounds)
                    {
                        Commands = new List<PathCommands>()
                    };
                    var cmd = new PathCommands(PathCommandType.Move, item, new double[] { xPath + halfX, yPath + halfY, xPath, yPath - halfY, xPath - halfX, yPath + halfY });
                    ((SvgRenderPathItem)item).Commands.Add(cmd);
                    ((SvgRenderPathItem)item).Commands.Add(new PathCommands(PathCommandType.End, item));
                    break;
                case eMarkerStyle.Diamond:
                    item = new SvgRenderPathItem(sc, sc.Drawing.GetBoundingBox())
                    {
                        Commands = new List<PathCommands>()
                    };

                    cmd = new PathCommands(PathCommandType.Move, item, new double[] { (xPath - halfX), yPath, xPath, yPath + halfY, xPath + halfX, yPath, xPath, yPath - halfY });
                    ((SvgRenderPathItem)item).Commands.Add(cmd);
                    ((SvgRenderPathItem)item).Commands.Add(new PathCommands(PathCommandType.End, item));
                    break;
                case eMarkerStyle.Dot:
                case eMarkerStyle.Dash:
                    if (isLegend)
                    {
                        item = null;
                    }
                    else
                    {
                        if (m.Style == eMarkerStyle.Dot)
                        {
                            item = new SvgRenderRectItem(sc, sc.Bounds)
                            {
                                Left = x,
                                Top = y - size / 8,
                                Width = size / 2,
                                Height = size / 4
                            };
                        }
                        else //Dash
                        {
                            item = new SvgRenderRectItem(sc, sc.Bounds)
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
                    item = new SvgRenderRectItem(sc, sc.Bounds)
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
                    var pathItem = new SvgRenderPathItem(sc, sc.Bounds)
                    {
                        Commands = new List<PathCommands>()
                    };
                    if (m.Style == eMarkerStyle.Star)
                    {
                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath - halfX, yPath - halfY, xPath + halfX, yPath + halfY }));

                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath, yPath + halfY, xPath, yPath - halfY }));

                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath + halfX, yPath - halfY, xPath - halfX, yPath + halfY }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End, pathItem));

                    }
                    else if (m.Style == eMarkerStyle.X)
                    {
                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath - halfX, yPath - halfY, xPath + halfX, yPath + halfY }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End, pathItem));

                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath - halfX, yPath + halfY, xPath + halfX, yPath - halfY }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End, pathItem));
                    }
                    else
                    {
                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath, yPath - halfY, xPath, yPath + halfY }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End, pathItem));

                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath - halfX, yPath, xPath + halfX, yPath }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End, pathItem));
                    }
                    item = pathItem;
                    break;
                default:
                    item = null;
                    break;
            }
            if (m.Fill.IsEmpty == false)
            {
                item?.SetDrawingPropertiesFill(m.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color);
            }
            else if (serie.Fill.IsEmpty)
            {
                item?.SetDrawingPropertiesFill(serie.Border.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color);
            }
            else
            {
                item?.SetDrawingPropertiesFill(serie.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color);
            }

            if (m.Border.Width > 0)
            {
                if (m.Border.Fill.IsEmpty)
                {
                    item?.SetDrawingPropertiesBorder(serie.Border, sc.Chart.StyleManager.Style.DataPointMarker.BorderReference.Color, serie.Border.Fill.Style != eFillStyle.NoFill, 0.75);
                }
                else
                {
                    item?.SetDrawingPropertiesBorder(m.Border, sc.Chart.StyleManager.Style.DataPointMarker.BorderReference.Color, m.Border.Fill.Style != eFillStyle.NoFill, 0.75);
                }
            }
            return item;
        }

        internal static RenderItem GetMarkerBackground(SvgChart sc, ExcelChartMarker m, double x, double y, bool isLegend)
        {
            SvgRenderItem item;
            float maxSize = isLegend ? 7f : float.MaxValue;
            var size = m.Size > maxSize ? maxSize : m.Size;
            //var line = sls.SeriesIcon as SvgRenderLineItem;
            item = new SvgRenderRectItem(sc, sc.Bounds)
            {
                Left = x - (size / 2),// line.X1 + (line.X2 - line.X1 - size) / 2,
                Top = y - (size / 2),
                Width = size,
                Height = size
            };
            item?.SetDrawingPropertiesFill(m.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color);
            return item;
        }

        internal static RenderItem GetMarkerItem(SvgChart sc, ExcelLineChartSerie ls, double x, double y, bool isLegend)
        {
            SvgRenderItem item;
            var m = ls.Marker;
            float maxSize = isLegend ? 7f : float.MaxValue;
            var size = m.Size > maxSize ? maxSize : m.Size;
            var halfSize = size / 2;
            var xPath = x / sc.ChartArea.Rectangle.Width;
            var yPath = y / sc.ChartArea.Rectangle.Height;
            var halfY = halfSize / sc.ChartArea.Rectangle.Height;
            var halfX = halfSize / sc.ChartArea.Rectangle.Width;
            switch (m.Style)
            {
                case eMarkerStyle.Circle:
                    item = new SvgRenderEllipseItem(sc, sc.Bounds)
                    { 
                        Rx = halfSize,
                        Ry = halfSize,
                        Cx = x,
                        Cy = y
                    };
                    break;
                case eMarkerStyle.Triangle:
                    item = new SvgRenderPathItem(sc, sc.Bounds)
                    {
                        Commands = new List<PathCommands>()
                    };
                    var cmd = new PathCommands(PathCommandType.Move, item, new double[] { xPath + halfX, yPath + halfY, xPath, yPath - halfY, xPath - halfX, yPath + halfY });
                    ((SvgRenderPathItem)item).Commands.Add(cmd);
                    ((SvgRenderPathItem)item).Commands.Add(new PathCommands(PathCommandType.End, item));
                    break;
                case eMarkerStyle.Diamond:
                    item = new SvgRenderPathItem(sc, sc.Drawing.GetBoundingBox())
                    {
                        Commands = new List<PathCommands>()
                    };

                    cmd = new PathCommands(PathCommandType.Move, item, new double[] { (xPath - halfX), yPath, xPath, yPath + halfY, xPath + halfX, yPath, xPath, yPath - halfY });
                    ((SvgRenderPathItem)item).Commands.Add(cmd);
                    ((SvgRenderPathItem)item).Commands.Add(new PathCommands(PathCommandType.End, item));
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
                            item = new SvgRenderRectItem(sc, sc.Bounds)
                            {
                                Left = x,
                                Top = y - size / 8,
                                Width = size / 2,
                                Height = size / 4
                            };
                        }
                        else //Dash
                        {
                            item = new SvgRenderRectItem(sc, sc.Bounds)
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
                    item = new SvgRenderRectItem(sc, sc.Bounds)
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
                    var pathItem = new SvgRenderPathItem(sc, sc.Bounds)
                    {
                        Commands = new List<PathCommands>()
                    };
                    if (m.Style == eMarkerStyle.Star)
                    {
                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath - halfX, yPath - halfY, xPath + halfX, yPath + halfY }));

                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath, yPath + halfY, xPath, yPath - halfY }));

                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath + halfX, yPath - halfY, xPath - halfX, yPath + halfY }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End, pathItem));

                    }
                    else if (m.Style == eMarkerStyle.X)
                    {
                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath - halfX, yPath - halfY, xPath + halfX, yPath + halfY }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End, pathItem));

                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath - halfX, yPath + halfY, xPath + halfX, yPath - halfY }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End, pathItem));
                    }
                    else
                    {
                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath, yPath - halfY, xPath, yPath + halfY }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End, pathItem));

                        pathItem.Commands.Add(new PathCommands(PathCommandType.Move, pathItem, new double[] { xPath - halfX, yPath, xPath + halfX, yPath }));
                        pathItem.Commands.Add(new PathCommands(PathCommandType.End, pathItem));
                    }
                    item = pathItem;
                    break;
                default:
                    item = null;
                    break;
            }
            if (ls.Marker.Fill.IsEmpty == false)
            {
                item?.SetDrawingPropertiesFill(ls.Marker.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color);
            }
            else if (ls.Fill.IsEmpty)
            {
                item?.SetDrawingPropertiesFill(ls.Border.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color);
            }
            else
            {
                item?.SetDrawingPropertiesFill(ls.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color);
            }

            if (ls.Marker.Border.Width > 0)
            {
                if (ls.Marker.Border.Fill.IsEmpty)
                {
                    item?.SetDrawingPropertiesBorder(ls.Border, sc.Chart.StyleManager.Style.DataPointMarker.BorderReference.Color, ls.Border.Fill.Style != eFillStyle.NoFill, 0.75);
                }
                else
                {
                    item?.SetDrawingPropertiesBorder(ls.Marker.Border, sc.Chart.StyleManager.Style.DataPointMarker.BorderReference.Color, ls.Marker.Border.Fill.Style != eFillStyle.NoFill, 0.75);
                }
            }
            return item;
        }
        internal static RenderItem GetMarkerBackground(SvgChart sc, ExcelLineChartSerie ls,  double x, double y, bool isLegend)
        {
            SvgRenderItem item;
            var m = ls.Marker;
            float maxSize = isLegend ? 7f : float.MaxValue;
            var size = m.Size > maxSize ? maxSize : m.Size;
            //var line = sls.SeriesIcon as SvgRenderLineItem;
            item = new SvgRenderRectItem(sc, sc.Bounds)
            {
                Left = x - (size / 2),// line.X1 + (line.X2 - line.X1 - size) / 2,
                Top = y - (size / 2),
                Width = size,
                Height = size
            };
            item?.SetDrawingPropertiesFill(ls.Marker.Fill, sc.Chart.StyleManager.Style.DataPointMarker.FillReference.Color);
            return item;
        }

    }
}
