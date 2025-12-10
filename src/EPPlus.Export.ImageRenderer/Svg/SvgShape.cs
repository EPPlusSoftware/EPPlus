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
using EPPlusImageRenderer.ShapeDefinitions;
using EPPlusImageRenderer.Text;
using EPPlusImageRenderer.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;
using System.Globalization;
using System.Text;
using TypeConv = OfficeOpenXml.Utils.TypeConversion;
using EPPlus.Fonts.OpenType;
using System.Collections.Generic;
using System;
using System.Linq;
using EPPlus.Export.ImageRenderer.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgShape : DrawingShape
    {
        SvgRenderRectItem _renderTextBox;
        TextBox textBox;
        private ExcelTheme _theme;
        ExcelWorkbook _wb;
        TextContainer _textContainer;

        public SvgShape(ExcelShape shape) : base(shape)
        {
            _wb = shape._drawings.Worksheet.Workbook;
            _theme = _wb.ThemeManager.GetOrCreateTheme();
            var style = shape.Style;

            if(style==eShapeStyle.CustomShape)
            {
                _renderTextBox = null;
                foreach (var path in shape.CustomGeom.DrawingPaths)
                {
                    AddFromPaths(path);
                }
            }
            else
            {
                var shapeDef = PresetShapeDefinitions.ShapeDefinitions[style].Clone();
                shapeDef.Calculate(shape);

                //RenderItems.Add(new SvgRenderPathItem(shape));
                RenderItems.Add(new SvgGroupItem(shapeDef.GetTransform(shape.Rotation)));

                //Draw Filled path's
                foreach (var path in shapeDef.ShapePaths)
                {
                    if (path.Fill != PathFillMode.None)
                    {
                        AddFromPaths(path, true, false);
                    }
                }

                //Draw border path's
                foreach (var path in shapeDef.ShapePaths)
                {
                    if (path.Stroke)
                    {
                        AddFromPaths(path, false, true);
                    }
                }

                if (_shape.Text != null)
                {
                    if (shapeDef.TextBoxRect != null)
                    {
                        if (shape.TextBody.TextAutofit != eTextAutofit.ShapeAutofit)
                        {
                            var rectItem = new SvgRenderRectItem(_shape);

                            rectItem.X = (float)shapeDef.TextBoxRect.LeftValue;
                            rectItem.Y = (float)shapeDef.TextBoxRect.TopValue;
                            rectItem.Width = (float)shapeDef.TextBoxRect.RightValue - rectItem.X;
                            rectItem.Height = (float)shapeDef.TextBoxRect.BottomValue - rectItem.Y;
                            rectItem.FillOpacity = 0.3d;
                            _renderTextBox = rectItem;
                        }
                        else
                        {
                            var rectItem = new SvgRenderRectItem(_shape);

                            rectItem.X = (float)shapeDef.TextBoxRect.LeftValue;
                            rectItem.Y = (float)shapeDef.TextBoxRect.TopValue;
                            rectItem.Width = (float)shapeDef.TextBoxRect.RightValue;
                            rectItem.Height = (float)shapeDef.TextBoxRect.BottomValue;
                            rectItem.FillOpacity = 0.3d;
                            _renderTextBox = rectItem;
                        }
                    }
                    else
                    {
                        _renderTextBox = null;
                    }

                    textBox = GetTextBox();
                    LoadTextBox();
                }
            }
        }

        private void LoadTextBox()
        {
            textBox.VerticalAlignment = _shape.TextAnchoring;
            textBox.WrapText = _shape.TextBody.WrapText != eTextWrappingType.None;
            var fontMeasurer = (FontMeasurerTrueType)_shape._drawings._package.Settings.TextSettings.GenericTextMeasurerTrueType;

            //Make width the doc width if meant to overflow
            if (_shape.TextBody.HorizontalTextOverflow == eTextHorizontalOverflow.Overflow)
            {
                textBox.Width = Size.Width;
            }

            string color = "#" + GetFontColor();
            textBox.fontColor = color;

            //Paragraph level begins
            foreach (var paragraph in _shape.TextBody.Paragraphs)
            {
                var svgParagraph = textBox.AddParagraph(paragraph);
            }
        }

        protected void AddFromPaths(DrawingPath path, bool drawFill = true, bool drawBorder = true)
        {
            var pi = new SvgRenderPathItem(_shape);
            var coordinates = new List<double>();
            PathCommands cmd = null;
            PathsBase pCmd = null;
            double cx = 0, cy = 0;
            foreach (var p in path.Paths)
            {
                switch (p.Type)
                {
                    case PathDrawingType.MoveTo:
                        AddCmd(pi, path, coordinates, ref cmd, pCmd, p, PathCommandType.Move);
                        break;
                    case PathDrawingType.LineTo:
                        AddCmd(pi, path, coordinates, ref cmd, pCmd, p, PathCommandType.Line);
                        break;
                    case PathDrawingType.CubicBezierTo:
                        AddCmd(pi, path, coordinates, ref cmd, pCmd, p, PathCommandType.CubicBézier);
                        break;
                    case PathDrawingType.QuadBezierTo:
                        AddCmd(pi, path, coordinates, ref cmd, pCmd, p, PathCommandType.QuadraticBézier);
                        break;
                    case PathDrawingType.ArcTo:
                        SetCmdCoordinats(cmd, p, coordinates);
                        AddArc(pi, path, coordinates, pCmd, out cx, out cy, p);
                        cmd = null;
                        break;
                    case PathDrawingType.Close:
                        if (pi.Commands[pi.Commands.Count - 1].Type != PathCommandType.Arc)
                        {
                            pi.Commands[pi.Commands.Count - 1].Coordinates = coordinates.ToArray();
                            coordinates.Clear();
                        }
                        pi.Commands.Add(new PathCommands(PathCommandType.End, pi));
                        cmd = null;
                        break;
                }
                pCmd = p;
            }
            if (coordinates.Count > 0)
            {
                pi.Commands[pi.Commands.Count - 1].Coordinates = coordinates.ToArray();
            }
            if (drawFill)
            {
                pi.FillColorSource = path.Fill;                
                pi.SetDrawingPropertiesFill(_shape.Fill, _shape.ThemeStyles.FillReference.Color);
            }
            else
            {
                pi.FillColorSource = PathFillMode.None;
                pi.FillColor = "none";
            }

            if (drawBorder)
            {
                pi.BorderColorSource = path.Stroke ? PathFillMode.Norm : PathFillMode.None;
                pi.SetDrawingPropertiesBorder(_shape.Border, _shape.ThemeStyles.BorderReference.Color, path.Stroke);
            }
            else
            {
                pi.BorderColorSource = PathFillMode.None;
                pi.BorderColor = "none";
            }

            RenderItems.Add(pi);
        }
        public string ViewBox 
        { 
            get
            {
                double l=0, t=0, r=1, b=1;
                foreach(var item in RenderItems)
                {
                    item.GetBounds(out var il, out var it, out var ir, out var ib);
                    if(il<l)
                    {
                        l = il;
                    }
                    if(it<t)
                    {
                        t = it;
                    }
                    if(ir>r)
                    {
                        r = ir;
                    }
                    if(ib>b)
                    {
                        b = ib;
                    }
                }
                return $"{(l * Size.Width).ToString(CultureInfo.InvariantCulture)},{(t * Size.Height).ToString(CultureInfo.InvariantCulture)},{((Math.Abs(l) + r) * Size.Width).ToString(CultureInfo.InvariantCulture)},{((Math.Abs(t) + b) * Size.Height).ToString(CultureInfo.InvariantCulture)}";
            }
        }

        public override void Render(StringBuilder sb)
        {
            sb.Append($"<svg width=\"{Size.Width}\" height=\"{Size.Height}\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:space=\"preserve\" Overflow=\"Hidden\" viewbox=\"{ViewBox}\">");

            //Write defs used for gradient colors
            var writer = new SvgDrawingWriter(this);
            writer.WriteSvgDefs(sb, RenderItems);
            
            SvgGroupItem gItemTest = null;
            foreach(var item in RenderItems)
            {
                item.Render(sb);
                if(item.Type == SvgItemType.Group && gItemTest == null)
                {
                    gItemTest = (SvgGroupItem)item;
                }
            }
            if (!string.IsNullOrEmpty(_shape.Text))
            {
                RenderText(sb);
            }

            if (gItemTest != null)
            {
                gItemTest.RenderEndGroup(sb);
            }
            sb.AppendLine("</svg>");
        }

        private string GetFontColor()
        {
            string color;

            if (_shape.Font.Fill.Style == eFillStyle.SolidFill)
            {
                var c = TypeConv.ColorConverter.GetThemeColor(_shape.Font.Fill.SolidFill.Color);
                color = ((uint)c.ToArgb()).ToString("x").Substring(2, 6);
            }
            else
            {
                var c = TypeConv.ColorConverter.GetThemeColor(_wb.ThemeManager.CurrentTheme, _shape.ThemeStyles.FontReference.Color);
                color = ((uint)c.ToArgb()).ToString("x").Substring(2, 6);
            }

            return color;
        }

        TextBox GetTextBox()
        {
            if(_renderTextBox != null)
            {
                return new TextBox(_shape.TextBody, _renderTextBox);
            }
            else
            {
                GetShapeInnerBound(out double x, out double y, out double width, out double height);
                return new TextBox(x, y, width, height);
            }
        }

        private void GetFontNameAndSize(ExcelFont nsFont, out string fontName, out double fontSize)
        {
            fontName = string.IsNullOrEmpty(_shape.Font.LatinFont) ? _shape.Font.ComplexFont : _shape.Font.LatinFont;

            fontSize = _shape.Font.Size;
            if (string.IsNullOrEmpty(fontName)) fontName = nsFont?.Name ?? _theme.FontScheme.MajorFont.First().Typeface;
            if (fontSize <= 0 && nsFont != null) fontSize = nsFont.Size;
        }

        private void RenderText(StringBuilder sb)
        {
            RenderDebugTextBox(sb);
            textBox.RenderParagraphs(sb);
        }

        private void RenderDebugTextBox(StringBuilder sb)
        {
            _renderTextBox.FillColor = "green";
            _renderTextBox.Render(sb);

            var area = textBox.GetTextArea();

            _renderTextBox.X = (float)area.Left;
            _renderTextBox.Y = (float)area.Top;
            _renderTextBox.Width = (float)area.Width;
            _renderTextBox.Height = (float)area.Height;
            _renderTextBox.FillColor = "blue";
            _renderTextBox.Render(sb);
        }

        private void GetShapeInnerBound(out double x, out double y, out double width, out double height)
        {
            double currentX = 0, currentY = 0, xe, ye;
            x = y = 0;
            width = xe = Size.Width;
            height = ye = Size.Height;
            foreach (var ri in RenderItems)
            {
                switch (ri.Type)
                {
                    case SvgItemType.Rect:
                        var rectItem = (SvgRenderRectItem)ri;
                        x = rectItem.X;
                        y = rectItem.Y;
                        width = rectItem.Width;
                        height = rectItem.Height;
                        break;
                    case SvgItemType.Path:
                        var pathItem = (SvgRenderPathItem)ri;
                        foreach (var cmd in pathItem.Commands)
                        {
                            var cmdCoordinates = new List<Coordinate>();
                            for (int i = 0; i < cmd.Coordinates.Length; i++)
                            {
                                switch (cmd.Type)
                                {
                                    case PathCommandType.Move:
                                        if (i == 0)
                                        {
                                            currentX = cmd.Coordinates[i];
                                            currentY = cmd.Coordinates[++i];
                                        }
                                        else
                                        {
                                            HandleLine(ref currentX, ref currentY, ref xe, ref ye, cmd, cmdCoordinates, ref i);
                                        }
                                        break;
                                    case PathCommandType.VerticalLine:
                                        HandleVertical(y, ref currentY, ref xe, cmd.Coordinates[i]);
                                        break;
                                    case PathCommandType.HorizontalLine:
                                        HandleHorizontal(x, ref currentX, ref xe, cmd.Coordinates[i]);
                                        break;
                                    case PathCommandType.Line:
                                        HandleLine(ref currentX, ref currentY, ref xe, ref ye, cmd, cmdCoordinates, ref i);
                                        break;
                                    case PathCommandType.CubicBézier:
                                        if (currentX > x)
                                        {
                                            x = currentX;
                                        }
                                        if (currentY > y)
                                        {
                                            y = currentY;
                                        }
                                        i += 4;
                                        break;

                                }
                            }
                        }
                        break;
                }
            }
            if (xe != double.MinValue)
            {
                width = xe - x;
            }
            if (ye != double.MinValue)
            {
                height = ye - y;
            }
        }

        private static void HandleLine(ref double currentX, ref double currentY, ref double xe, ref double ye, PathCommands cmd, List<Coordinate> cmdCoordinates, ref int i)
        {
            xe = cmd.Coordinates[i];
            ye = cmd.Coordinates[++i];
            if (xe == currentX || ye == currentY)
            {
                if (cmdCoordinates.Count == 0)
                {
                    cmdCoordinates.Add(new Coordinate(currentX, currentY));
                }
                cmdCoordinates.Add(new Coordinate(xe, ye));
            }
            else
            {
                var w = Math.Abs(xe - currentX);
                var h = Math.Abs(ye - currentY);
                cmdCoordinates.Add(new Coordinate((Math.Min(xe, currentX) + w) / 2, (Math.Min(ye, currentY) + h) / 2));
            }
            currentX = xe;
            currentY = ye;
        }


        private static void HandleVertical(double y, ref double currentY, ref double ye, double yec)
        {
            if (currentY < y || currentY == double.MinValue)
            {
                currentY = y;
            }
            if (ye > yec || ye == double.MinValue)
            {
                ye = yec;
            }
        }
        private static void HandleHorizontal(double x, ref double currentX, ref double xe, double xec)
        {
            if (currentX < x || currentX == double.MinValue)
            {
                currentX = x;
            }
            if (xe > xec || xe == double.MinValue)
            {
                xe = xec;
            }
        }
    }
}