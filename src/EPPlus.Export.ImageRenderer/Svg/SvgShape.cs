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
using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.ShapeDefinitions;
using EPPlusImageRenderer.Text;
using EPPlusImageRenderer.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using TypeConv = OfficeOpenXml.Utils.TypeConversion;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgShape : DrawingShape
    {
        /// <summary>
        /// Calculated shape textbox
        /// </summary>
        SvgRenderRectItem InsetTextBox;
        /// <summary>
        /// Textbox from memory
        /// </summary>
        public SvgTextBoxItem TextBox { get; internal set; }

        public SvgShape(ExcelShape shape) : base(shape)
        {
            var style = shape.Style;

            if (style==eShapeStyle.CustomShape)
            {
                foreach (var path in shape.CustomGeom.DrawingPaths)
                {
                    AddFromPaths(path);
                }
            }
            else
            {
                var shapeDef = PresetShapeDefinitions.ShapeDefinitions[style].Clone();
                shapeDef.Calculate(shape);

                RenderItems.Add(new SvgGroupItem(this, shape.GetBoundingBox(), shape.Rotation));

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

                RenderItems.Add(new SvgEndGroupItem(this, Bounds));

                if (_shape.Text != null)
                {
                    if (shapeDef.TextBoxRect != null)
                    {
                        InsetTextBox = new SvgRenderRectItem(this, Bounds);
                        InsetTextBox.Bounds.Left = (float)shapeDef.TextBoxRect.LeftValue;
                        InsetTextBox.Bounds.Top = (float)shapeDef.TextBoxRect.TopValue;
                        InsetTextBox.FillOpacity = 0.3d;

                        if (shape.TextBody.TextAutofit != eTextAutofit.ShapeAutofit)
                        {
                            InsetTextBox.Width = (float)shapeDef.TextBoxRect.RightValue - (float)shapeDef.TextBoxRect.LeftValue;
                            InsetTextBox.Height = (float)shapeDef.TextBoxRect.BottomValue - (float)shapeDef.TextBoxRect.TopValue;
                        }
                        else
                        {
                            InsetTextBox.Width = (float)shapeDef.TextBoxRect.RightValue;
                            InsetTextBox.Height = (float)shapeDef.TextBoxRect.BottomValue;
                        }
                    }
                    else
                    {
                        InsetTextBox = null;
                    }

                    TextBox = CreateTextBodyItem();
                    TextBox.ImportTextBody(_shape.TextBody);
                    TextBox.AppendRenderItems(RenderItems);
                }
            }
        }

        protected void AddFromPaths(DrawingPath path, bool drawFill = true, bool drawBorder = true)
        {
            var pi = new SvgRenderPathItem(this, _shape.GetBoundingBox());
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
                return $"{(Bounds.Left).ToString(CultureInfo.InvariantCulture)},{Bounds.Top.ToString(CultureInfo.InvariantCulture)},{Bounds.Right.ToString(CultureInfo.InvariantCulture)},{Bounds.Bottom.ToString(CultureInfo.InvariantCulture)}";
            }
        }

        public void Render(StringBuilder sb)
        {
            sb.Append($"<svg width=\"{Bounds.Width}\" height=\"{Bounds.Height}\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:space=\"default\" Overflow=\"Hidden\" viewbox=\"{ViewBox}\">");

            //Write defs used for gradient colors
            var writer = new SvgDrawingWriter(this);
            writer.WriteSvgDefs(sb, RenderItems);
            
            //SvgGroupItem gItemTest = null;
            foreach(var item in RenderItems)
            {
                item.Render(sb);
                //if(item.Type == RenderItemType.Group && gItemTest == null)
                //{
                //    gItemTest = (SvgGroupItem)item;
                //}
                //if (item.IsEndOfGroup && gItemTest != null)
                //{
                //    gItemTest.RenderEndGroup(sb);
                //}
            }
            //if (!string.IsNullOrEmpty(_shape.Text))
            //{
            //    //RenderText(sb);
            //}

            //if (gItemTest != null)
            //{
            //    gItemTest.RenderEndGroup(sb);
            //}
            sb.AppendLine("</svg>");
        }

        SvgTextBoxItem CreateTextBodyItem()
        {
            if (InsetTextBox == null)
            {
                GetShapeInnerBound(out double x, out double y, out double width, out double height);
                InsetTextBox = new SvgRenderRectItem(this, Bounds);
                InsetTextBox.Bounds.Left = x;
                InsetTextBox.Bounds.Top = y;
                InsetTextBox.Width = width;
                InsetTextBox.Height = height;
                InsetTextBox.Bounds.Parent = TextBox.Bounds; //TODO:Check that textBody is correct.
            }
            var txtBodyItem = new SvgTextBoxItem(this, Bounds, InsetTextBox.Bounds);

            return txtBodyItem;
        }

        //private void RenderText(StringBuilder sb)
        //{
        //    //RenderDebugTextBox(sb);
        //    textBody.Render(sb);
        //}

        private void RenderDebugTextBox(StringBuilder sb)
        {
            InsetTextBox.FillColor = "green";
            InsetTextBox.Render(sb);

            InsetTextBox.GetBounds(out double l, out double t, out double r, out double b);

            //var area = textBody.Bounds;

            ////Temporarily set as child bounds
            //insetTextBox.Bounds.Left = (float)area.Left + l;
            //insetTextBox.Bounds.Top = (float)area.Top + t;
            //insetTextBox.Bounds.Width = (float)area.Width;
            //insetTextBox.Bounds.Height = (float)area.Height;

            //insetTextBox.FillColor = "blue";

            ////Render the inner area
            //insetTextBox.Render(sb);

            ////Reset variables so that the rendering of children later aren't affected
            //insetTextBox.Bounds.Left = l;
            //insetTextBox.Bounds.Top = t;
            //insetTextBox.Bounds.Right = r;
            //insetTextBox.Bounds.Bottom = b;
        }

        private void GetShapeInnerBound(out double x, out double y, out double width, out double height)
        {
            double currentX = 0, currentY = 0, xe, ye;
            x = y = 0;
            width = xe = Bounds.Width;
            height = ye = Bounds.Height;
            foreach (var ri in RenderItems)
            {
                switch (ri.Type)
                {
                    case RenderItemType.Rect:
                        var rectItem = (SvgRenderRectItem)ri;
                        x = rectItem.Left;
                        y = rectItem.Top;
                        width = rectItem.Width;
                        height = rectItem.Height;
                        break;
                    case RenderItemType.Path:
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