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
using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.ShapeDefinitions;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Export.Utils;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using OfficeOpenXml.Utils.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.Renderer
{
    internal class ShapeRenderer : DrawingRenderer
    {
        /// <summary>
        /// Calculated shape textbox
        /// </summary>
        RectRenderItem InsetTextBox;

        RectRenderItem MarginTextBox;

        /// <summary>
        /// Textbox from memory
        /// </summary>
        public DrawingTextbody TextBody{ get; internal set; }

        //public ShapeRenderer(eShapeStyle style, double top, double left, double width, double height, eTextAutofit autofit) : base()
        //{
        //    var parentBounds = new BoundingBox(top, left, width, height);
        //    var shapeGroup = new GroupRenderItem(parentBounds);

        //}

        public ShapeRenderer(ExcelShape shape) : base(shape)
        {
            var style = shape.Style;

            var parentBounds = shape.GetBoundingBox();
            var shapeGroup = new GroupRenderItem(parentBounds, shape.Rotation);
            RenderItems.Add(shapeGroup);

            if (style==eShapeStyle.CustomShape)
            {
                foreach (var path in shape.CustomGeom.DrawingPaths)
                {
                    shapeGroup.RenderItems.Add(AddFromPaths(Bounds, path));
                }
            }
            else
            {
                var shapeDef = PresetShapeDefinitions.ShapeDefinitions[(ShapeStyle)style].Clone();
                if (shape.HasCustomAdjustmentPoints())
                {
                    shapeDef.Calculate(shape._width, shape._height, shape.TextBody.TextAutofit == eTextAutofit.ShapeAutofit, shape.GetAdjustmentPointsNames().ToList(), shape.GetAdjustmentPointsList().ToList());
                }
                else
                {
                    shapeDef.Calculate(shape._width, shape._height, shape.TextBody.TextAutofit == eTextAutofit.ShapeAutofit, null, null);
                }

                //Draw Filled path's
                foreach (var path in shapeDef.ShapePaths)
                {
                    if (path.Fill != PathFillMode.None)
                    {
                        shapeGroup.RenderItems.Add(AddFromPaths(parentBounds, path, true, false));
                    }
                }

                //Draw border path's
                foreach (var path in shapeDef.ShapePaths)
                {
                    if (path.Stroke)
                    {
                        shapeGroup.RenderItems.Add(AddFromPaths(parentBounds, path, false, true));
                    }
                }

                if (shape.Text != null)
                {
                    if (shapeDef.TextBoxRect != null)
                    {
                        InsetTextBox = new RectRenderItem(Bounds);
                        InsetTextBox.Bounds.Left = (float)shapeDef.TextBoxRect.LeftValue.PixelToPoint();
                        InsetTextBox.Bounds.Top = (float)shapeDef.TextBoxRect.TopValue.PixelToPoint();
                        InsetTextBox.Bounds.Left = (float)shapeDef.TextBoxRect.LeftValue.PixelToPoint();
                        InsetTextBox.Bounds.Top = (float)shapeDef.TextBoxRect.TopValue.PixelToPoint();
                        InsetTextBox.FillOpacity = 0.3d;

                        if (shape.TextBody.TextAutofit != eTextAutofit.ShapeAutofit)
                        {
                            InsetTextBox.Width = ((double)((float)shapeDef.TextBoxRect.RightValue - (float)shapeDef.TextBoxRect.LeftValue)).PixelToPoint();
                            InsetTextBox.Height = ((double)((float)shapeDef.TextBoxRect.BottomValue - (float)shapeDef.TextBoxRect.TopValue)).PixelToPoint();
                            InsetTextBox.Width = ((double)((float)shapeDef.TextBoxRect.RightValue - (float)shapeDef.TextBoxRect.LeftValue)).PixelToPoint();
                            InsetTextBox.Height = ((double)((float)shapeDef.TextBoxRect.BottomValue - (float)shapeDef.TextBoxRect.TopValue)).PixelToPoint();
                        }
                        else
                        {
                            InsetTextBox.Width = (float)shapeDef.TextBoxRect.RightValue.PixelToPoint();
                            InsetTextBox.Height = (float)shapeDef.TextBoxRect.BottomValue.PixelToPoint();
                            InsetTextBox.Width = (float)shapeDef.TextBoxRect.RightValue.PixelToPoint();
                            InsetTextBox.Height = (float)shapeDef.TextBoxRect.BottomValue.PixelToPoint();
                        }
                    }
                    else
                    {
                        InsetTextBox = null;
                    }

                    InsetTextBox.FillOpacity = 0.3d;

                    TextBody = CreateTextBodyItem(shape.TextBody);
                }
            }
        }

        protected RenderItem AddFromPaths(BoundingBox parent, DrawingPath path, bool drawFill = true, bool drawBorder = true)
        {
            var pi = new PathRenderItem(parent);
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
                        pi.Commands.Add(new PathCommands(PathCommandType.End));
                        cmd = null;
                        break;
                }
                //p.TranslateCoordiantesToPointsAndDegrees(ExcelDrawing.EMU_PER_POINT, 1);
                pCmd = p;
            }
            if (coordinates.Count > 0)
            {
                pi.Commands[pi.Commands.Count - 1].Coordinates = coordinates.ToArray();
            }
            var shape = (ExcelShape)Drawing;
            if (drawFill)
            {
                pi.FillColorSource = path.Fill;                
                pi.SetDrawingPropertiesFill(Theme, shape.Fill, shape.ThemeStyles.FillReference.Color);
            }
            else
            {
                pi.FillColorSource = PathFillMode.None;
                pi.FillColor = "none";
            }

            if (drawBorder)
            {
                pi.BorderColorSource = path.Stroke ? PathFillMode.Norm : PathFillMode.None;
                pi.SetDrawingPropertiesBorder(Theme, shape.Border, shape.ThemeStyles.BorderReference.Color, path.Stroke);
            }
            else
            {
                pi.BorderColorSource = PathFillMode.None;
                pi.BorderColor = "none";
            }

            return pi;
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
                return $"{(Bounds.Left).PointToPixelString()},{Bounds.Top.PointToPixelString()},{Bounds.Right.PointToPixelString()},{Bounds.Bottom.PointToPixelString()}";
            }
        }
        DrawingTextbody CreateTextBodyItem(ExcelTextBody bodyOrig)
        {
            if (InsetTextBox == null)
            {
                GetShapeInnerBound(out double x, out double y, out double width, out double height);
                InsetTextBox = new RectRenderItem(Bounds);
                InsetTextBox.Bounds.Left = x.PixelToPoint();
                InsetTextBox.Bounds.Top = y.PixelToPoint();
                InsetTextBox.Width = width.PixelToPoint();
                InsetTextBox.Height = height.PixelToPoint();
                //InsetTextBox.Bounds.Parent = RenderTextbox.Parent; //TODO:Check that textBody is correct.
                InsetTextBox.Bounds.Left = x.PixelToPoint();
                InsetTextBox.Bounds.Top = y.PixelToPoint();
                InsetTextBox.Width = width.PixelToPoint();
                InsetTextBox.Height = height.PixelToPoint();
                //InsetTextBox.Bounds.Parent = RenderTextbox.Parent; //TODO:Check that textBody is correct.
            }

            double l, r, t, b;
            bodyOrig.GetInsetsOrDefaults(out l, out t, out r, out b);

            MarginTextBox = new RectRenderItem(this.Bounds);

            MarginTextBox.Top = t + InsetTextBox.Top;
            MarginTextBox.Left = l + InsetTextBox.Left;
            MarginTextBox.Width = InsetTextBox.Width - r - l;
            MarginTextBox.Height = InsetTextBox.Height - b - t;

            var grp = new GroupRenderItem(MarginTextBox.Bounds);
            RenderItems.Add(grp);

            var txtBodyItem = new DrawingTextbody(Drawing, MarginTextBox.Bounds, MarginTextBox.Left, MarginTextBox.Top, MarginTextBox.Width, MarginTextBox.Height);
            txtBodyItem.ImportTextBody(bodyOrig);

            txtBodyItem.AppendRenderItems(grp.RenderItems);

            //ChartAreaRenderItems.Add(new SvgEndGroupItem(this, Bounds));
            
            return txtBodyItem;
        }

        //private void RenderText(StringBuilder sb)
        //{
        //    //RenderDebugTextBox(sb);
        //    textBody.Render(sb);
        //}

        //private void RenderDebugTextBox(StringBuilder sb)
        //{
        //    InsetTextBox.FillOpacity = 0.3d;
        //    InsetTextBox.FillColor = "green";
        //    InsetTextBox.Render(sb);

        //    MarginTextBox.FillColor = "red";
        //    MarginTextBox.FillOpacity = 0.3;
        //    MarginTextBox.Render(sb);
        //    //InsetTextBox.GetBounds(out double l, out double t, out double r, out double b);

        //    //var area = textBody.Bounds;

        //    ////Temporarily set as child bounds
        //    //insetTextBox.Bounds.Left = (float)area.Left + l;
        //    //insetTextBox.Bounds.Top = (float)area.Top + t;
        //    //insetTextBox.Bounds.Width = (float)area.Width;
        //    //insetTextBox.Bounds.Height = (float)area.Height;

        //    //insetTextBox.FillColor = "blue";

        //    ////Render the inner area
        //    //insetTextBox.Render(sb);

        //    ////Reset variables so that the rendering of children later aren't affected
        //    //insetTextBox.Bounds.Left = l;
        //    //insetTextBox.Bounds.Top = t;
        //    //insetTextBox.Bounds.Right = r;
        //    //insetTextBox.Bounds.Bottom = b;
        //}

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
                        var rectItem = (RectRenderItem)ri;
                        x = rectItem.Left;
                        y = rectItem.Top;
                        width = rectItem.Width;
                        height = rectItem.Height;
                        break;
                    case RenderItemType.Path:
                        var pathItem = (PathRenderItem)ri;
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
        protected static void AddCmd(PathRenderItem pi, DrawingPath path, List<double> coordinates, ref PathCommands cmd, PathsBase pp, PathsBase p, PathCommandType commandType)
        {
            if (pp == null || pp.Type != p.Type)
            {
                SetCmdCoordinats(cmd, p, coordinates);
                cmd = new PathCommands(commandType);
                pi.Commands.Add(cmd);
            }
            AddToCoordinates(path, coordinates, p);
        }
        protected static void AddArc(PathRenderItem pi, DrawingPath path, List<double> coordinates, PathsBase pCmd, out double startPointX, out double startPointY, PathsBase p)
        {
            //var width = ((double)path.Width.Value / ExcelDrawing.EMU_PER_PIXEL);
            //var height = ((double)path.Height.Value / ExcelDrawing.EMU_PER_PIXEL);
            var arc = (ArcTo)p;
            PathCommands c = null;
            startPointX = pCmd.EndX;
            startPointY = pCmd.EndY;
            if (startPointX != 0) startPointX /= ExcelDrawing.EMU_PER_POINT;
            if (startPointY != 0) startPointY /= ExcelDrawing.EMU_PER_POINT;
            var wR = arc.WidthRadius.Value / ExcelDrawing.EMU_PER_POINT;
            var hR = arc.HeightRadius.Value / ExcelDrawing.EMU_PER_POINT;
            if (wR == 0 && hR == 0)
            {
                return;
            }
            var stA = arc.StartAngle.Value / 60000d;
            var swA = arc.SwingAngle.Value / 60000d;

            while (swA != 0)
            {
                var aAdd = swA < 0 ? Math.Max(swA, -180) : Math.Min(swA, 180);
                var endAngle = AngleToRadians(stA + aAdd);

                var stA_Adj = stA < 0 ? (stA + 360) % 360 : stA;
                var adjRads = AngleToRadians(stA_Adj);

                //Start and End angles are NOT the 't' angle of the equations we use.
                //The angles we are given are DIRECTLY against the ellipse. Or point 'P' in a parametric form
                //Therefore we have to use the angle we have to calculate the angles needed for our formulas.
                var angleT = Math.Atan((wR * Math.Tan(adjRads)) / hR);
                var angleTEnd = Math.Atan((wR * Math.Tan(endAngle)) / hR);

                //Atan can only return values on positive x 90° to -90°
                //So we must adjust by adding Pi (180°) if x of the angle is negative
                if (Math.Cos(adjRads) < 0)
                {
                    angleT += (Math.Round((double)System.Math.PI, 14));
                }
                if (Math.Cos(endAngle) < 0)
                {
                    angleTEnd += (Math.Round((double)System.Math.PI, 14));
                }

                var centerX = startPointX - (wR * Math.Cos(angleT));
                var centerY = startPointY - (hR * Math.Sin(angleT));
                var endX = (double)centerX + (wR * Math.Cos(angleTEnd));
                var endY = (double)centerY + (hR * Math.Sin(angleTEnd));
                c = new PathCommands(PathCommandType.Arc, wR, hR, 0, 0, swA < 0 ? 0 : 1, endX, endY);
                pi.Commands.Add(c);
                stA += aAdd;
                swA -= aAdd;
                if (wR != 0)
                {
                    startPointX = endX;
                }
                if (hR != 0)
                {
                    startPointY = endY;
                }
                ((ArcTo)p).SetEndCoordinates(endX * ExcelDrawing.EMU_PER_POINT, endY * ExcelDrawing.EMU_PER_POINT);
            }
        }

        protected static double AngleToRadians(double angle)
        {
            return MConverter.DegreesToRadians(angle);
        }
        protected static void SetCmdCoordinats(PathCommands cmd, PathsBase p, List<double> coordinates)
        {
            if (cmd != null)
            {
                cmd.Coordinates = coordinates.ToArray();
                if (cmd.Coordinates.Length > 0)
                {
                    coordinates.Clear();
                }
            }
        }
        private static void AddToCoordinates(DrawingPath path, List<double> coordinates, PathsBase p)
        {
            var mt = (PathWithCoordinates)p;
            foreach (var c in mt.Coordinates)
            {
                coordinates.Add(c.X.Value / ExcelDrawing.EMU_PER_POINT);
                coordinates.Add(c.Y.Value / ExcelDrawing.EMU_PER_POINT);
            }
        }
    }
}