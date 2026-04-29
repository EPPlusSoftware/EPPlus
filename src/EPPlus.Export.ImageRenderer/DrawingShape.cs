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
using EPPlus.Export.ImageRenderer.Utils;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Collections.Generic;

namespace EPPlusImageRenderer
{
    internal abstract class DrawingShape : DrawingBase
    {
        protected ExcelShape _shape;
        protected DrawingShape(ExcelShape shape) : base(shape)
        {
            var style = shape.Style;

            _shape = shape;
        }
        protected static void AddCmd(SvgRenderPathItem pi, DrawingPath path, List<double> coordinates, ref PathCommands cmd, PathsBase pp, PathsBase p, PathCommandType commandType)
        {
            if (pp == null || pp.Type != p.Type)
            {
                SetCmdCoordinats(cmd, p, coordinates);
                cmd = new PathCommands(commandType, pi);
                pi.Commands.Add(cmd);
            }
            AddToCoordinates(path, coordinates, p);
        }
        protected static void AddArc(SvgRenderPathItem pi, DrawingPath path, List<double> coordinates, PathsBase pCmd, out double startPointX, out double startPointY, PathsBase p)
        {
            var width = ((double)path.Width.Value / ExcelDrawing.EMU_PER_PIXEL);
            var height = ((double)path.Height.Value / ExcelDrawing.EMU_PER_PIXEL);
            var arc = (ArcTo)p;
            PathCommands c = null;
            startPointX = pCmd.EndX;
            startPointY = pCmd.EndY;
            if (startPointX != 0) startPointX /= ExcelDrawing.EMU_PER_PIXEL;
            if (startPointY != 0) startPointY /= ExcelDrawing.EMU_PER_PIXEL;
            var wR = arc.WidthRadius.Value / (float)ExcelDrawing.EMU_PER_PIXEL;
            var hR = arc.HeightRadius.Value / (float)ExcelDrawing.EMU_PER_PIXEL;
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
                c = new PathCommands(PathCommandType.Arc, pi, (float)wR / width, (float)hR / height, 0, 0, swA < 0 ? 0 : 1, endX / width, endY / height);
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
                ((ArcTo)p).SetEndCoordinates(endX * ExcelDrawing.EMU_PER_PIXEL, endY * ExcelDrawing.EMU_PER_PIXEL);
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
                coordinates.Add(c.X.Value / path.Width.Value);
                coordinates.Add(c.Y.Value / path.Height.Value);
            }
        }
    }
}