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
using EPPlusImageRenderer.Constants;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using TypeConv = OfficeOpenXml.Utils.TypeConversion;

namespace EPPlusImageRenderer.Utils
{
    internal class SvgDrawingWriter
    {
        ExcelTheme _theme;
        DrawingBase _svgDrawing;

        public SvgDrawingWriter(DrawingBase svgDrawing)
        {
            _svgDrawing = svgDrawing;
            var wb = svgDrawing.Drawing._drawings.Worksheet.Workbook;
            _theme = wb.ThemeManager.GetOrCreateTheme();
        }
        internal void WriteSvgDefs(StringBuilder sb, List<RenderItem> renderItems)
        {
            var defSb = new StringBuilder();
            var hs = new HashSet<string>();
            var ix = 1;
            foreach (RenderItem item in renderItems)
            {
                string filter = "";
                if (item.GradientFill != null)
                {
                    string name = WriteGradient("Gradient", defSb, hs, item.GradientFill, item.FillColorSource, true);
                    item.FillColor = $"Url(#{name})";
                }
                else if (item.PatternFill != null)
                {
                    string name = WritePattern($"Pattern{item.PatternFill.PatternType}", defSb, hs, item.PatternFill, item.FillColorSource);
                    item.FillColor = $"Url(#{name})";
                }
                else if (item.BlipFill != null)
                {
                    if (item.FillColorSource != PathFillMode.Norm)
                    {
                        item.FilterName = GetFilterName(ix);
                    }

                    string name = WriteBlip("Blip", defSb, hs, item, ref filter);
                    item.FillColor = $"Url(#{name})";
                }
                if (item.BorderGradientFill != null)
                {
                    string name = WriteGradient("StrokeGradient", defSb, hs, item.BorderGradientFill, item.BorderColorSource, true);
                    item.BorderColor = $"Url(#{name})";
                }
                if(item.GlowColor!=null)
                {
                    if(string.IsNullOrEmpty(item.FilterName))
                    {
                        var filterName = GetFilterName(ix);
                        item.FilterName = $"Url(#{filterName})"; 
                        filter = $"<filter id=\"{filterName}\">";
                    }
                    
                    filter += $"<feGaussianBlur in=\"SourceAlpha\" stdDeviation=\"{item.GlowRadius ?? 0 / 2}\" result=\"blur\"/>" +
                    $"<feFlood flood-color=\"{item.GlowColor}\" flood-opacity=\"0.8\" result=\"glowColor\"/>" +
                    $"<feComposite in=\"glowColor\" in2=\"blur\" operator=\"in\" result=\"coloredBlur\"/>" +
                    $"<feMerge><feMergeNode in=\"coloredBlur\"/><feMergeNode in=\"SourceGraphic\"/></feMerge>";
                }
                if(string.IsNullOrEmpty(filter)==false)
                {
                    defSb.Append(filter+"</filter>");
                }
                ix++;
            }
            if (defSb.Length > 0)
            {
                sb.Append("<defs>");
                sb.Append(defSb);
                sb.Append("</defs>");
            }
        }

        private static string GetFilterName(int ix)
        {
            return $"item{ix}Filter";
        }

        private string WriteBlip(string namePrefix, StringBuilder defSb, HashSet<string> hs, RenderItem item, ref string filter)
        {
            //, item.BlipFill, item.FillColorSource
            var name = $"{namePrefix}";
            var fillMode = item.FillColorSource;
            if (fillMode != PathFillMode.Norm)
            {
                if (hs.Contains(item.FilterName) == false)
                {
                    switch (fillMode)
                    {
                        case PathFillMode.Lighten:
                            filter = $"<filter id=\"{item.FilterName}\"><feColorMatrix type=\"matrix\"\r\n values=\"0.6 0 0 0 0.4\r\n0 0.6 0 0 0.4\r\n0 0 0.6 0 0.4\r\n0 0 0 1 0\" />";
                            break;
                        case PathFillMode.LightenLess:
                            filter = $"<filter id=\"{item.FilterName}\"><feColorMatrix type=\"matrix\"\r\n values=\"0.804 0 0 0 0.196\r\n0 0.804 0 0 0.196\r\n0 0 0.804 0 0.196\r\n0 0 0 1 0\" />";
                            break;
                        case PathFillMode.DarkenLess:
                            filter = $"<filter id=\"{item.FilterName}\"><feColorMatrix type=\"matrix\"\r\n values=\"0.804 0 0 0 0\r\n0 0.804 0 0 0\r\n0 0 0.804 0 0\r\n0 0 0 1 0\" />";
                            break;
                        case PathFillMode.Darken:
                            filter = $"<filter id=\"{item.FilterName}\"><feColorMatrix type=\"matrix\"\r\n values=\"0.6 0 0 0 0\r\n0 0.6 0 0 0\r\n0 0 0.6 0 0\r\n0 0 0 1 0\" />";
                            break;
                    }
                }
                hs.Add(item.FilterName);
            }

            if (hs.Contains(name)) return name;
            hs.Add(name);

            defSb.Append($"<pattern id=\"{name}\" width=\"{item.BlipFill.Image.Bounds.Width}\" height=\"{item.BlipFill.Image.Bounds.Height}\" patternUnits=\"userSpaceOnUse\">");
            defSb.Append($"<image xlink:href=\"{GetImageAsHref(item.BlipFill)}\" {SetStretchTileProps(item.BlipFill)} />");
            defSb.Append($"</pattern>");
            return name;
        }
        private string WriteGradient(string namePrefix, StringBuilder defSb, HashSet<string> hs, DrawGradientFill gradientFill, PathFillMode fillMode, bool userSpaceOnUse)
        {
            var gs = gradientFill.Settings;
            var name = $"{namePrefix}{fillMode}";
            var grUnits = userSpaceOnUse ? " gradientUnits=\"userSpaceOnUse\"" : "";
            if (gs.ShadePath == eShadePath.Linear && hs.Contains(name) == false)
            {
                hs.Add(name);
                var xy = GetXy(gs.LinearSettings?.Angle);
                defSb.Append($"<linearGradient id=\"{name}\"{grUnits} {xy}>");
                SetStopColors(defSb, gradientFill, fillMode);
                defSb.Append("</linearGradient>");
            }
            else if (hs.Contains(name) == false)
            {
                hs.Add(name);
                defSb.Append($"<radialGradient id=\"{name}\" {GetScaling(gradientFill)}>");
                SetStopColors(defSb, gradientFill, fillMode);
                defSb.Append($"</radialGradient>");
            }

            return name;
        }
        private string WritePattern(string namePrefix, StringBuilder defSb, HashSet<string> hs, ExcelDrawingPatternFill patternFill, PathFillMode fillMode)
        {
            var name = $"{namePrefix}{fillMode}";
            var fc = TypeConv.ColorConverter.GetThemeColor(_theme, patternFill.ForegroundColor);
            var bc = TypeConv.ColorConverter.GetThemeColor(_theme, patternFill.BackgroundColor);
            var afc = ColorUtils.GetAdjustedColor(fillMode, fc);
            var abc = ColorUtils.GetAdjustedColor(fillMode, bc);
            switch (patternFill.PatternType)
            {
                case eFillPatternStyle.Pct5:
                    SetPatternHalf(defSb, name, afc, abc, 10, 10);
                    break;
                case eFillPatternStyle.Pct10:
                    SetPatternHalf(defSb, name, afc, abc, 10, 5);
                    break;
                case eFillPatternStyle.Pct20:
                    SetPatternHalf(defSb, name, afc, abc, 4, 4);
                    break;
                case eFillPatternStyle.Pct25:
                    SetPatternHalf(defSb, name, afc, abc, 4, 2);
                    break;
                case eFillPatternStyle.Pct30:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Pct30);
                    break;
                case eFillPatternStyle.Pct40:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Pct40);
                    break;
                case eFillPatternStyle.Pct50:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Pct50);
                    break;
                case eFillPatternStyle.Pct60:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Pct60);
                    break;
                case eFillPatternStyle.Pct70:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Pct70);
                    break;
                case eFillPatternStyle.Pct75:
                    SetPatternHalf(defSb, name, abc, afc, 4, 2);
                    break;
                case eFillPatternStyle.Pct80:
                    SetPatternHalf(defSb, name, abc, afc, 4, 4);
                    break;
                case eFillPatternStyle.Pct90:
                    SetPatternHalf(defSb, name, abc, afc, 10, 5);
                    break;
                case eFillPatternStyle.LtHorz:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LtHorz);
                    break;
                case eFillPatternStyle.LtVert:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LtVert);
                    break;
                case eFillPatternStyle.LtUpDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LtUpDiag);
                    break;
                case eFillPatternStyle.LtDnDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LtDnDiag);
                    break;
                case eFillPatternStyle.DkVert:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DkVert);
                    break;
                case eFillPatternStyle.DkHorz:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DkHorz);
                    break;
                case eFillPatternStyle.DkUpDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DkUpDiag);
                    break;
                case eFillPatternStyle.DkDnDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DkDnDiag);
                    break;
                case eFillPatternStyle.WdUpDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.WdUpDiag);
                    break;
                case eFillPatternStyle.WdDnDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.WdDnDiag);
                    break;
                case eFillPatternStyle.NarVert:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.NarVert);
                    break;
                case eFillPatternStyle.NarHorz:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.NarHorz);
                    break;
                case eFillPatternStyle.Vert:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Vert);
                    break;
                case eFillPatternStyle.Horz:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Horz);
                    break;
                case eFillPatternStyle.DashDnDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DashDnDiag);
                    break;
                case eFillPatternStyle.DashUpDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DashUpDiag);
                    break;
                case eFillPatternStyle.DashHorz:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DashHorz);
                    break;
                case eFillPatternStyle.DashVert:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DashVert);
                    break;
                case eFillPatternStyle.SmConfetti:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.SmConfetti);
                    break;
                case eFillPatternStyle.LgConfetti:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LgConfetti);
                    break;
                case eFillPatternStyle.ZigZag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.ZigZag);
                    break;
                case eFillPatternStyle.Wave:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Wave);
                    break;
                case eFillPatternStyle.DiagBrick:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DiagBrick);
                    break;
                case eFillPatternStyle.HorzBrick:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.HorzBrick);
                    break;
                case eFillPatternStyle.Weave:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Weave);
                    break;
                case eFillPatternStyle.Plaid:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Plaid);
                    break;
                case eFillPatternStyle.Divot:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Divot);
                    break;
                case eFillPatternStyle.DotGrid:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DotGrid);
                    break;
                case eFillPatternStyle.DotDmnd:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DotDmnd);
                    break;
                case eFillPatternStyle.Shingle:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Shingle);
                    break;
                case eFillPatternStyle.Trellis:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Trellis);
                    break;
                case eFillPatternStyle.Sphere:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Sphere);
                    break;
                case eFillPatternStyle.SmGrid:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.SmGrid);
                    break;
                case eFillPatternStyle.LgGrid:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LgGrid);
                    break;
                case eFillPatternStyle.SmCheck:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.SmCheck);
                    break;
                case eFillPatternStyle.LgCheck:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LgCheck);
                    break;
                case eFillPatternStyle.OpenDmnd:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.OpenDmnd);
                    break;
                case eFillPatternStyle.SolidDmnd:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.SolidDmnd);
                    break;
                default:
                    break;
            }
            return name;
        }
        /// <summary>
        /// Sets the pattern to the size with one point top-left and on point middle (x/y).
        /// </summary>
        /// <param name="defSb"></param>
        /// <param name="name"></param>
        /// <param name="afc"></param>
        /// <param name="abc"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        private static void SetPatternHalf(StringBuilder defSb, string name, Color afc, Color abc, int width, int height)
        {
            defSb.Append($"<pattern id=\"{name}\" width=\"{width}\" height=\"{height}\" patternUnits=\"userSpaceOnUse\">");
            defSb.Append($"<rect width=\"{width}\" height=\"{height}\" fill=\"#{abc.To6CharHexString()}\"/>");
            defSb.Append($"<rect x=\"0\" y=\"0\" width=\"1\" height=\"1\" fill=\"#{afc.To6CharHexString()}\"/>");
            defSb.Append($"<rect x=\"{Math.Round(width / 2D, 0)}\" y=\"{Math.Round(height / 2D, 0)}\" width=\"1\" height=\"1\" fill=\"#{afc.To6CharHexString()}\"/>");
            defSb.Append($"</pattern>");
        }
        private static void SetPatternArray(StringBuilder defSb, string name, Color afc, Color abc, short[][] pathArray)
        {
            var height = pathArray.GetLength(0);
            var width = pathArray[0].GetLength(0);
            defSb.Append($"<pattern id=\"{name}\" width=\"{width}\" height=\"{height}\" patternUnits=\"userSpaceOnUse\">");
            defSb.Append($"<rect width=\"{width}\" height=\"{height}\" fill=\"#{abc.To6CharHexString()}\"/>");
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    if (pathArray[y][x] == 1)
                    {
                        defSb.Append($"<rect x=\"{x}\" y=\"{y}\" width=\"1\" height=\"1\" fill=\"#{afc.To6CharHexString()}\" />");
                    }
                }
            }
            defSb.Append($"</pattern>");
        }

        private string GetScaling(DrawGradientFill gradientFill)
        {
            var sb = new StringBuilder();
            var t = gradientFill.Settings.FocusPoint.TopOffset / 100;
            var b = gradientFill.Settings.FocusPoint.BottomOffset / 100;
            var l = gradientFill.Settings.FocusPoint.LeftOffset / 100;
            var r = gradientFill.Settings.FocusPoint.RightOffset / 100;

            var tt = gradientFill.Settings.TileRectangle.TopOffset / 100;
            var tb = gradientFill.Settings.TileRectangle.BottomOffset / 100;
            var tl = gradientFill.Settings.TileRectangle.LeftOffset / 100;
            var tr = gradientFill.Settings.TileRectangle.RightOffset / 100;

            var dy = Math.Abs(t - b);
            var dx = Math.Abs(l - r);
            //var scaleTo = Math.Min(dy, dx);
            //var mult = 0.5 + (scaleTo / 2);
            var cx = _svgDrawing.Bounds.Width * l;
            var cy = _svgDrawing.Bounds.Height * t;
            //var radX = Math.Abs((t - b )) * _svgDrawing.FontSize.Item1;
            //var radY = Math.Abs((l - r)) * _svgDrawing.FontSize.Item2;
            var rx = _svgDrawing.Bounds.Width * 0.5 * (Math.Abs(t - tb) + Math.Abs(b - tt));
            var ry = _svgDrawing.Bounds.Height * 0.5 * (Math.Abs(l - tr) + Math.Abs(r - tl));

            var rad = Math.Sqrt(rx * rx + ry * ry);

            sb.Append($"cx=\"{cx.ToString(CultureInfo.InvariantCulture)}\" cy=\"{cy.ToString(CultureInfo.InvariantCulture)}\" ");

            //if(tt-tb != 0)
            //{
            //    sb.Append($"fy=\"{fy.ToString(CultureInfo.InvariantCulture)}\" ");
            //}
            if (dy > 0 && dy < 1)
            {
                var fy = cy * (1 + dy);
                sb.Append($"fy=\"{fy.ToString(CultureInfo.InvariantCulture)}\" ");
            }

            if (dx > 0 && dx < 1)
            {
                var fx = cx * (1 + dx);
                sb.Append($"fx=\"{fx.ToString(CultureInfo.InvariantCulture)}\" ");
            }
            sb.Append($"r=\"{rad.ToString(CultureInfo.InvariantCulture)}\" ");
            sb.Append($"gradientUnits=\"userSpaceOnUse\" spreadMethod=\"pad\" gradientTransform=\"matrix(1 0 0 1 1 1)\" ");

            return sb.ToString();
        }

        private void SetStopColors(StringBuilder defSb, DrawGradientFill gradientFill, PathFillMode fillMode)
        {
            int ix = 0;
            foreach (var c in gradientFill.Colors)
            {
                var color = ColorUtils.GetAdjustedColor(fillMode, c.Color);
                // TODO: check if ix should be increased...?
                defSb.Append($"<stop offset=\"{c.Position}%\" stop-color=\"#{color.To6CharHexString()}\" {GetOpacity(gradientFill.Settings.Colors[ix].Color)} />");
            }
        }

        private string GetXy(double? angle)
        {
            if (angle.HasValue && angle != 0)
            {
                var x1 = 0D;
                var x2 = 0D;
                var y1 = 0D;
                var y2 = 0D;
                angle %= 360;
                if (angle <= 90)
                {
                    x2 = 1D - Math.Sin(MathHelper.Radians(angle.Value));
                    y2 = Math.Sin(MathHelper.Radians(angle.Value));
                }
                else if (angle <= 180)
                {
                    y2 = Math.Sin(MathHelper.Radians(angle.Value));
                    x1 = 1D - Math.Sin(MathHelper.Radians(angle.Value));
                }
                else if (angle <= 270)
                {
                    y1 = Math.Sin(MathHelper.Radians(angle.Value - 180));
                    x1 = 1D - Math.Sin(MathHelper.Radians(angle.Value - 180));
                }
                else
                {
                    y1 = Math.Sin(MathHelper.Radians(angle.Value - 180));
                    x2 = 1D - Math.Sin(MathHelper.Radians(angle.Value - 180));
                }

                return $" x1=\"{x1.ToString("0.00", CultureInfo.InvariantCulture)}\" x2=\"{x2.ToString("0.00", CultureInfo.InvariantCulture)}\" y1=\"{y1.ToString("0.00", CultureInfo.InvariantCulture)}\" y2=\"{y2.ToString("0.00", CultureInfo.InvariantCulture)}\"";
            }
            return "";
        }

        private string GetOpacity(ExcelDrawingColorManager c)
        {
            var opacetyTransform = c.Transforms?.FirstOrDefault(x => x.Type == OfficeOpenXml.Drawing.Style.Coloring.eColorTransformType.Alpha);
            if (opacetyTransform == null) return "";

            return $"stop-opacity=\"{opacetyTransform.Value.ToString("0")}%\"";
        }

        private string SetStretchTileProps(ExcelDrawingBlipFill blipFill)
        {
            if (blipFill.Stretch)
            {
                var x = _svgDrawing.Bounds.Width * blipFill.StretchOffset.LeftOffset / 100;
                var y = _svgDrawing.Bounds.Height * blipFill.StretchOffset.TopOffset / 100;
                var width = _svgDrawing.Bounds.Width - x - _svgDrawing.Bounds.Width * blipFill.StretchOffset.RightOffset / 100;
                var height = _svgDrawing.Bounds.Height - x - _svgDrawing.Bounds.Height * blipFill.StretchOffset.BottomOffset / 100;
                return $" preserveAspectRatio=\"none\" x=\"{x.ToString(CultureInfo.InvariantCulture)}\" y=\"{y.ToString(CultureInfo.InvariantCulture)}\" width=\"{width.ToString(CultureInfo.InvariantCulture)}\" height=\"{height.ToString(CultureInfo.InvariantCulture)}\" ";
            }
            else if (!(blipFill.Tile.HorizontalOffset == 0 && blipFill.Tile.VerticalOffset == 0 &&
                    blipFill.Tile.HorizontalRatio == 100 && blipFill.Tile.VerticalRatio == 100 && blipFill.Tile.FlipMode == eTileFlipMode.None))
            {
                var flip = "";
                switch (blipFill.Tile.FlipMode)
                {
                    case eTileFlipMode.X:
                        flip = $" transform=\"translate({_svgDrawing.Bounds.Width.ToString(CultureInfo.InvariantCulture)}, 0) scale(-1, 1)\"";
                        break;
                    case eTileFlipMode.Y:
                        flip = $" transform=\"translate(0, {_svgDrawing.Bounds.Height.ToString(CultureInfo.InvariantCulture)}) scale(1, -1)\"";
                        break;
                    case eTileFlipMode.XY:
                        flip = $" transform=\"translate({_svgDrawing.Bounds.Width.ToString(CultureInfo.InvariantCulture)}, {_svgDrawing.Bounds.Height.ToString(CultureInfo.InvariantCulture)}) scale(-1, -1)\"";
                        break;
                }
                return $"{flip}";
            }

            return "";
        }

        private object GetImageAsHref(ExcelDrawingBlipFill blipFill)
        {
            return $"data:{blipFill.ContentType};base64," + Convert.ToBase64String(blipFill.Image.ImageBytes);
        }
    }
}
