using DrawingRenderer.Constants;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.DrawingRenderer.Utils;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using EPPlusImageRenderer.Utils;
using System.Drawing;
using System.Globalization;
using System.Text;
using OfficeOpenXml.Utils;
using EPPlus.Fonts.OpenType.Utils;
using System.Security.Cryptography.Xml;
using System.ComponentModel;

namespace EPPlus.DrawingRenderer
{
    public class SvgShapeRenderer : IShapeRenderer<StringBuilder>
    {
        public SvgShapeRenderer(BoundingBox bounds, StringBuilder outputStream, SvgRenderOptions options)
        {
            BasicShapesRenderer = new SvgBasicShapesRenderer(outputStream);

            //Override size.
            if (options.Size.Width.HasValue)
            {
                bounds.Width = options.Size.WidthPixels;
            }
            if (options.Size.Height.HasValue)
            {
                bounds.Height = options.Size.HeightPixels;
            }

            Bounds = bounds;
            OutputStream = outputStream;
        }
        public IBasicIShapesRenderer<StringBuilder> BasicShapesRenderer { get; }

        public StringBuilder OutputStream { get; }

        public BoundingBox Bounds { get; }
        public string ViewBox { get; set; }
        public bool Render(List<RenderItem> items)
        {
            OutputStream.Clear();
            OutputStream.Append($"<svg width=\"{Bounds.Width.PointToPixelString()}\" height=\"{Bounds.Height.PointToPixelString()}\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:space=\"default\" Overflow=\"Hidden\" viewbox=\"{ViewBox}\">");
            PreRender(items);
            foreach (var item in items)
            {
                switch (item.Type)
                {
                    case RenderItemType.Group:
                        BasicShapesRenderer.GroupRenderer.Render((GroupRenderItem)item);
                        break;
                    case RenderItemType.Line:
                        BasicShapesRenderer.LineRenderer.Render((LineRenderItem)item);
                        break;
                    case RenderItemType.Rect:
                        BasicShapesRenderer.RectangleRenderer.Render((RectRenderItem)item);
                        break;
                    case RenderItemType.Ellipse:
                        BasicShapesRenderer.EllipseRenderer.Render((EllipseRenderItem)item);
                        break;
                    case RenderItemType.Path:
                        BasicShapesRenderer.PathRenderer.Render((PathRenderItem)item);
                        break;
                    case RenderItemType.Paragraph:
                        BasicShapesRenderer.ParagraphRenderer.Render((ParagraphRenderItem)item);
                        break;
                    case RenderItemType.UseReference:
                        BasicShapesRenderer.UseReferenceRenderer.Render((UseReferenceRenderItem)item);
                        break;
                }
            }
            OutputStream.AppendLine("</svg>");
            return true;
        }


        public void PreRenderGroup(List<RenderItem> items, StringBuilder defSb, HashSet<string> hs, ref int ix)
        {
            foreach (RenderItem item in items)
            {
                if (item is GroupRenderItem group)
                {
                    PreRenderGroup(group.RenderItems, defSb, hs, ref ix);
                    //foreach(var child in group.RenderItems)
                    //{
                    //    WriteDefsForRenderItem(defSb, hs, ref ix, child);
                    //}
                }
                else
                {
                    WriteDefsForRenderItem(defSb, hs, ref ix, item);
                }
            }
        }


        public bool PreRender(List<RenderItem> items)
        {
            var defSb = new StringBuilder();
            var hs = new HashSet<string>();
            var ix = 1;

            foreach (RenderItem item in items)
            {
                PreRenderGroup(items, defSb, hs, ref ix);
            }

            if (defSb.Length > 0)
            {
                OutputStream.Append("<defs>");
                OutputStream.Append(defSb);
                OutputStream.Append("</defs>");
            }

            ////Remove all items that have already been rendered
            //renderItems.RemoveAll(x => x is SvgRenderItem svgX && string.IsNullOrEmpty(svgX.DefId) == false);
            return true;
        }
        Dictionary<string, string> _defsCache = new Dictionary<string, string>();
        private void WriteDefsForRenderItem(StringBuilder defSb, HashSet<string> hs, ref int ix, RenderItem item)
        {
            string filter = "";
            if (item.GradientFill != null)
            {
                var key = item.GradientFill.GetKey();
                if (_defsCache.TryGetValue(key, out string? name) == false)
                {
                    name = WriteGradient($"Gradient{ix}", defSb, hs, item, item.GradientFill, item.FillColorSource);
                    _defsCache[key] = name;
                }
                item.FillColor = $"Url(#{name})";
            }
            else if (item.PatternFill != null)
            {
                var key = item.PatternFill.GetKey();
                if (_defsCache.TryGetValue(key, out string? name) == false)
                {
                    name = WritePattern($"Pattern{item.PatternFill.PatternType}{ix}", defSb, hs, item.PatternFill, item.FillColorSource);
                    _defsCache[key] = name;
                }
                item.FillColor = $"Url(#{name})";
            }
            else if (item.BlipFill != null)
            {
                var key = item.BlipFill.GetKey();

                if (_defsCache.TryGetValue(key, out string? name) == false)
                {
                    if (item.FillColorSource != PathFillMode.Norm)
                    {
                        item.FilterName = GetFilterName(ix);
                    }

                    name = WriteBlip("Blip", defSb, hs, item, ref filter);
                    _defsCache[key] = name;
                }
                item.FillColor = $"Url(#{name})";
            }
            if (item.BorderGradientFill != null)
            {
                var key = item.BorderGradientFill.GetKey() + item.Bounds.UniqueKey;
                if (_defsCache.TryGetValue(key, out string? name) == false)
                {
                    name = WriteGradient($"StrokeGradient{ix}", defSb, hs, item, item.BorderGradientFill, item.BorderColorSource);
                    _defsCache[key] = name;
                }
                item.BorderColor = $"Url(#{name})";
            }
            if (item.GlowColor != null)
            {
                var key = item.GetFilterKey();
                if (_defsCache.TryGetValue(key, out string? name) == false)
                {
                    if (string.IsNullOrEmpty(item.FilterName))
                    {
                        name = GetFilterName(ix);
                        item.FilterName = $"Url(#{name})";
                        filter = $"<filter id=\"{name}\" filterUnits=\"userSpaceOnUse\">";
                        double stdev = (item.GlowRadius??0) / 4D;
                        filter += $"<feMorphology in=\"SourceAlpha\" operator=\"dilate\" radius=\"{stdev.ToString("#.###", CultureInfo.InvariantCulture)}\" result=\"thick\"/>" +
                        $"<feGaussianBlur in=\"thick\" stdDeviation=\"{stdev.ToString("#.###",CultureInfo.InvariantCulture)}\" result=\"blur\"/>" +
                        $"<feFlood flood-color=\"{item.GlowColor}\" flood-opacity=\"{item.GlowOpacity}%\" result=\"glowColor\"/>" +
                        $"<feComposite in=\"glowColor\" in2=\"blur\" operator=\"in\" result=\"coloredBlur\"/>" +
                        $"<feMerge><feMergeNode in=\"coloredBlur\"/><feMergeNode in=\"SourceGraphic\"/></feMerge>";

                        _defsCache[key] = name;
                    }
                }
                else
                {
                    item.FilterName = name;
                }
            }
            if (item.OuterShadowEffect != null)
            {
                if (string.IsNullOrEmpty(item.FilterName))
                {
                    var key = item.GetFilterKey();
                    if (_defsCache.TryGetValue(key, out string? name) == false)
                    {
                        var filterName = GetFilterName(ix);
                        item.FilterName = $"Url(#{filterName})";
                        filter = $"<filter id=\"{filterName}\" >";
                        _defsCache[key] = filterName;
                    }
                    else
                    {
                        item.FilterName = $"Url(#{name})";
                    }
                }
                if (string.IsNullOrEmpty(filter) == false)
                {
                    item.GetOuterShadowColor(out string shadowColor, out double opacity);
                    var dx = Math.Round(item.OuterShadowEffect.Distance * Math.Cos(MathHelper.Radians(item.OuterShadowEffect.Direction ?? 0D)), 2);
                    var dy = Math.Round(item.OuterShadowEffect.Distance * Math.Sin(MathHelper.Radians(item.OuterShadowEffect.Direction ?? 0D)), 2);
                    var blurRadius = item.OuterShadowEffect.BlurRadius ?? 0D / 2;
                    filter += $"<feDropShadow dx=\"{dx.PointToPixelString()}\" dy=\"{dy.PointToPixelString()}\" stdDeviation=\"{blurRadius.PointToPixelString()}\" flood-color=\"{shadowColor}\" flood-opacity=\"{opacity.ToString("N2", CultureInfo.InvariantCulture)}\" />";
                }
            }
            if (string.IsNullOrEmpty(filter) == false)
            {
                defSb.Append(filter + "</filter>");
            }
            ix++;
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

            defSb.Append($"<pattern id=\"{name}\" width=\"{item.BlipFill.ImageBounds.Width}\" height=\"{item.BlipFill.ImageBounds.Height}\" patternUnits=\"userSpaceOnUse\">");
            defSb.Append($"<image xlink:href=\"{GetImageAsHref(item.BlipFill)}\" {SetStretchTileProps(item.BlipFill)} />");
            defSb.Append($"</pattern>");
            return name;
        }
        private string WriteGradient(string namePrefix, StringBuilder defSb, HashSet<string> hs, RenderItem item, RenderGradientFill? gradientFill, PathFillMode fillMode)
        {
            var name = GetGradientName(namePrefix, fillMode, gradientFill);

            if (gradientFill.ShadePath == ShadePath.Linear && hs.Contains(name) == false)
            {
                hs.Add(name);
                var grUnits = gradientFill.UserSpaceOnUse != UserSpaceSettings.ObjectBoundingBox ? " gradientUnits=\"userSpaceOnUse\"" : "";
                var xy = GetXy(item, gradientFill.UserSpaceOnUse, gradientFill.LinearSettings?.Angle);
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

        private string GetGradientName(string namePrefix, PathFillMode fillMode, RenderGradientFill? gradientFill)
        {
            //if(gradientFill?.UserSpaceOnUse==UserSpaceSettings.UserSpaceOnUse_Parent || gradientFill?.LinearSettings?.Angle!=0)
            //{
            //    return $"{namePrefix}{fillMode}";
            //}
            return $"{namePrefix}{fillMode}";
        }

        private string WritePattern(string namePrefix, StringBuilder defSb, HashSet<string> hs, RenderPatternFill patternFill, PathFillMode fillMode)
        {
            var name = $"{namePrefix}{fillMode}";
            var afc = ColorUtils.GetAdjustedColor(fillMode, patternFill.ForegroundColor);
            var abc = ColorUtils.GetAdjustedColor(fillMode, patternFill.BackgroundColor);
            switch (patternFill.PatternType)
            {
                case FillPatternStyle.Pct5:
                    SetPatternHalf(defSb, name, afc, abc, 10, 10);
                    break;
                case FillPatternStyle.Pct10:
                    SetPatternHalf(defSb, name, afc, abc, 10, 5);
                    break;
                case FillPatternStyle.Pct20:
                    SetPatternHalf(defSb, name, afc, abc, 4, 4);
                    break;
                case FillPatternStyle.Pct25:
                    SetPatternHalf(defSb, name, afc, abc, 4, 2);
                    break;
                case FillPatternStyle.Pct30:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Pct30);
                    break;
                case FillPatternStyle.Pct40:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Pct40);
                    break;
                case FillPatternStyle.Pct50:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Pct50);
                    break;
                case FillPatternStyle.Pct60:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Pct60);
                    break;
                case FillPatternStyle.Pct70:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Pct70);
                    break;
                case FillPatternStyle.Pct75:
                    SetPatternHalf(defSb, name, abc, afc, 4, 2);
                    break;
                case FillPatternStyle.Pct80:
                    SetPatternHalf(defSb, name, abc, afc, 4, 4);
                    break;
                case FillPatternStyle.Pct90:
                    SetPatternHalf(defSb, name, abc, afc, 10, 5);
                    break;
                case FillPatternStyle.LtHorz:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LtHorz);
                    break;
                case FillPatternStyle.LtVert:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LtVert);
                    break;
                case FillPatternStyle.LtUpDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LtUpDiag);
                    break;
                case FillPatternStyle.LtDnDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LtDnDiag);
                    break;
                case FillPatternStyle.DkVert:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DkVert);
                    break;
                case FillPatternStyle.DkHorz:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DkHorz);
                    break;
                case FillPatternStyle.DkUpDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DkUpDiag);
                    break;
                case FillPatternStyle.DkDnDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DkDnDiag);
                    break;
                case FillPatternStyle.WdUpDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.WdUpDiag);
                    break;
                case FillPatternStyle.WdDnDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.WdDnDiag);
                    break;
                case FillPatternStyle.NarVert:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.NarVert);
                    break;
                case FillPatternStyle.NarHorz:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.NarHorz);
                    break;
                case FillPatternStyle.Vert:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Vert);
                    break;
                case FillPatternStyle.Horz:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Horz);
                    break;
                case FillPatternStyle.DashDnDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DashDnDiag);
                    break;
                case FillPatternStyle.DashUpDiag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DashUpDiag);
                    break;
                case FillPatternStyle.DashHorz:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DashHorz);
                    break;
                case FillPatternStyle.DashVert:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DashVert);
                    break;
                case FillPatternStyle.SmConfetti:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.SmConfetti);
                    break;
                case FillPatternStyle.LgConfetti:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LgConfetti);
                    break;
                case FillPatternStyle.ZigZag:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.ZigZag);
                    break;
                case FillPatternStyle.Wave:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Wave);
                    break;
                case FillPatternStyle.DiagBrick:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DiagBrick);
                    break;
                case FillPatternStyle.HorzBrick:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.HorzBrick);
                    break;
                case FillPatternStyle.Weave:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Weave);
                    break;
                case FillPatternStyle.Plaid:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Plaid);
                    break;
                case FillPatternStyle.Divot:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Divot);
                    break;
                case FillPatternStyle.DotGrid:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DotGrid);
                    break;
                case FillPatternStyle.DotDmnd:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.DotDmnd);
                    break;
                case FillPatternStyle.Shingle:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Shingle);
                    break;
                case FillPatternStyle.Trellis:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Trellis);
                    break;
                case FillPatternStyle.Sphere:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.Sphere);
                    break;
                case FillPatternStyle.SmGrid:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.SmGrid);
                    break;
                case FillPatternStyle.LgGrid:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LgGrid);
                    break;
                case FillPatternStyle.SmCheck:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.SmCheck);
                    break;
                case FillPatternStyle.LgCheck:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.LgCheck);
                    break;
                case FillPatternStyle.OpenDmnd:
                    SetPatternArray(defSb, name, afc, abc, PatternArrays.OpenDmnd);
                    break;
                case FillPatternStyle.SolidDmnd:
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


            defSb.Append($"<rect width=\"{width}\" height=\"{height}\" {WriteFillColor(abc)}");
            defSb.Append("/>");
            defSb.Append($"<rect x=\"0\" y=\"0\" width=\"1\" height=\"1\" \"/>");
            defSb.Append($"<rect x=\"{Math.Round(width / 2D, 0)}\" y=\"{Math.Round(height / 2D, 0)}\" width=\"1\" height=\"1\" {WriteFillColor(afc)}/>");
            defSb.Append($"</pattern>");
        }

        private static object WriteFillColor(Color c)
        {
            double opacity = -1;
            if (c.A < 255 && c != Color.Empty)
            {
                opacity = c.A / 255D * 100;
            }
            if (opacity == -1)
            {
                return $"fill=\"#{c.To6CharHexString()}\"";
            }
            else
            {
                return $"fill=\"#{c.To6CharHexString()}\" fill-opacity=\"{opacity}%\"";
            }
        }

        private static void SetPatternArray(StringBuilder defSb, string name, Color afc, Color abc, short[][] pathArray)
        {
            var height = pathArray.GetLength(0);
            var width = pathArray[0].GetLength(0);
            defSb.Append($"<pattern id=\"{name}\" width=\"{width}\" height=\"{height}\" patternUnits=\"userSpaceOnUse\">");
            defSb.Append($"<rect width=\"{width}\" height=\"{height}\" {WriteFillColor(abc)}/>");
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    if (pathArray[y][x] == 1)
                    {
                        defSb.Append($"<rect x=\"{x}\" y=\"{y}\" width=\"1\" height=\"1\" {WriteFillColor(afc)}/>");
                    }
                }
            }
            defSb.Append($"</pattern>");
        }

        private string GetScaling(RenderGradientFill gradientFill)
        {
            var sb = new StringBuilder();
            var t = gradientFill.FocusPoint.TopOffset / 100;
            var b = gradientFill.FocusPoint.BottomOffset / 100;
            var l = gradientFill.FocusPoint.LeftOffset / 100;
            var r = gradientFill.FocusPoint.RightOffset / 100;

            var tt = gradientFill.TileRectangle.TopOffset / 100;
            var tb = gradientFill.TileRectangle.BottomOffset / 100;
            var tl = gradientFill.TileRectangle.LeftOffset / 100;
            var tr = gradientFill.TileRectangle.RightOffset / 100;

            var dy = Math.Abs(t - b);
            var dx = Math.Abs(l - r);
            //var scaleTo = Math.Min(dy, dx);
            //var mult = 0.5 + (scaleTo / 2);
            var cx = Bounds.Width * l;
            var cy = Bounds.Height * t;
            //var radX = Math.Abs((t - b )) * _svgDrawing.FontSize.Item1;
            //var radY = Math.Abs((l - r)) * _svgDrawing.FontSize.Item2;
            var rx = Bounds.Width * 0.5 * (Math.Abs(t - tb) + Math.Abs(b - tt));
            var ry = Bounds.Height * 0.5 * (Math.Abs(l - tr) + Math.Abs(r - tl));

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

        private void SetStopColors(StringBuilder defSb, RenderGradientFill gradientFill, PathFillMode fillMode)
        {
            //Svg requires starting at 0 and moving towards 100% Excel sometimes starts at 100
            //Sort to get around that
            var sortedGradientColors = gradientFill.Colors.OrderBy(x => x.Position);

            foreach (var c in sortedGradientColors)
            {
                var color = ColorUtils.GetAdjustedColor(fillMode, c.Color);
                // TODO: check if ix should be increased...?
                var op = c.Opacity;
                defSb.Append($"<stop offset=\"{c.Position}%\" stop-color=\"#{color.To6CharHexString()}\" {(op == 1 ? "" : "stop-opacity=\"" + (op * 100).ToString("N0") + "%\"")} />");
            }
        }

        private string GetXy(RenderItem item, UserSpaceSettings userSpace, double? angle)
        {
            //if (userSpace == UserSpaceSettings.UserSpaceOnUse_Parent)
            //{
                double theta = MathHelper.Radians((angle ?? 90) % 360);

                var l = item.Bounds.Left;
                var t =  item.Bounds.Top;
                var w = item.Bounds.Width;
                var h = item.Bounds.Height;

                double dx = Math.Cos(theta);
                double dy = Math.Sin(theta);
                double length = w * Math.Abs(dx) + h * Math.Abs(dy);

                double cx = l + w / 2.0;
                double cy = t + h / 2.0;

                double halfX = (length / 2.0) * dx;
                double halfY = (length / 2.0) * dy;

                double x1 = cx - halfX, y1 = cy - halfY;
                double x2 = cx + halfX, y2 = cy + halfY;

                return $" x1=\"{(x1).PointToPixelString("0.00")}\" x2=\"{(x2).PointToPixelString("0.00")}\" y1=\"{y1.PointToPixelString("0.00")}\" y2=\"{y2.PointToPixelString("0.00")}\"";
            //}
            //return "";
        }

        private string GetOpacity(double opacity)
        {
            if (opacity > 0 && opacity < 1)
            {
                return $"stop-opacity=\"{opacity.ToString("0")}%\"";
            }
            return "";
        }

        private string SetStretchTileProps(RenderBlipFill blipFill)
        {
            if (blipFill.Stretch)
            {
                var x = Bounds.Width * blipFill.StretchOffset.LeftOffset / 100;
                var y = Bounds.Height * blipFill.StretchOffset.TopOffset / 100;
                var width = Bounds.Width - x - Bounds.Width * blipFill.StretchOffset.RightOffset / 100;
                var height = Bounds.Height - x - Bounds.Height * blipFill.StretchOffset.BottomOffset / 100;
                return $" preserveAspectRatio=\"none\" x=\"{x.ToString(CultureInfo.InvariantCulture)}\" y=\"{y.ToString(CultureInfo.InvariantCulture)}\" width=\"{width.PointToPixelString()}\" height=\"{height.PointToPixelString()}\" ";
            }
            else if (!(blipFill.Tile.HorizontalOffset == 0 && blipFill.Tile.VerticalOffset == 0 &&
                    blipFill.Tile.HorizontalRatio == 100 && blipFill.Tile.VerticalRatio == 100 && blipFill.Tile.FlipMode == TileFlipMode.None))
            {
                var flip = "";
                switch (blipFill.Tile.FlipMode)
                {
                    case TileFlipMode.X:
                        flip = $" transform=\"translate({Bounds.Width.ToString(CultureInfo.InvariantCulture)}, 0) scale(-1, 1)\"";
                        break;
                    case TileFlipMode.Y:
                        flip = $" transform=\"translate(0, {Bounds.Height.ToString(CultureInfo.InvariantCulture)}) scale(1, -1)\"";
                        break;
                    case TileFlipMode.XY:
                        flip = $" transform=\"translate({Bounds.Width.ToString(CultureInfo.InvariantCulture)}, {Bounds.Height.ToString(CultureInfo.InvariantCulture)}) scale(-1, -1)\"";
                        break;
                }
                return $"{flip}";
            }

            return "";
        }

        private object GetImageAsHref(RenderBlipFill blipFill)
        {
            return $"data:{blipFill.ContentType};base64," + Convert.ToBase64String(blipFill.ImageBytes);
        }
    }
}
