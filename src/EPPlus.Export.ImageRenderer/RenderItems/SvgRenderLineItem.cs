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
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Globalization;
using System.Text;
namespace EPPlusImageRenderer.RenderItems
{
    internal class SvgRenderLineItem : SvgRenderItem
    {

        public SvgRenderLineItem(DrawingBase renderer, BoundingBox parent) : base(renderer, parent)
        {

        }
        double _x1, _y1, _x2, _y2;
        public double X1
        {
            get
            {
                return _x1;
            }
            set
            {
                _x1 = value;
                UpdateBounds();
            }
        }
        public double Y1
        {
            get
            {
                return _y1;
            }
            set
            {
                _y1 = value;
                UpdateBounds();
            }
        }
        public double X2
        {
            get
            {
                return _x2;
            }
            set
            {
                _x2 = value;
                UpdateBounds();
            }
        }
        public double Y2
        {
            get
            {
                return _y2;
            }
            set
            {
                _y2 = value;
                UpdateBounds();
            }
        }
        private void UpdateBounds()
        {
            var px = Math.Min(X1, X2);
            var py = Math.Min(Y1, Y2);
            var sizeX = Math.Abs(X2 - X1);
            var sizeY = Math.Abs(Y2 - Y1);

            Bounds.Position = new Vector2(px, py);
            Bounds.Size = new Vector2(sizeX, sizeY);
        }

        public override RenderItemType Type => RenderItemType.Line;

        public override void Render(StringBuilder sb)
        {
            //Draw transparent lines to create the compond line effect, as SVG does not support compound lines natively
            switch (CompoundLineStyle)
            {
                case eCompoundLineStyle.Double:
                    LineCap = eLineCap.Flat;
                    var name = $"double-stroke-{Guid.NewGuid().ToString()}";
                    sb.Append($"<defs><mask id=\"{name}\">");

                    RenderLineItem(sb, BorderWidth, "white", null);
                    RenderLineItem(sb, BorderWidth * (3D / 7D), "black", null);
                    sb.Append($"</mask></defs><rect width=\"100%\" height=\"100%\" fill=\"{BorderColor}\" mask=\"url(#{name})\" />");
                    break;
                case eCompoundLineStyle.DoubleThickThin:
                    WriteThickThin(sb, "double-thick-thin-stroke-{0}", (BorderWidth ?? 1D) * 1D / 7D);
                    break;
                case eCompoundLineStyle.DoubleThinThick:
                    WriteThickThin(sb, "double-thin-thick-stroke-{0}", ((BorderWidth ?? 1D) * 1D / 7D) * -1);
                    break;
                case eCompoundLineStyle.TripleThinThickThin:
                    var guid= Guid.NewGuid().ToString();
                    var gapOffset = 5 * BorderWidth.Value / 16;
                    name = $"triple-stroke-{guid}";
                    sb.Append($"<defs>");
                    sb.Append($"<filter id=\"gap-left-{guid}\" x=\"-500%\" y=\"-500%\" width=\"1100%\" height=\"1100%\" filterUnits=\"userSpaceOnUse\"><feOffset dx=\"0\" dy=\"-{gapOffset.PointToPixel().ToString(CultureInfo.InvariantCulture)}\" /></filter>"); 
                    sb.Append($"<filter id=\"gap-right-{guid}\" x=\"-500%\" y=\"-500%\" width=\"1100%\" height=\"1100%\" filterUnits=\"userSpaceOnUse\"><feOffset dx=\"0\" dy=\"{gapOffset.PointToPixel().ToString(CultureInfo.InvariantCulture)}\" /></filter>");
                    sb.Append($"<mask id=\"{name}\">");
                    RenderLineItem(sb, BorderWidth, "white", null);
                    RenderLineItem(sb, BorderWidth * (1D / 8D), "black", $"filter=\"url(#gap-left-{guid})\"");
                    RenderLineItem(sb, BorderWidth * (1D / 8D), "black", $"filter=\"url(#gap-right-{guid})\"");
                    sb.Append($"</mask></defs><rect width=\"100%\" height=\"100%\" fill=\"{BorderColor}\" mask=\"url(#{name})\" />");
                    break;
                default:
                    RenderLineItem(sb, null, null, null);
                    break;
            }
        }
        private void WriteThickThin(StringBuilder sb, string name, double gapOffset)
        {
            var guid = Guid.NewGuid().ToString();
            name = string.Format(name, guid);
            string gapFilterName = $"f-gap-shift-{guid}";
            sb.Append("<defs>");
            sb.Append($"<filter id=\"{gapFilterName}\" x=\"-50%\" y=\"-50%\" width=\"200%\" height=\"200%\" filterUnits=\"userSpaceOnUse\"><feOffset in=\"SourceGraphic\" dy=\"{gapOffset.PointToPixel().ToString(CultureInfo.InvariantCulture)}\"/></filter>");
            sb.Append($"<mask id=\"{name}\">");
            RenderLineItem(sb, BorderWidth, "white", null);
            RenderLineItem(sb, BorderWidth * (1D / 4D), "black", $"filter=\"url(#{gapFilterName})\"");
            sb.Append($"</mask></defs><rect width=\"100%\" height=\"100%\" fill=\"{BorderColor}\" mask=\"url(#{name})\" />");
        }

        internal string Suffix = "px";

        private void RenderLineItem(StringBuilder sb, double? borderWidth, string color, string filter)
        {
            if(Suffix == "%")
            {
                sb.AppendFormat("<line x1=\"{0}\" y1=\"{1}\" x2=\"{2}\" y2=\"{3}\" ",
                X1 + Suffix,
                Y1 + Suffix,
                X2 + Suffix,
                Y2 + Suffix);
            }
            else
            {
                sb.AppendFormat("<line x1=\"{0}\" y1=\"{1}\" x2=\"{2}\" y2=\"{3}\" ",
                X1.PointToPixelString(),
                Y1.PointToPixelString(),
                X2.PointToPixelString(),
                Y2.PointToPixelString());
            }
            
            RenderCompoundItems(sb, borderWidth, color, filter);
        }

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            var clone = new SvgRenderLineItem(svgDocument, svgDocument.Bounds);
            CloneBase(clone);
            return clone;
        }

        internal SvgRenderLineItem Clone(DrawingBase baseItem)
        {
            var clone = new SvgRenderLineItem(baseItem, baseItem.Bounds);
            clone.X1 = X1;
            clone.Y1 = Y1;
            clone.X2 = X2;
            clone.Y2 = Y2;
            CloneBase(clone);
            return clone;
        }

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = X1;
            it = Y1;
            ir = X2;
            ib = Y2;
        }
    }
}