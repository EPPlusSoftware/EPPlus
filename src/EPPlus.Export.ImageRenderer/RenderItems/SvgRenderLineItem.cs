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
using EPPlus.Graphics;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using System.Globalization;
using System.Text;
namespace EPPlusImageRenderer.RenderItems
{
    internal class SvgRenderLineItem : SvgRenderItem
    {
        public SvgRenderLineItem(DrawingBase renderer, BoundingBox parent) : base(renderer, parent)
        {

        }
        public float X1 { get; set; }
        public float Y1 { get; set; }
        public float X2 { get; set; }
        public float Y2 { get; set; }
        public override RenderItemType Type => RenderItemType.Line;

        public override void Render(StringBuilder sb)
        {
            sb.AppendFormat("<line x1=\"{0}\" y1=\"{1}\" x2=\"{2}\" y2=\"{3}\" ", 
                X1.ToString(CultureInfo.InvariantCulture), 
                Y1.ToString(CultureInfo.InvariantCulture), 
                X2.ToString(CultureInfo.InvariantCulture), 
                Y2.ToString(CultureInfo.InvariantCulture));
            base.Render(sb);
            if(LineCap!=eLineCap.Flat)
            {
                sb.AppendFormat(" stroke-linecap=\"{0}\"", LineCap == eLineCap.Round ? "round" : "square");
            }
            sb.AppendFormat("/>");
        }

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            var clone = new SvgRenderLineItem(svgDocument, svgDocument.Bounds);
            CloneBase(clone);
            //if (adjustmentpoints != null && adjustmentpoints.commands == null && adjustmentpoints.adjustmenttype == adjustmenttype.adjusttowidthheight)
            //{
            //    var wh = math.min(width, height);
            //    clone.x = x * svgdocument.size.item1;
            //    clone.y = y * svgdocument.size.item2;
            //    clone.width = svgdocument.size.item1 * wh;
            //    clone.height = svgdocument.size.item2 * wh;
            //}
            //else
            //{
            //    clone.x = x * svgdocument.size.item1;
            //    clone.y = y * svgdocument.size.item2;
            //    clone.width = svgdocument.size.item1 * width;
            //    clone.height = svgdocument.size.item2 * height;
            //}
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