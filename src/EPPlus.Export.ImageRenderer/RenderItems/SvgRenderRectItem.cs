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
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using EPPlus.Graphics;
using EPPlus.Export.ImageRenderer.Utils;
using OfficeOpenXml.Drawing.Theme;
using EPPlus.Fonts.OpenType.Utils;

namespace EPPlusImageRenderer.RenderItems
{
    internal class SvgRenderRectItem : SvgRenderItem
    {
        public SvgRenderRectItem(DrawingBase renderer, BoundingBox parent) : base(renderer, parent)
        {

        }

        public double Left { get { return Bounds.Left; } set { Bounds.Left = value; } }
        public double Top { get { return Bounds.Top; } set { Bounds.Top = value; } }
        public double Width { get { return Bounds.Width; } set { Bounds.Width = value; } }
        public double Height { get { return Bounds.Height; } set { Bounds.Height = value; } }
        public double Right { get { return Bounds.Left + Width; }  }
        public double Bottom { get { return Bounds.Top + Height; }  }
        public double GlobalLeft => Bounds.GlobalLeft;
        public double GlobalTop => Bounds.GlobalTop;
        public double GlobalRight => Bounds.GlobalLeft + Width; 
        public double GlobalBottom => Bounds.GlobalTop + Height;
        public override RenderItemType Type => RenderItemType.Rect;

        public override void Render(StringBuilder sb)
        {
            //var groupItem = new SvgGroupItem(DrawingRenderer, Bounds);
            //groupItem.Render(sb);

            RenderRect(sb);

            //groupItem.RenderEndGroup(sb);
        }

        internal void RenderRect(StringBuilder sb)
        {
            sb.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" ",
                Left.PointToPixelString(),
                Top.PointToPixelString(),
                Width.PointToPixelString(),
                Height.PointToPixelString());
            base.Render(sb);
            sb.AppendFormat("/>");
        }

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            var clone = new SvgRenderRectItem(svgDocument, svgDocument.Bounds);
            CloneBase(clone);

            clone.Left = Left * svgDocument.Bounds.Width;
            clone.Top = Top * svgDocument.Bounds.Height;
            clone.Width = svgDocument.Bounds.Width * Width;
            clone.Height = svgDocument.Bounds.Height * Height;

            return clone;
        }
        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = Left;
            it = Top;
            ir = Width;
            ib = Height;
        }
    }
}