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
using EPPlusImageRenderer.Text;
using OfficeOpenXml.Drawing;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using EPPlus.Graphics;

namespace EPPlusImageRenderer.RenderItems
{
    internal class SvgRenderRectItem : SvgRenderItem
    {
        public SvgRenderRectItem() : base()
        {

        }

        public SvgRenderRectItem(ExcelDrawing drawing) : base(drawing)
        {

        }

        public double X { get { return Bounds.Left; } set { Bounds.Left = value; } }
        public double Y { get { return Bounds.Top; } set { Bounds.Top = value; } }
        public double Width { get { return Bounds.Width; } set { Bounds.Width = value; } }
        public double Height { get { return Bounds.Height; } set { Bounds.Height = value; } }

        public override RenderItemType Type => RenderItemType.Rect;

        public override void Render(StringBuilder sb)
        {
            var groupItem = new SvgGroupItem("");
            groupItem.Render(sb);

            RenderRect(sb);

            groupItem.RenderEndGroup(sb);
        }

        private void RenderRect(StringBuilder sb)
        {
            sb.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" ",
                X.ToString(CultureInfo.InvariantCulture),
                Y.ToString(CultureInfo.InvariantCulture),
                Width.ToString(CultureInfo.InvariantCulture),
                Height.ToString(CultureInfo.InvariantCulture));
            base.Render(sb);
            sb.AppendFormat("/>");
        }

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            var clone = new SvgRenderRectItem(_drawing);
            clone._theme = _theme;
            CloneBase(clone);

            clone.X = X * svgDocument.Size.Width;
            clone.Y = Y * svgDocument.Size.Height;
            clone.Width = svgDocument.Size.Width * Width;
            clone.Height = svgDocument.Size.Height * Height;

            return clone;
        }
        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = X;
            it = Y;
            ir = Width;
            ib = Height;
        }
    }
}