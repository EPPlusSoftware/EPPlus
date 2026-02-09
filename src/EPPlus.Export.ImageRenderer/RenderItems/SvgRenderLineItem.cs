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
            sb.AppendFormat("<line x1=\"{0}\" y1=\"{1}\" x2=\"{2}\" y2=\"{3}\" ", 
                X1.PointToPixelString(), 
                Y1.PointToPixelString(), 
                X2.PointToPixelString(), 
                Y2.PointToPixelString());
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