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
using System.Globalization;
using System.Text;
namespace EPPlusImageRenderer.RenderItems
{
    internal class SvgRenderEllipseItem : SvgRenderItem
    {
        public SvgRenderEllipseItem(ExcelDrawing drawing) : base(drawing)
        {

        }
        public float Cx { get; set; }
        public float Cy { get; set; }
        public float Rx { get; set; }
        public float Ry { get; set; }

        public override SvgItemType Type => SvgItemType.Rect;

        public override void Render(StringBuilder sb)
        {
            sb.AppendFormat("<ellipse cx=\"{0}\" cy=\"{1}\" rx=\"{2}\" ry=\"{3}\" ", 
                Cx.ToString(CultureInfo.InvariantCulture), 
                Cy.ToString(CultureInfo.InvariantCulture), 
                Rx.ToString(CultureInfo.InvariantCulture), 
                Ry.ToString(CultureInfo.InvariantCulture));
            
            base.Render(sb);

            sb.AppendFormat("/>");
        }

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            var clone = new SvgRenderEllipseItem(_drawing);
            clone._theme = _theme;
            CloneBase(clone);
            //if (AdjustmentPoints != null && AdjustmentPoints.Commands == null && AdjustmentPoints.AdjustmentType == AdjustmentType.AdjustToWidthHeight)
            //{
            //    var wh = Math.Min(Width, Height);
            //    clone.Left = Left * svgDocument.FontSize.Item1;
            //    clone.Top = Top * svgDocument.FontSize.Item2;
            //    clone.Width = svgDocument.FontSize.Item1 * wh;
            //    clone.Height = svgDocument.FontSize.Item2 * wh;
            //}
            //else
            //{
            //    clone.Left = Left * svgDocument.FontSize.Item1;
            //    clone.Top = Top * svgDocument.FontSize.Item2;
            //    clone.Width = svgDocument.FontSize.Item1 * Width;
            //    clone.Height = svgDocument.FontSize.Item2 * Height;
            //}
            return clone;
        }
        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = Cx - Rx;
            it = Cy - Ry;
            ir = Cx + Rx;
            ib = Cx + Rx;
        }
    }
}