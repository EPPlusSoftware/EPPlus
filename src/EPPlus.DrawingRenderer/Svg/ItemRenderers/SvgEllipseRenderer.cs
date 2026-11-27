/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  23/11/2026         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Utils;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgEllipseRenderer : SvgBaseRenderer<EllipseRenderItem>
    {
        public SvgEllipseRenderer(StringBuilder outputStream) : base(outputStream)
        {
        }
        public override void Render(EllipseRenderItem item)
        {
            OutputStream.AppendFormat("<ellipse cx=\"{0}\" cy=\"{1}\" rx=\"{2}\" ry=\"{3}\" ",
                item.Cx.PointToPixelString(),
                item.Cy.PointToPixelString(),
                item.Rx.PointToPixelString(),
                item.Ry.PointToPixelString());

            RenderBase(item);

            OutputStream.AppendFormat("/>");
        }

    }
}
