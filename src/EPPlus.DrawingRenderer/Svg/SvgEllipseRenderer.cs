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
