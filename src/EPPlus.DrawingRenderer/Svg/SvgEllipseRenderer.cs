using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Utils;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgEllipseRenderer : SvgBaseRenderer
    {
        public SvgEllipseRenderer(StringBuilder outputStream) : base(outputStream)
        {
        }
        public override void Render(RenderItem item)
        {
            var re = (RenderEllipseItem)item;

            OutputStream.AppendFormat("<ellipse cx=\"{0}\" cy=\"{1}\" rx=\"{2}\" ry=\"{3}\" ",
                re.Cx.PointToPixelString(),
                re.Cy.PointToPixelString(),
                re.Rx.PointToPixelString(),
                re.Ry.PointToPixelString());

            RenderBase(re);

            OutputStream.AppendFormat("/>");
        }

    }
}
