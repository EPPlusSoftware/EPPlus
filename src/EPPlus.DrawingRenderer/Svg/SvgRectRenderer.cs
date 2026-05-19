using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Utils;
using System.Globalization;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgRectRenderer : SvgBaseRenderer
    {
        public SvgRectRenderer(StringBuilder outputStream) : base(outputStream)
        {

        }
        internal string Suffix = "px";
        public override void Render(RenderItem item)
        {
            var ri = (RectRenderItem)item;
            if (Suffix == "%")
            {
                OutputStream.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" ",
                ri.Left.ToString(CultureInfo.InvariantCulture) + Suffix,
                ri.Top.ToString(CultureInfo.InvariantCulture) + Suffix,
                ri.Width.ToString(CultureInfo.InvariantCulture) + Suffix,
                ri.Height.ToString(CultureInfo.InvariantCulture) + Suffix);
            }
            else
            {
                OutputStream.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" ",
                ri.Left.PointToPixelString(),
                ri.Top.PointToPixelString(),
                ri.Width.PointToPixelString(),
                ri.Height.PointToPixelString());
            }
            RenderBase(ri);
            OutputStream.AppendFormat("/>");
        }
    }
}
