using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Utils;
using System.Globalization;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgRectRenderer : SvgBaseRenderer<RectRenderItem>
    {
        public SvgRectRenderer(StringBuilder outputStream) : base(outputStream)
        {

        }
        internal string Suffix = "px";
        public override void Render(RectRenderItem item)
        {
            if (Suffix == "%")
            {
                OutputStream.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" ",
                item.Left.ToString(CultureInfo.InvariantCulture) + Suffix,
                item.Top.ToString(CultureInfo.InvariantCulture) + Suffix,
                item.Width.ToString(CultureInfo.InvariantCulture) + Suffix,
                item.Height.ToString(CultureInfo.InvariantCulture) + Suffix);
            }
            else
            {
                OutputStream.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" ",
                item.Left.PointToPixelString(),
                item.Top.PointToPixelString(),
                item.Width.PointToPixelString(),
                item.Height.PointToPixelString());
            }
            if(item.RoundedCornerRadius>0)
            { 
                OutputStream.AppendFormat("rx=\"{0}\" ry=\"{1}\" ", item.RoundedCornerRadius.PointToPixelString(), item.RoundedCornerRadius.PointToPixelString());
            }
            RenderBase(item);
            OutputStream.AppendFormat("/>");
        }
    }
    public class SvgUseReferenceRenderer : SvgBaseRenderer<UseReferenceRenderItem>
    {
        public SvgUseReferenceRenderer(StringBuilder outputStream) : base(outputStream)
        {
        }
        internal string Suffix = "px";
        public override void Render(UseReferenceRenderItem item)
        {
            string renderStr = $"<use x=\"{item.X.PointToPixelString()}\" y=\"{item.Y.PointToPixelString()}\" href=\"{item.Href}\"/>";
            OutputStream.Append(renderStr);
        }
    }
}
