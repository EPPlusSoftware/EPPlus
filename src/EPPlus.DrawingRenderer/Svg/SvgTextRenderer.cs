using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgTextRenderer : SvgBaseRenderer<ParagraphRenderItem>
    {
        public SvgTextRenderer(StringBuilder outputStream) : base(outputStream)
        {

        }
        internal string Suffix = "px";
        public override void Render(ParagraphRenderItem item)
        {
            foreach(var c in item.Runs)
            {

            }
            //if (Suffix == "%")
            //{
            //    OutputStream.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" ",
            //    ri.Left.ToString(CultureInfo.InvariantCulture) + Suffix,
            //    ri.Top.ToString(CultureInfo.InvariantCulture) + Suffix,
            //    ri.Width.ToString(CultureInfo.InvariantCulture) + Suffix,
            //    ri.Height.ToString(CultureInfo.InvariantCulture) + Suffix);
            //}
            //else
            //{
            //    OutputStream.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" ",
            //    ri.Left.PointToPixelString(),
            //    ri.Top.PointToPixelString(),
            //    ri.Width.PointToPixelString(),
            //    ri.Height.PointToPixelString());
            //}
            //RenderBase(ri);
            //OutputStream.AppendFormat("/>");
        }
    }
}
