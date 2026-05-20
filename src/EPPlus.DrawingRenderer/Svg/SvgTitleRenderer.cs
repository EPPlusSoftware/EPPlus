using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Utils;
using System.Globalization;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgTitleRenderer : SvgBaseRenderer<TitleRenderItem>
    {
        public SvgTitleRenderer(StringBuilder outputStream) : base(outputStream)
        { 

        }
        internal string Suffix = "px";
        public override void Render(TitleRenderItem item)
        {
            OutputStream.Append($"<title>{item.Title}</title>");
        }
    }
}
