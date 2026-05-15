using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlusImageRenderer;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgInnerGroup : InnerGroup
    {
        public SvgInnerGroup(DrawingBase renderer) : base(renderer)
        {
        }

        public override void Render(StringBuilder sb)
        {
            sb.Append($"<g>");

            foreach (var item in _childItems)
            {
                if (item is InnerGroup)
                {
                    var subGroup = item as InnerGroup;
                    subGroup.Render(sb);
                }
                else
                {
                    item.Render(sb);
                }
            }

            sb.Append("</g>");
        }
    }
}
