using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgTitleItem : RenderItem
    {
        internal string Title { get; private set; }

        public SvgTitleItem(DrawingBase renderer, string titleName) : base(renderer)
        {
            Title = titleName;
        }

        public override RenderItemType Type => RenderItemType.CommentTitle;

        public override void Render(StringBuilder sb)
        {
            sb.Append($"<title>{Title}</title>");
        }
    }
}
