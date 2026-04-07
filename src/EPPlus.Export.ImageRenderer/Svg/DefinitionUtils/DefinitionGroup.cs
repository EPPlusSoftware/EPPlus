using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.DefinitionUtils
{
    internal class DefinitionGroup : RenderItem
    {
        internal List<RenderItem> Items = new List<RenderItem>();

        public DefinitionGroup(DrawingBase renderer) : base(renderer)
        {
        }

        public override RenderItemType Type => RenderItemType.Group;

        public override void Render(StringBuilder sb)
        {
            sb.Append("<defs>");

            foreach(var item in Items)
            {
                item.Render(sb);
            }

            sb.Append("</defs>");
        }
    }
}
