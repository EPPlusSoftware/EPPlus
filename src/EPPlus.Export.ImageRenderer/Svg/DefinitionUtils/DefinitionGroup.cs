using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.DefinitionUtils
{
    internal class DefinitionGroup : RenderItemBase
    {
        internal List<RenderItem> Items;

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
