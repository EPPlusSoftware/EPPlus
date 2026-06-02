using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace EPPlus.Export.ImageRenderer.Svg.DefinitionUtils
{
    internal class PatternItem : RenderItem
    {
        string _id = null;

        protected List<RenderItem> _items = new List<RenderItem>();

        public PatternItem(DrawingBase baseRend, string id)  : base(baseRend)
        {
            _id = id;
        }

        protected double heightPercent = 100d;
        protected double widthPercent = 100d;

        public override RenderItemType Type => RenderItemType.Group;

        public override void Render(StringBuilder sb)
        {
            sb.Append($"<pattern id=\"{_id}\" width=\"{widthPercent.ToString(CultureInfo.InvariantCulture)}%\" height=\"{heightPercent.ToString(CultureInfo.InvariantCulture)}%\">");

            foreach (var item in _items)
            {
                item.Render(sb);
            }

            sb.Append("</pattern>");
        }
    }
}
