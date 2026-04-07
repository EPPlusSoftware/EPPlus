using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.DefinitionUtils
{
    internal class MaskGroup : RenderItem
    {
        protected string _id = null;

        protected List<RenderItem> _items = new List<RenderItem>();

        public MaskGroup(DrawingBase renderer, string id) : base(renderer)
        {
            _id = id;
        }

        public override RenderItemType Type => RenderItemType.Group;

        public override void Render(StringBuilder sb)
        {
            sb.Append($"<mask id=\"{_id}\" x=\"{Bounds.Left.PointToPixelString()}\" y=\"{Bounds.Top.PointToPixelString()}%\" width=\"{100.ToString(CultureInfo.InvariantCulture)}%\" height=\"{100.ToString(CultureInfo.InvariantCulture)}%\">");

            foreach (var item in _items)
            {
                item.Render(sb);
            }

            sb.Append("</mask>");
        }
    }
}
