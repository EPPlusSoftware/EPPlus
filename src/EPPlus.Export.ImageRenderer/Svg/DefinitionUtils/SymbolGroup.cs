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
    internal class SymbolGroup : RenderItem
    {
        protected const string _urlRef = "url(#{0})";

        protected string _id = null;

        protected List<RenderItem> _items = new List<RenderItem>();

        internal string Mask = null;

        public SymbolGroup(DrawingBase renderer, string id) : base(renderer)
        {
            _id = id;
        }

        public override RenderItemType Type => RenderItemType.Group;

        public override void Render(StringBuilder sb)
        {
            sb.Append($"<symbol id=\"{_id}\" " +
                $"viewBox=\"0 0 100% 100%\" ");

            if(string.IsNullOrEmpty(Mask) == false)
            {
                sb.Append($"mask=\"{Mask}\" ");
            }
            sb.Append(">");

            foreach (var item in _items)
            {
                item.Render(sb);
            }

            sb.Append("</symbol>");
        }
    }
}
