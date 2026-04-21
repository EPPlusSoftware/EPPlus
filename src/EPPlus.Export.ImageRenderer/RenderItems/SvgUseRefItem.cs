using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems
{
    internal class SvgUseRefItem : RenderItem
    {
        internal string Href { get; private set; } = null;

        internal double X
        {
            get
            {
                return Bounds.Left;
            }
            set
            {
                Bounds.Left = value; 
            }
        }

        internal double Y
        {
            get
            {
                return Bounds.Top;
            }
            set
            {
                Bounds.Top = value;
            }
        }

        /// <summary>
        /// Any bounds on this property are considered OFFSETS to the original object
        /// Width and Height ONLY matter if this refers to an <svg> or <symbol> element
        /// </summary>
        /// <param name="idForHref"></param>
        internal SvgUseRefItem(DrawingBase renderer, BoundingBox parent, string idForHref) : base(renderer, parent)
        {
            Href = "#" + idForHref;
        }

        public override RenderItemType Type => RenderItemType.Reference;



        public override void Render(StringBuilder sb)
        {
            string renderStr = $"<use x=\"{X.PointToPixelString()}\" y=\"{Y.PointToPixelString()}\" href=\"{Href}\"/>";
            sb.Append(renderStr);
        }
    }
}
