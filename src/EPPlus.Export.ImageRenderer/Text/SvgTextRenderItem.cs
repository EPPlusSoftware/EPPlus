using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Text
{
    internal class SvgTextRenderItem : SvgRenderItem
    {


        public override RenderItemType Type => throw new NotImplementedException();

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            throw new NotImplementedException();
        }

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            throw new NotImplementedException();
        }
        public override void Render(StringBuilder sb)
        {

        }
    }
}
