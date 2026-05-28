using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using OfficeOpenXml.Interfaces.RichText;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.DrawingRenderer.RenderItems.SvgItem
{
    internal class SvgTextRunRenderItem : TextRunRenderItem
    {
        public SvgTextRunRenderItem(BoundingBox parent) : base(parent)
        {
        }

        public SvgTextRunRenderItem(BoundingBox parent, string text, int origRtIdx) : base(parent, text, origRtIdx)
        {
        }

        public SvgTextRunRenderItem(BoundingBox parent, IFontFormatBase font, string displayText) : base(parent, font, displayText)
        {
        }

        public SvgTextRunRenderItem(BoundingBox parent, string text, IFontFormatBase font, string displayText) : base(parent, text, font, displayText)
        {
        }
    }
}
