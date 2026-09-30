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
    public class SvgTextRunRenderItem : TextRunRenderItem
    {
        public SvgTextRunRenderItem(RenderItem parent) : base(parent)
        {
        }

        public SvgTextRunRenderItem(RenderItem parent, string text, int origRtIdx) : base(parent, text, origRtIdx)
        {
        }

        public SvgTextRunRenderItem(RenderItem parent, IFontFormatBase font, string displayText, bool renderTextNode = false) : base(parent, font, displayText)
        {
            RenderTextNode = renderTextNode;
        }

        public SvgTextRunRenderItem(RenderItem parent, string text, IFontFormatBase font, string displayText) : base(parent, text, font, displayText)
        {
        }


        /// <summary>
        /// If set to true will render its own parent Text Node
        /// Will not work properly within paragraphs
        /// </summary>
        internal bool RenderTextNode { get; private set; } = false;

        public override RenderItem Clone()
        {
            throw new NotImplementedException();
        }
    }
}
